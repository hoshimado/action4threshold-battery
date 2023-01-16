#-----------------------------------------------------------------
# ノートPCのバッテリ残量（100-0）が閾値以下の時、作業継続を問うダイアログボックス
# を表示し、応答が無ければスリープ状態へ移行する。
#
# usage: 
#   powershell   message-sleep-battery.ps1
#
# Appendix:
#   Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
#
# タスクスケジューラ利用時:
#   プログラム／スクリプト:
#     message-sleep-battery-wrapper.vbsへのフルパスを指定
#   引数の追加（オプション）:
#     message-sleep-battery.ps1へのフルパスを指定
#   ※上記のようにWrapperを指定する理由
#     powershellへの引数に「-WindowStyle Hidden -ExecutionPolicy RemoteSigned -Command "＜.ps1ファイルのパス＞"」という
#     指定をしたとしても、そのHiddenするpowershellを開始するためのpowershellウィンドウが開いてしまう、と言う仕様があり、
#     ウィンドウを開く事を抑止できないため。
#     なお、「ユーザーがログオンしているかどうかに関わらず実行する：On」だと、確かにpowershellウィンドウは表示されないが、
#     ポップアップウィンドウも表示されなくなってしまうのでNG。
#     https://yanor.net/wiki/?PowerShell/%E3%82%B9%E3%82%AF%E3%83%AA%E3%83%97%E3%83%88/%E3%82%BF%E3%82%B9%E3%82%AF%E3%82%B9%E3%82%B1%E3%82%B8%E3%83%A5%E3%83%BC%E3%83%AB%E5%AE%9F%E8%A1%8C%E6%99%82%E3%81%AB%E3%82%A6%E3%82%A3%E3%83%B3%E3%83%89%E3%82%A6%E3%82%92%E9%9A%A0%E3%81%99
#-----------------------------------------------------------------
# BOMを付けてUTF-8保存 for 日本語


# 「作業を続ける？（寝落ちしてない？）」の確認ダイアログを出し始めるバッテリー残量の閾値
Set-Variable -Name THRESHOLD_BATTERY_CHARGE_REMAINING -Value 50 -Option Constant

# （未実装）繰り返し確認ダイアログを出す場合の、バッテリー残量の間隔
# Set-Variable -Name INTERVAL_OF_BATTERY_CHARGE_VALUE -Value 10 -Option Constant

# 確認ダイアログで待機する秒数（この秒数を超えて応答が無ければ「寝落ちしている」と判断）
# と、通知への気づきの観点で、繰返し表示する回数。実際は「秒数×回数」が待機秒数、となる。
Set-Variable -Name WAIT_FOR_DIALOG_RESPONSE_SEC -Value 60 -Option Constant
Set-Variable -Name NUMBER_OF_REPEAT_DIALOG      -Value  3 -Option Constant

# スクリプトのタイトル（ダイアログのダイアログ）
Set-Variable -Name TITLE_TEXT -Value "寝落ち時スリープ移行支援Ver.0.01" -Option Constant

# 出力するログファイル名
Set-Variable -Name LOGFILE_NAME -Value "\log-test.log" -Option Constant

# ログファイルに記載するアクション種別
Set-Variable -Name ACT_COFIRM_SKIP  -Value "Confirm-Skip" -Option Constant
Set-Variable -Name ACT_CONTINUE_YES -Value "Continue-Yes" -Option Constant
Set-Variable -Name ACT_TIMEOUT      -Value "Timeout" -Option Constant




# reading battery-property value:
Get-CimInstance -ClassName Win32_Battery | Select-Object -Property DeviceID, EstimatedChargeRemaining, EstimatedRunTime | Foreach-Object {
  # $DeviceID = $_.DeviceID

  # 残バッテリ量(100-0％) 
  $EstimatedChargeRemaining = $_.EstimatedChargeRemaining

  # 残り使用時間(分)
  $EstimatedRunTime = $_.EstimatedRunTime
}
# https://powershell.one/wmi/root/cimv2/win32_battery

# ToDo:
# AC電源接続時は、スキップする
# if ($EstimatedRunTime -eq 71582788)
# {
#   'AC Power'
# }
# else
# {
#   'EstimatedRunTime = {0:n1} hours' -f ($EstimatedRunTime/60)
# }



# ToDo:
# 再問い掛け用に、現時点のバッテリー残量をファイル出力して記録する
# ＆前回のを比較して、問い掛けへ進むか否かを分岐



# https://forums.powershell.org/t/show-opup-with-timeout-and-focus/15982/3
#
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# 利用する「.NET Framwork（のアセンブリ）」を明示的にロードする。
# ※デフォルトロードされるものもあり、その場合はそのまま利用可能。
#   ただし、期待のアセンブリが「デフォルト」に入っているとは
#   限らないので、利用するものを明示的にロードするのが良い。
# 
# ref.
# 実行環境でロードされているアセンブリを確認するには「[System.AppDomain]::CurrentDomain.GetAssemblies()」を使います。
#  [System.AppDomain]::CurrentDomain.GetAssemblies() | % { $_.GetName().Name }
# https://www.vwnet.jp/windows/PowerShell/CheckAssemblis.htm
function messageDialogWithTimeout ($timeout,$message,$titleText)
{

    # フォームControlを作成し、サイズを指定する
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $titleText
    $form.Size = New-Object System.Drawing.Size(360,200)
    $form.StartPosition = 'CenterScreen'

    # ラベルControlを作成して、任意のテキストを設定し、
    # それをフォームControlの上に貼り付け。
    $label = New-Object System.Windows.Forms.label
    $label.Text = $message
    $label.Location = New-Object System.Drawing.Point(8,8)
    $label.Size = New-Object System.Drawing.Size(304,108)
    $form.Controls.Add($label)

    # ボタンControlを作成し、（以下略）
    # ボタン押下時の戻り値に「DialogResult::OK」を指定。
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(132,120)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    # タイマーを設定。
    # トリガーされたら、自身を停止後にフォームControlを閉じる。
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = $timeout * 1000
    $timer.add_tick({ # ToDo: この記法は「J#」か？　「C#」では記法が異なるのでは？
        $timer.Stop();
        $form.Close();
        return;
    })
    $timer.Start()

    # フォームControlの表示位置（z軸）を「最前面」に指定
    $form.Topmost = $true

    # フォーム終了時にタイマーを無効にするように設定
    # ToDo: Stop()の方が適切なのでは？
    $form.add_formclosed({ # ToDo: この記法は「J#」か？　「C#」では記法が異なるのでは？
        $timer.Enabled = $false;
    })

    # フォームを表示
    $result = $form.ShowDialog()

    # フォームを破棄
    $form.Dispose()

    return $result
}

# https://www.tekizai.net/entry/powershell_messagebox_1
#                    vbYesNo = 4
# Set-Variable -Name vbYes         -Value  6 -Option Constant
# Set-Variable -Name POPUP_TIMEOUT -Value -1 -Option Constant
# $wsobj = new-object -comobject wscript.shell
# $wsobj.popup() は使わなくなったが、暫しコメントに残しておく。
function Confirm-Popup4ContinueWorking {
  param (
    $titleText, $nSecondsToWait, $EstimatedRunTime, $EstimatedChargeRemaining
  )

  $msgText = "バッテリー残り時間は" + $EstimatedRunTime + "分です`n作業を継続しますか？`n`n※一定時間内に「はい」が押されない場合は、スリープ状態へ移行します。"

  $i = $NUMBER_OF_REPEAT_DIALOG -1
  do {
    $result = messageDialogWithTimeout -message $msgText -titleText $titleText -timeout $nSecondsToWait

    $i--
  } while (($i -ge 0) -and ([System.Windows.Forms.DialogResult]::OK -ne $result))
  
  Return $result
}

function output-ActionsLog {
  param (
    $EstimatedChargeRemaining, $ActionsText
  )

  # 「$NowEpochTime = Get-Date -UFormat "%s"」だとUTCで取得される且つ、
  # 小数点を含む（ナノ秒ベースだから？）ので、.NetのDateTimeOffsetを利用する方式とする
  # ref, https://qiita.com/SAITO_Keita/items/95ca9f536328fa58a912
  #      https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-date?view=powershell-7.2#notes
  $TimeTextNow = Get-Date -Format "yyyy/MM/dd HH:mm"
  $DateTimeNow = Get-Date
  $NowEpochTimeJst = ([datetimeoffset]$DateTimeNow).ToUnixTimeSeconds()
  
  $LogText = [string]$NowEpochTimeJst + "," + $TimeTextNow + "," + [string]$EstimatedChargeRemaining + "," + [string]$ActionsText
  $OutputLogPath = $PSScriptRoot + $LOGFILE_NAME

  # ファイルパスに「[]」（角括弧・ブラケット）を含む場合を考慮して
  # 与えらえたパスをリテラル文字として扱うようオプションを付けておく。
  Write-Output $LogText | Out-File -Append -LiteralPath $OutputLogPath
}

If($EstimatedChargeRemaining -lt $THRESHOLD_BATTERY_CHARGE_REMAINING){
  $result = Confirm-Popup4ContinueWorking ($TITLE_TEXT) ($WAIT_FOR_DIALOG_RESPONSE_SEC) ($EstimatedRunTime) ($EstimatedChargeRemaining)

  If($result -ne [System.Windows.Forms.DialogResult]::OK){
    output-ActionsLog ($EstimatedChargeRemaining) ($ACT_TIMEOUT) 

    Add-Type -Assembly System.Windows.Forms;[System.Windows.Forms.Application]::SetSuspendState(‘Suspend’, $false, $false);
    # https://www.fenet.jp/infla/column/technology/powershell%E3%81%AEsleep%E3%81%A8%E3%81%AF%EF%BC%9Ftimeout%E3%81%A8%E3%81%84%E3%81%86start-sleep%E3%81%AB%E4%BC%BC%E3%81%9F%E3%82%B3%E3%83%9E%E3%83%B3%E3%83%89%E3%82%82%E7%B4%B9%E4%BB%8B%EF%BC%81/
  }else{
    output-ActionsLog ($EstimatedChargeRemaining) ($ACT_CONTINUE_YES) 
  }
}else{
  output-ActionsLog ($EstimatedChargeRemaining) ($ACT_COFIRM_SKIP) 
}


