#-----------------------------------------------------------------
# ノートPCのバッテリ残量（100-0）が閾値以下の時、作業継続を問うダイアログボックス
# を表示し、応答が無ければスリープ状態へ移行する。
#
# usage: 
#   powershell   message-sleep-battery.ps1
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



Set-Variable -Name THRESHOLD_BATTERY_CHARGE_REMAINING -Value 50 -Option Constant
# Set-Variable -Name INTERVAL_OF_BATTERY_CHARGE_VALUE -Value 10 -Option Constant
Set-Variable -Name WAIT_FOR_DIALOG_RESPONSE_SEC -Value 180 -Option Constant
Set-Variable -Name TITLE_TEXT -Value "寝落ち時スリープ移行支援Ver.0.01" -Option Constant



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



function Confirm-Popup4ContinueWorking {
  param (
    $titleText, $nSecondsToWait, $EstimatedRunTime, $EstimatedChargeRemaining
  )

  $wsobj = new-object -comobject wscript.shell
  $msgText = "バッテリー残り時間は" + $EstimatedRunTime + "分です`n作業を継続しますか？`n`n※一定時間内に「はい」が押されない場合は、スリープ状態へ移行します。"

  Return $wsobj.popup($msgText, $nSecondsToWait, $titleText, "4")
}


If($EstimatedChargeRemaining -lt $THRESHOLD_BATTERY_CHARGE_REMAINING){
  $result = Confirm-Popup4ContinueWorking ($TITLE_TEXT) ($WAIT_FOR_DIALOG_RESPONSE_SEC) ($EstimatedRunTime) ($EstimatedChargeRemaining)

  If($result -ne "6"){
    Add-Type -Assembly System.Windows.Forms;[System.Windows.Forms.Application]::SetSuspendState(‘Suspend’, $false, $false);
    # https://www.fenet.jp/infla/column/technology/powershell%E3%81%AEsleep%E3%81%A8%E3%81%AF%EF%BC%9Ftimeout%E3%81%A8%E3%81%84%E3%81%86start-sleep%E3%81%AB%E4%BC%BC%E3%81%9F%E3%82%B3%E3%83%9E%E3%83%B3%E3%83%89%E3%82%82%E7%B4%B9%E4%BB%8B%EF%BC%81/
  }
}


