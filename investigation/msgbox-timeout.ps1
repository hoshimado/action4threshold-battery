# Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned


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
function messageDialogWithTimeout ($timeout,$message)
{

    # フォームControlを作成し、サイズを指定する
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Company Name"
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = 'CenterScreen'

    # ラベルControlを作成して、任意のテキストを設定し、
    # それをフォームControlの上に貼り付け。
    $label = New-Object System.Windows.Forms.label
    $label.Text = $message
    # Locationは省略で「(0,0)」の位置から張り付ける
    $label.Size = New-Object System.Drawing.Size(200,40)
    $form.Controls.Add($label)

    # ボタンControlを作成し、（以下略）
    # ボタン押下時の戻り値に「DialogResult::OK」を指定。
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75,120)
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

$result = messageDialogWithTimeout -message “This is the message you will see on the window`nMessage on new line” -timeout 5

if([System.Windows.Forms.DialogResult]::OK -ne $result){
    Write-Output "続行OK以外"
}else{
    Write-Output "OKボタンが押された"
}
