Set objWshShell = WScript.CreateObject("Wscript.Shell")
objWshShell.run "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -WindowStyle Hidden -ExecutionPolicy RemoteSigned -Command " & Wscript.Arguments(0), vbHide

' https://qiita.com/trumpet_developer/items/cf7b8cb0981bdcab6c20
' https://qiita.com/tsukamoto/items/59e34fe55cbae5ca9070
' https://4thsight.xyz/37398
