'-----------------------------------------------------------------
' ノートPCのバッテリ残量（100-0）が閾値以下の時、作業継続を問うダイアログボックス
' を表示し、応答が無ければスリープ状態へ移行する。
'
' usage: 
'   cscript /Nologo message-sleep-battery.vbs
'   C:\windows\system32\cscript.exe
'-----------------------------------------------------------------

Const THRESHOLD_BATTERY_CHARGE_REMAINING = 80 '[100-0]
Const INTERVAL_OF_BATTERY_CHARGE_VALUE = 10   '[99-1]
Const WAIT_FOR_DIALOG_RESPONSE_SEC = 30
Const TITLE_TEXT = "寝落ち時スリープ移行支援"


Set rows = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select * from Win32_Battery",,48)

'UnixTime(が必要なら、VBScriptでは次のように算出できる)
'Dim epochSeconds
'epochSeconds = DateDiff("s", "1970/01/01 00:00:00", Now()) - 32400


Dim EstimatedRunTime, EstimatedChargeRemaining
For Each row in rows
    '残り使用時間(分) 
    EstimatedRunTime = row.EstimatedRunTime
    '残バッテリ容量
    EstimatedChargeRemaining = row.EstimatedChargeRemaining
Next


'ToDo:
'再問い掛け用に、現時点のバッテリー残量をファイル出力して記録する
'＆前回のを比較して、問い掛けへ進むか否かを分岐


Function ConfirmPopup4ContinueWorking(nSecondsToWait, EstimatedRunTime, EstimatedChargeRemaining)
    ' https://atmarkit.itmedia.co.jp/ait/articles/0410/21/news099_3.html
    Set objWshShell = WScript.CreateObject("Wscript.Shell")
    Dim msgText
    msgText = "バッテリー残り時間は" & EstimatedRunTime & "分です" & vbCrLf & "作業を継続しますか？" & vbCrLf & vbCrLf & "※一定時間内に「はい」が押されない場合は、スリープ状態へ移行します。"

    ConfirmPopup4ContinueWorking = objWshShell.Popup(msgText, nSecondsToWait, TITLE_TEXT, vbYesNo )
End Function


Dim result
If EstimatedChargeRemaining =< THRESHOLD_BATTERY_CHARGE_REMAINING Then
    result = ConfirmPopup4ContinueWorking(WAIT_FOR_DIALOG_RESPONSE_SEC, EstimatedRunTime, EstimatedChargeRemaining)
    If result <> vbYes Then
        CreateObject("WScript.Shell").Run "rundll32 powrprof.dll, SetSuspendState"
        ' CreateObject("WScript.Shell").Run "rundll32 powrprof.dll, SetSuspendState Hibernate"
    End If
End If



