'-----------------------------------------------------------------
' �m�[�gPC�̃o�b�e���c�ʁi100-0�j��臒l�ȉ��̎��A��ƌp����₤�_�C�A���O�{�b�N�X
' ��\�����A������������΃X���[�v��Ԃֈڍs����B
'
' usage: 
'   cscript /Nologo message-sleep-battery.vbs
'   C:\windows\system32\cscript.exe
'-----------------------------------------------------------------

Const THRESHOLD_BATTERY_CHARGE_REMAINING = 80 '[100-0]
Const INTERVAL_OF_BATTERY_CHARGE_VALUE = 10   '[99-1]
Const WAIT_FOR_DIALOG_RESPONSE_SEC = 30
Const TITLE_TEXT = "�Q�������X���[�v�ڍs�x��"


Set rows = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select * from Win32_Battery",,48)

'UnixTime(���K�v�Ȃ�AVBScript�ł͎��̂悤�ɎZ�o�ł���)
'Dim epochSeconds
'epochSeconds = DateDiff("s", "1970/01/01 00:00:00", Now()) - 32400


Dim EstimatedRunTime, EstimatedChargeRemaining
For Each row in rows
    '�c��g�p����(��) 
    EstimatedRunTime = row.EstimatedRunTime
    '�c�o�b�e���e��
    EstimatedChargeRemaining = row.EstimatedChargeRemaining
Next


'ToDo:
'�Ė₢�|���p�ɁA�����_�̃o�b�e���[�c�ʂ��t�@�C���o�͂��ċL�^����
'���O��̂��r���āA�₢�|���֐i�ނ��ۂ��𕪊�


Function ConfirmPopup4ContinueWorking(nSecondsToWait, EstimatedRunTime, EstimatedChargeRemaining)
    ' https://atmarkit.itmedia.co.jp/ait/articles/0410/21/news099_3.html
    Set objWshShell = WScript.CreateObject("Wscript.Shell")
    Dim msgText
    msgText = "�o�b�e���[�c�莞�Ԃ�" & EstimatedRunTime & "���ł�" & vbCrLf & "��Ƃ��p�����܂����H" & vbCrLf & vbCrLf & "����莞�ԓ��Ɂu�͂��v��������Ȃ��ꍇ�́A�X���[�v��Ԃֈڍs���܂��B"

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



