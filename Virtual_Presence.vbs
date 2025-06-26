Set wshShell = WScript.CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")

StartTime = Timer 'set exe time

Dim max,min,rand
max=130000
min=100000

Do While (Timer - StartTime) < (9*60*60) ' 9h=32400s

    Set colProcessList = objWMIService.ExecQuery("select * from Win32_Process where Name='LogonUI.exe'") 'locked var

    If colProcessList.count = 0 Then 'check if not locked
        Randomize
        rand = Int((max-min+1)*Rnd+min)
    
        WScript.Sleep rand '(2*60*1000/500) ' 2m=120000ms
        wshShell.SendKeys "{NUMLOCK}"

        WScript.Sleep (0.1*1000) ' 1s=1000ms
        wshShell.SendKeys "{NUMLOCK}"

        'WScript.Echo rand
    Else
        'if locked do nothing
    End If

Loop
