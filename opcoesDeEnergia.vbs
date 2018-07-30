Dim objShell

Set objShell = CreateObject("WScript.Shell")

' configura para nunca hibernar e nunca desligar o video do monitor

' objShell.Run "cmd /c powercfg /change ""Always On"" /standby-timeout-ac 0", 0, True
' objShell.Run "cmd /c powercfg /change ""Always On"" /hibernate-timeout-ac 0", 0, True
' objShell.Run "cmd /c powercfg /setactive ""Always On""", 0, True

objShell.Run "cmd /c powercfg /change /standby-timeout-ac 0"
objShell.Run "cmd /c powercfg /change /hibernate-timeout-ac 0"

Set objShell = Nothing