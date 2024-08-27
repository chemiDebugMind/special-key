#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%


; URL to download the latest version of the main script
scriptURL := "https://raw.githubusercontent.com/chemiDebugMind/special-key/main/Special_Key.ahk"
localScript := "Special_Key.ahk"

run_as_admin()


; Close any running process that starts with "Special_Key"
ProcessList := []
for process in ComObject("WbemScripting.SWbemLocator").ConnectServer().ExecQuery("Select * from Win32_Process")
{
    if (RegExMatch(process.Name, "^Special_Key"))
    {
        ProcessList.Push(process.ProcessId)
    }
}

for index, pid in ProcessList
{
    Process, Close, %pid%
}

Sleep, 1000 ; Wait a second to ensure the processes are fully closed
; Get the Windows startup folder path
EnvGet, startupFolder, APPDATA
startupFolder := startupFolder . "\Microsoft\Windows\Start Menu\Programs\Startup"

; Download the script to the startup folder
UrlDownloadToFile, %scriptURL%, %startupFolder%\%localScript%

if (ErrorLevel = 0) {    
    ; Define paths
    sourceScript := startupFolder . "\" . localScript
    outputExe := startupFolder . "\Special_Key.exe"
    compilerPath := "C:\Program Files\AutoHotkey\Compiler\Ahk2Exe.exe"
    
    ; Compile the AHK script to EXE
    RunWait, %compilerPath% /in "%sourceScript%" /out "%outputExe%", , Hide
    
    ; Check if the EXE was created successfully
    if FileExist(outputExe)
    {
        FileDelete, %sourceScript%
        if (ErrorLevel = 0) {
            MsgBox, The downloaded AHK file was deleted successfully.
            ; Run the newly created executable
            Run, %outputExe%
            if (ErrorLevel = 0) {
                MsgBox, The new executable has been started.
            } else {
                MsgBox, There was an error starting the new executable.
            }
        } else {
            MsgBox, There was an error deleting the AHK file.
        }
    }
    else
    {
        MsgBox, There was an error compiling the script.
    }
} else {
    MsgBox, Error downloading the script. Error level: %ErrorLevel%
}


; Run as admin
run_as_admin(){
    ; Snippet taken from AHK documentation
    full_command_line := DllCall("GetCommandLine", "str")

    if not (A_IsAdmin or RegExMatch(full_command_line, " /restart(?!\S)"))
    {
        try
        {
            if A_IsCompiled
                Run *RunAs "%A_ScriptFullPath%" /restart
            else
                Run *RunAs "%A_AhkPath%" /restart "%A_ScriptFullPath%"
        }
        ExitApp
    }

}