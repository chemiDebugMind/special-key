#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%
run_as_admin()

; URL to download the latest version of the main script
scriptURL := "https://raw.githubusercontent.com/chemiDebugMind/special-key/main/Special_Keys.ahk"
localScript := "Special_Keys.ahk"

; Close any running process that starts with "Special_Key"
; Using AutoHotkey's built-in commands
SetTimer, CloseSpecialKey, -1

Sleep, 1000 ; Wait a second to ensure the processes are fully closed

; Get the Windows startup folder path
EnvGet, startupFolder, APPDATA
startupFolder := startupFolder . "\Microsoft\Windows\Start Menu\Programs\Startup"

; Download the script to the startup folder
UrlDownloadToFile, %scriptURL%, %startupFolder%\%localScript%
if (ErrorLevel = 0) {    
    ; Define paths
    sourceScript := startupFolder . "\" . localScript
    outputExe := startupFolder . "\Special_Keys.exe"
    compilerPath := "C:\Program Files\AutoHotkey\Compiler\Ahk2Exe.exe"
   
    ; Compile the AHK script to EXE
    RunWait, %compilerPath% /in "%sourceScript%" /out "%outputExe%", , Hide
   
    ; Check if the EXE was created successfully
    if FileExist(outputExe)
    {
        FileDelete, %sourceScript%
        if (ErrorLevel = 0) {
            alert_message("The downloaded AHK file was deleted successfully.")
            ; Run the newly created executable
            Run, %outputExe%
            if (ErrorLevel = 0) {
                alert_message("The new executable has been started.")

            } else {
                alert_message("There was an error starting the new executable.")
                msgbox, %ErrorLevel%
            }
        } else {
            alert_message("There was an error deleting the AHK file.")

        }
    }
    else
    {
        alert_message("There was an error compiling the script.")

    }
} else {
    alert_message("Error downloading the script.")
}
ExitApp


; Function to close Special_Key processes
CloseSpecialKey:
    Process, Close, Special_Key.exe
return


alert_message(msg){
    MsgBox,262144,,%msg%, 1
}

; Run as admin
run_as_admin(){
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