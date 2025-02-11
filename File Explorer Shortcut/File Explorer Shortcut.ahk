#Requires AutoHotkey v2.0

; Define common paths
user_path := "C:\Users\itsam"
downloads_path := "C:\Users\itsam\Downloads"
documents_path := "C:\Users\itsam\Documents"
code_path := "C:\Users\itsam\Code"
recycle_bin_path := "shell:RecycleBinFolder"

; Only activate when File Explorer is active
#HotIf WinActive("ahk_exe explorer.exe")

; User folder (Ctrl + Alt + U)
^!u::
{
    Send("!d")  ; Alt + D to focus the address bar
    Sleep(50)  ; Small delay to ensure address bar is focused
    Send("{Text}" user_path)  ; Using {Text} mode for reliable input
    Send("{Enter}")
}

; Downloads folder (Ctrl + Alt +J)
^!j::
{
    Send("!d")
    Sleep(50)  ; Small delay to ensure address bar is focused
    Send("{Text}" downloads_path)  ; Using {Text} mode for reliable input
    Send("{Enter}")
}

; Documents folder (Ctrl+ Alt + D)
^!d::
{
    Send("!d")
    Sleep(50)
    Send("{Text}" documents_path)
    Send("{Enter}")
}

; Code folder (Ctrl + Alt + C)
^!c::
{
    Send("!d")
    Sleep(50)
    Send("{Text}" code_path)
    Send("{Enter}")
}

; Recycle Bin (Ctrl + Alt + B)
^!b::
{
    Send("!d")
    Sleep(50)
    Send("{Text}" recycle_bin_path)
    Send("{Enter}")
}

; For all other windows, let the shortcuts perform their default actions
#HotIf
^!j::Send("^!j")
^!d::Send("^!d")
^!c::Send("^!c")
^!b::Send("^!b")