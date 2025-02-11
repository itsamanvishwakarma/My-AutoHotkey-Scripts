#Requires AutoHotkey v2.0


CapsLock::
{
    static modifierMode := false  ; Make it static so it persists between function calls
    pressStartTime := A_TickCount
    modifierSent := false
    
    ; Wait for either 2 seconds or CapsLock release
    while (GetKeyState("CapsLock", "P")) {
        if (A_TickCount - pressStartTime >= 300 && !modifierSent) {
            ; Enter modifier mode and send modifier keys
            modifierMode := true
            Send "{LWin Down}{Ctrl Down}{Alt Down}{Shift Down}"
            modifierSent := true
        }
        Sleep 10
    }
    
    ; If we entered modifier mode
    if (modifierMode) {
        ; Release all modifier keys
        Send "{LWin Up}{Ctrl Up}{Alt Up}{Shift Up}"
        modifierMode := false
    } else {
        ; If CapsLock was pressed briefly, toggle CapsLock state
        SetCapsLockState !GetKeyState("CapsLock", "T")
    }
}

; Block native CapsLock behavior to prevent double-toggling
*CapsLock::Return