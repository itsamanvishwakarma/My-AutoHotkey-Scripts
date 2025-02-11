#Requires AutoHotkey v2.0

; Various utility functions. This script doesn't run anything by itself, but is meant to be included in other scripts

; -------------------------------------------------------------------------------

; Gets all the controls of a window as objects from the Windows API, given the window's HWND
; Can be used as a replacement for WinGetControls() which only returns control names, this way you can get the names and HWNDs in one go
;    Return Type: Array of control objects with properties: Class (String), ClassNN (String), ControlID (Int), Hwnd (Int / Pointer)
;    Optional Parameter: getText (Bool) - If true, will also get the text of each control and put it into a Text property
GetAllControlsAsObjects_ViaWindowsAPI(windowHwnd, getText:=unset) {
    ; ---------------- Local Functions ----------------
    EnumChildProc(hwnd, lParam) {
        controlsArray := ObjFromPtrAddRef(lParam)
        
        ; Get control class
        classNameBuffer := Buffer(256)
        DllCall("GetClassName",
            "Ptr", hwnd,
            "Ptr", classNameBuffer,
            "Int", 256
        )
        className := StrGet(classNameBuffer) ; Convert buffer to string

        classNN := ControlGetClassNN(hwnd)
        
        ; Get control ID
        id := DllCall("GetDlgCtrlID", "Ptr", hwnd)

        controlObject := { Hwnd: hwnd, Class: className, ClassNN: classNN, ControlID: id }

        if (IsSet(getText) and getText) {
            text := ControlGetText(hwnd)
            controlObject.Text := text
        }
        
        ; Add control info to the array
        controlsArray.Push(controlObject)
        
        return true  ; Continue enumeration
    }
    ; ------------------------------------------------

    controlsArray := []

    ; Enumerate child windows
    DllCall("EnumChildWindows",
        "Ptr", windowHwnd,
        "Ptr", CallbackCreate(EnumChildProc, "F", 2),
        "Ptr", ObjPtr(controlsArray)
    )
    
    return controlsArray
}

; Checks if a window has a control with a specific class name. Allows wildcards in the form of an asterisk (*), otherwise exact matching is done
;    Return Type: Bool (True/False)
CheckWindowHasControlName(hwnd, pattern) {
    try {
        controlsObj := WinGetControls("ahk_id " hwnd)
        
        ; Wildcard pattern matching
        if InStr(pattern, "*") {
            pattern := "^" StrReplace(pattern, "*", ".*") "$"
            for ctrlName in controlsObj {
                if RegExMatch(ctrlName, pattern)
                    return true
            }
        }
        ; Exact matching
        else {
            for ctrlName in controlsObj {
                if InStr(ctrlName, pattern)
                    return true
            }
        }
    }
    return false
}

; Shows a dialog box to enter a directory path and validates it
;    Return Type: String (Path) or empty string if cancelled or invalid
ShowDirectoryPathEntryBox() {
    path := InputBox("Enter a path to navigate to", "Path", "w300 h100")
    
    ; Check if user cancelled the InputBox
    if (path.Result = "Cancel")
        return ""

    ; Trim whitespace
    trimmedPath := Trim(path.Value)
        
    ; Check if the input is empty
    if (trimmedPath = "")
        return ""

    ; Use Windows API to check if the directory exists. Also works for UNC paths
    callResult := DllCall("Shlwapi\PathIsDirectoryW", "Str", trimmedPath)
    if callResult = 0 {
        ;MsgBox("Invalid path format. Please enter a valid path.")
        return ""
    } else {
        return trimmedPath
    }
}

; Get the handle of the window under the mouse cursor
GetWindowHwndUnderMouse() {
    MouseGetPos(unset, unset, &WindowhwndOut)
    ;MsgBox("Window Hwnd: " Windowhwnd)
    return WindowhwndOut
}

; Get the class name of the control under the mouse cursor
GetControlClassUnderMouse(winTitle, ctl) {
    MouseGetPos(unset, unset, &hWnd, &classNN)
    return classNN
}

; Get the Handle/Hwnd of the specific control under the mouse cursor (not the window's handle)
GetControlUnderMouseHandleID() {
    MouseGetPos(unset, unset, &windowHandleID, &controlHandleID, 2) ; If controlHandleID is not provided, it will be an empty string
    return controlHandleID
}

; Sets the theme of menus by the process - Adapted from https://www.autohotkey.com/boards/viewtopic.php?style=19&f=82&t=133886#p588184
; Usage: Put this before creating any menus. AllowDark will folow system theme. Seems that once set, the app must restart to change it.
;SetMenuTheme("AllowDark")
SetContextMenuTheme(appMode:=0) {
    static preferredAppMode       :=  {Default:0, AllowDark:1, ForceDark:2, ForceLight:3, Max:4}
    static uxtheme                :=  dllCall("Kernel32.dll\GetModuleHandle", "Str", "uxtheme", "Ptr")

    if (uxtheme) {
        fnSetPreferredAppMode := dllCall("Kernel32.dll\GetProcAddress", "Ptr", uxtheme, "Ptr", 135, "Ptr")
        fnFlushMenuThemes := dllCall("Kernel32.dll\GetProcAddress", "Ptr", uxtheme, "Ptr", 136, "Ptr")
    } else {
        return -1
    }

    if (preferredAppMode.hasProp(appMode))
        appMode:=preferredAppMode.%appMode%

    if (fnSetPreferredAppMode && fnFlushMenuThemes) { ; Ensure the functions were found
        prev := dllCall(fnSetPreferredAppMode, "Int", appMode)
        dllCall(fnFlushMenuThemes)
        return prev
    } else {
        return -1
    }
}

; Uses Windows API SendMessage to directly send a mouse wheel movement message to a window, instead of using multiple wheel scroll events
;     > This is useful for apps that ignore the Windows scroll speed / scroll line amount settings
; Use positive multiplier for scrolling up and negative for scrolling down. The handle ID can be either for the window or a control inside it
MouseScrollMultiplied(multiplier, forceWindowHandle:=false, targetHandleID:=unset, mousePosX:=unset, mousePosY:=unset) {
    ; Gets the mouse position and handles under the mouse if not provided
    if !IsSet(targetHandleID) or !IsSet(mousePosX) or !IsSet(mousePosY) {
        MouseGetPos(&mousePosX, &mousePosY, &windowHandleID, &controlHandleID, 2) ; If controlHandleID is not provided, it will be an empty string
        if (forceWindowHandle)
            controlHandleID := ""
        ; Use the control handle if available, since most apps seem to work with that than if only the window handle is used
        targetHandleID := controlHandleID ? controlHandleID : windowHandleID
    } 

    ; 120 is the default delta for one scroll notch in Windows, regardless of mouse setting for number of lines to scroll
    delta := Round(120 * multiplier) 
    ; Construct wParam: shift delta to high-order word (left 16 bits)
    wParam := delta << 16
    ; Construct lParam: combine x and y coordinates: x goes in low word, y in high word
    lParam := (mousePosY << 16) | (mousePosX & 0xFFFF)
    ; WM_MOUSEWHEEL = 0x020A
    result := SendMessage(0x020A, wParam, lParam, unset, "ahk_id " targetHandleID)

    ; Uncomment below for debugging. Sometimes apps return a failed result even if it works, so leaving this commented out since it's not reliable
    ; ---------------------------------------------
    ; using := (IsSet(controlHandleID) && controlHandleID) ? "Control" : ((IsSet(windowHandleID) && windowHandleID) ? "Window" : "Given Parameter")
    ; resultStr := result > 0 ? "Failed (" result ")" : "Success (" result ")"
    ; ToolTip("SendMessage WM_MOUSEWHEEL returned: " resultStr "`n wParam: " wParam "`n Delta: " delta "`n" using " Handle: " targetHandleID)
    ; ---------------------------------------------

    return result
}

; Check if mouse is over a specific window and control. Allows for wildcards in the control name
; Example:          #HotIf mouseOver("ahk_exe dopus.exe", "dopus.tabctrl1")
CheckMouseOverControl(winTitle, ctl := '') {
    MouseGetPos(unset, unset, &hWnd, &classNN)
    
    ; Allow for wildcards
    if (ctl = '*') {
        Return WinExist(winTitle ' ahk_id' hWnd)
    } else if InStr(ctl, '*') {
        ctl := StrReplace(ctl, '*', '.*')
        return WinExist(winTitle ' ahk_id' hWnd) && RegExMatch(classNN, '^' ctl '$')
    } else {
        Return WinExist(winTitle ' ahk_id' hWnd) && (ctl = '' || ctl = classNN)
    }
}

; Optimized version for exact control matches:
CheckMouseOverControlExact(winTitle, ctl) {
    MouseGetPos(unset, unset, &hWnd, &classNN)
    return WinExist(winTitle " ahk_id" hWnd) && (ctl = classNN)
}

; Check if mouse is over a specific window and control (allows wildcards), with additional parameters for various properties of the control
CheckMouseOverControlAdvanced(winTitle, ctl := '', ctlMinWidth := 0) {
    ; ------ Local Functions ------
    checkWidth(ctrlToCheck, winHwnd, ctlMinWidth) {
        if ctlMinWidth = 0
            return true
        ControlGetPos(&OutX, &OutY, &OutWidth, &OutHeight, ctrlToCheck, winHwnd)
        return OutWidth > ctlMinWidth
    }
    ; ----------------------------

    MouseGetPos(unset, unset, &hWnd, &classNN)

    ; Main check for window and control
    matched := false
    if (ctl = '*') {
        if WinExist(winTitle ' ahk_id' hWnd)
            matched := true
    } else if InStr(ctl, '*') {
        ctl := StrReplace(ctl, '*', '.*')
        if WinExist(winTitle ' ahk_id' hWnd) && RegExMatch(classNN, '^' ctl '$')
            matched := true
    } else {
        if WinExist(winTitle ' ahk_id' hWnd) && (ctl = '' || ctl = classNN)
            matched := true
    }

    ; If window and control match, check further criteria
    if (matched){
        if (ctlMinWidth > 0) {
            ctrlHwnd := ControlGetHwnd(classNN, "ahk_id " hWnd) ; Get the control's handle ID
            return (checkWidth(ctrlHwnd, hWnd, ctlMinWidth) = true)
        }
    }
    ; If nothing returned true, return false
    return false
}

; Check if mouse is over a specific window by program name (even if not focused)
; Example:          #HotIf mouseOverProgram("ahk_exe notepad.exe")
CheckMouseOverProgram(programTitle) {
    MouseGetPos(unset, unset, &hWnd)
    Return WinExist(programTitle " ahk_id" hWnd)
}

; Launch any program and move it to the mouse position, with parameters for relative offset vs mouse position
; Optionally, you can provide the path to the executable to launch (which may be faster, and should be more reliable), otherwise it will use the program title
LaunchProgramAtMouse(programTitle, xOffset := 0, yOffset := 0, exePath := "", forceWinActivate := false, sizeX:=0, sizeY:=0) {
    ; Store original settings to restore later
    originalMouseMode := A_CoordModeMouse
    originalWinDelay := A_WinDelay
    
    ; Optimize window operations delay
    SetWinDelay(0)
    
    ; Set coordinate mode for consistent positioning
    CoordMode("Mouse", "Screen")
    
    ; Get mouse position once and calculate new position
    MouseGetPos(&mouseX, &mouseY)
    newX := mouseX + xOffset
    newY := mouseY + yOffset
    
    ; Launch program. If path is provided, use it. Otherwise, use the program title
    If (exePath != "")
        Run(exePath)
    Else    
        Run(programTitle)
    
    timeoutAt := A_TickCount + 3000
    
    ; Instead of using WinWait, this uses a timer to check if the window exists which is considerably faster (at least 100ms)
    CheckWindow() {
        if (A_TickCount > timeoutAt) {
            SetTimer(CheckWindow, 0)
            MsgBox("Timed out after 3 seconds")
            return
        }
        
        ; Use direct window title for faster checking
        if WinExist("ahk_exe " programTitle) {
            SetTimer(CheckWindow, 0)
            
            ; Combine move and activate into one operation if possible
            if (forceWinActivate) {
                if (sizeX > 0)
                    WinMove(newX, newY, sizeX, sizeY)
                else
                    WinMove(newX, newY)
                WinActivate()
            } else {
                if (sizeX > 0)
                    WinMove(newX, newY, sizeX, sizeY)
                else
                    WinMove(newX, newY)
            }
        }
    }
    
    ; Use a faster timer interval
    SetTimer(CheckWindow, 5)
    
    ; Restore original settings
    SetWinDelay(originalWinDelay)
    CoordMode("Mouse", originalMouseMode)
}

; Gets the raw bytes data of a specific clipboard format, given the format's name string or ID number
GetClipboardFormatRawData(formatName := "", formatIDInput := unset) {
    if IsSet(formatIDInput) {
        formatId := formatIDInput
    } else if IsSet(formatName) {
        formatId := DllCall("RegisterClipboardFormat", "Str", formatName, "UInt")
        if !formatId
            Throw("Failed to register clipboard format: " formatName)
    } else {
        Throw("Error in Getting clipboard format data: No format name or ID provided.")
    }

    ; Get all clipboard data
    clipData := ClipboardAll()
    
    if clipData.Size = 0
        return []
    
    ; Create buffer to read from
    bufferObj := Buffer(clipData.Size)
    DllCall("RtlMoveMemory", 
        "Ptr", bufferObj.Ptr, 
        "Ptr", clipData.Ptr, 
        "UPtr", clipData.Size)
    
    offset := 0
    while offset < clipData.Size {
        ; Read format type (4 bytes)
        currentFormat := NumGet(bufferObj, offset, "UInt")
        
        ; If we hit a zero format, we've reached the end
        if currentFormat = 0
            break
            
        ; Read size of this format's data block (4 bytes)
        dataSize := NumGet(bufferObj, offset + 4, "UInt")
        
        ; If this is the format we want
        if currentFormat = formatId {
            ; Create array to hold the bytes
            bytes := []
            bytes.Capacity := dataSize  ; Pre-allocate for better performance
            
            ; Extract each byte directly from the buffer
            loop dataSize {
                bytes.Push(NumGet(bufferObj, offset + 8 + A_Index - 1, "UChar"))
            }
            
            return bytes
        }
        
        ; Move to next format block
        offset += 8 + dataSize  ; Skip format (4) + size (4) + data block
    }
    
    return []  ; Format not found
}
