#Requires AutoHotkey v2.0

mygui := Gui("+Resize")
phase := mygui.AddText("w700", "Getting apps information...")
msg := mygui.AddText("wp y+10 h500", "This may take minutes")
mygui.Show()

start:
EnvSet("TERM", "dumb") ; disable colors
appsinfo := StdoutToVar(
    "pwsh -NoProfile -Command " .
    "
    (Join ; treat as one-liner
        [console]::InputEncoding = [console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
        ;
        Get-StartApps | ForEach {
            $_.AppID + ',' + $_.Name
        }
        ;
        Get-AppxPackage | ForEach {
            $maa = (Get-AppxPackageManifest $_).Package.Applications.Application
            ;
            $_.Name + ',' +
            $_.PackageFamilyName + ',' +
            $_.InstallLocation + ',' +
            $maa.VisualElements.foreach{
                '' + $_.Square44x44Logo + ' '
            } + ',' +
            $maa.foreach{
                '' + $_.id + ' '
            }
        }
    )",
    ,
    "CP65001", ; UTF-8
    msg
)
if (appsinfo.ExitCode != 0) {
    if (MsgBox(
        "Install PowerShell?",
        "Error getting apps information",
        "YesNo"
    ) == "Yes") {
        RunWait("winget install --id Microsoft.Powershell --source winget")
        MsgBox("Restart", "Perhaps PowerShell has been installed")
        Goto("start")
    }
    ExitApp()
}

phase.Text := "Parsing apps information..."
msg.Text := ""
apps := Object()
appx := Object()
errs := ""
for (line in StrSplit(appsinfo.Output, "`n", "`r ")) {
    csv := StrSplit(line, ",", " ")
    if (csv.Length == 0) {
        continue
    }
    else if (csv.Length == 2) {
        appid := csv[1]
        app := apps.%appid% := Object()
        app.name := csv[2]
        msg.Text := Format("[{}]=[{}]`r`n", appid, app.name)
    }
    else if (csv.Length == 5) {
        name := csv[1]
        app := appx.%name% := Object()
        app.packagefamilyname := csv[2]
        app.installlocation := csv[3]
        app.icons := Array()
        for (icon in StrSplit(csv[4], " ", " ")) {
            app.icons.Push(icon ? app.installlocation . "\" . icon : "")
        }
        app.ids := Array()
        for (id in StrSplit(csv[5], " ", " ")) {
            app.ids.Push(id)
        }
        msg.Text := Format(
            "[{}]=[{}{}] ({})`r`n",
            name,
            app.packagefamilyname,
            (app.ids.Has(1) && app.ids[1]) ? "!" . app.ids[1] : "",
            app.icons.Has(1) ? app.icons[1] : ""
        )
    }
    else {
        errs .= Format("parse error: [{}]`r`n", line)
    }
}

; copy properties to apps
for (name, x in appx.OwnProps()) {
    for (id in (x.ids.Length ? x.ids : [""])) {
        appid := x.packagefamilyname . (id ? "!" . id : "")
        if (apps.HasOwnProp(appid)) {
            apps.%appid%.installlocation := x.installlocation
            apps.%appid%.icon := x.icons.Has(A_Index) ? x.icons[A_Index] : ""
        }
    }
}
mygui.Destroy()

;; non-store apps
;for (appid in apps) {
;    if (!apps.%appid%.HasOwnProp("installlocation")) {
;        errs .= Format("app doesn't have installlocation: {}`r`n", i)
;    }
;}
if (errs) {
    MsgBox(errs)
}

mygui := Gui("+Resize")
lv := mygui.Add(
    "ListView",
    "r20 w700", ; initial dimensions
    ["Name", "Executable", "Folder"]
)
icons := IL_Create(50, 20) ; may be optimized more
lv.SetImageList(icons)
list := Array()
for (appid, app in apps.OwnProps()) {
    if (app.HasOwnProp("installlocation")) {
        icon := findIcon(app.icon)
        findIcon(path) {
            if (FileExist(path)) {
                return path
            }
            SplitPath(path,, &dir, &ext, &name)
            ; can contain various options like "target-*"
            Loop Files Format("{}\{}.*.{}", dir, name, ext), "F" {
                return A_LoopFileFullPath
            }
        }
        list.Push({
            icon: IL_Add(icons, icon, 0xFFFFFF, True), ; mask white
            name: app.name,
            appid: appid,
            installlocation: app.installlocation
        })
    }
}
for (i in list) {
    lv.Add("Icon" . i.icon, i.name, i.appid, i.installlocation)
}
lv.ModifyCol(1)
lv.ModifyCol(2, 250)

mygui.OnEvent("Size", resizeLV)
resizeLV(g, minmax, w, h) {
    global lv
    lv.Move(,, w-20, h-15)
}

lv.OnEvent("DoubleClick", runItem)
runItem(v, item) {
    if (item == 0) {
        return
    }
    shellapp := Format('explorer.exe "shell:Appsfolder\{}"', v.GetText(item, 2))
    Run(shellapp)
}

lv.OnEvent("ContextMenu", showContext)
showContext(v, item, isRightClick, x, y) {
    if (item == 0) {
        return
    }
    shellapp := Format('explorer.exe "shell:Appsfolder\{}"', v.GetText(item, 2))
    folder := Format('explorer.exe "{}"', v.GetText(item, 3))

    cm := Menu()
    cm.Add("Copy Executable (AppID)", substClip.Bind(shellapp))
    cm.Add("Open Folder", runsimply.Bind(folder))
    cm.Add("Run", runsimply.Bind(shellapp))
    cm.Show(x, y)

    substClip(text, *) {
        A_Clipboard := text
    }

    runsimply(text, *) {
        Run(text)
    }
}

mygui.Show()
return


; ----------------------------------------------------------------------------------------------------------------------
; Function .....: StdoutToVar
; Description ..: Runs a command line program and returns its output.
; Parameters ...: sCmd - Commandline to be executed.
; ..............: sDir - Working directory.
; ..............: sEnc - Encoding used by the target process. Look at StrGet() for possible values.
; ..............: gMsg - Gui control with Text.
; Return .......: Command output as a string on success, empty string on error.
; AHK Version ..: AHK v2 x32/64 Unicode
; Author .......: Sean (http://goo.gl/o3VCO8), modified by nfl and by Cyruz
; License ......: WTFPL - http://www.wtfpl.net/txt/copying/
; Changelog ....: Feb. 20, 2007 - Sean version.
; ..............: Sep. 21, 2011 - nfl version.
; ..............: Nov. 27, 2013 - Cyruz version (code refactored and exit code).
; ..............: Mar. 09, 2014 - Removed input, doesn't seem reliable. Some code improvements.
; ..............: Mar. 16, 2014 - Added encoding parameter as pointed out by lexikos.
; ..............: Jun. 02, 2014 - Corrected exit code error.
; ..............: Nov. 02, 2016 - Fixed blocking behavior due to ReadFile thanks to PeekNamedPipe.
; ..............: Apr. 13, 2021 - Code restyling. Fixed deprecated DllCall types.
; ..............: Oct. 06, 2022 - AHK v2 version. Throw exceptions on failure.
; ..............: Oct. 08, 2022 - Exceptions management and handles closure fix. Thanks to lexikos and iseahound.
; ..............: Added gMsg.
; ----------------------------------------------------------------------------------------------------------------------
StdoutToVar(sCmd, sDir:="", sEnc:="CP0", gMsg:=False) {
    ; Create 2 buffer-like objects to wrap the handles to take advantage of the __Delete meta-function.
    oHndStdoutRd := { Ptr: 0, __Delete: delete(this) => DllCall("CloseHandle", "Ptr", this) }
    oHndStdoutWr := { Base: oHndStdoutRd }
    
    If !DllCall( "CreatePipe"
               , "PtrP" , oHndStdoutRd
               , "PtrP" , oHndStdoutWr
               , "Ptr"  , 0
               , "UInt" , 0 )
        Throw OSError(,, "Error creating pipe.")
    If !DllCall( "SetHandleInformation"
               , "Ptr"  , oHndStdoutWr
               , "UInt" , 1
               , "UInt" , 1 )
        Throw OSError(,, "Error setting handle information.")

    PI := Buffer(A_PtrSize == 4 ? 16 : 24,  0)
    SI := Buffer(A_PtrSize == 4 ? 68 : 104, 0)
    NumPut( "UInt", SI.Size,          SI,  0 )
    NumPut( "UInt", 0x100,            SI, A_PtrSize == 4 ? 44 : 60 )
    NumPut( "Ptr",  oHndStdoutWr.Ptr, SI, A_PtrSize == 4 ? 60 : 88 )
    NumPut( "Ptr",  oHndStdoutWr.Ptr, SI, A_PtrSize == 4 ? 64 : 96 )

    If !DllCall( "CreateProcess"
               , "Ptr"  , 0
               , "Str"  , sCmd
               , "Ptr"  , 0
               , "Ptr"  , 0
               , "Int"  , True
               , "UInt" , 0x08000000
               , "Ptr"  , 0
               , "Ptr"  , sDir ? StrPtr(sDir) : 0
               , "Ptr"  , SI
               , "Ptr"  , PI )
        Throw OSError(,, "Error creating process.")

    ; The write pipe must be closed before reading the stdout so we release the object.
    ; The reading pipe will be released automatically on function return.
    oHndStdOutWr := ""

    ; Before reading, we check if the pipe has been written to, so we avoid freezings.
    nAvail := 0, nLen := 0
    While DllCall( "PeekNamedPipe"
                 , "Ptr"   , oHndStdoutRd
                 , "Ptr"   , 0
                 , "UInt"  , 0
                 , "Ptr"   , 0
                 , "UIntP" , &nAvail
                 , "Ptr"   , 0 ) != 0
    {
        ; If the pipe buffer is empty, sleep and continue checking.
        If !nAvail && Sleep(100)
            Continue
        cBuf := Buffer(nAvail+1)
        DllCall( "ReadFile"
               , "Ptr"  , oHndStdoutRd
               , "Ptr"  , cBuf
               , "UInt" , nAvail
               , "PtrP" , &nLen
               , "Ptr"  , 0 )
        newStr := StrGet(cBuf, nLen, sEnc)
        sOutput .= newStr
        If gMsg
            gMsg.Text := newStr
    }
    
    ; Get the exit code, close all process handles and return the output object.
    DllCall( "GetExitCodeProcess"
           , "Ptr"   , NumGet(PI, 0, "Ptr")
           , "UIntP" , &nExitCode:=0 )
    DllCall( "CloseHandle", "Ptr", NumGet(PI, 0, "Ptr") )
    DllCall( "CloseHandle", "Ptr", NumGet(PI, A_PtrSize, "Ptr") )
    Return { Output: sOutput, ExitCode: nExitCode } 
}
