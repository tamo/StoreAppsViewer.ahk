#Requires AutoHotkey v2.0

mygui := Gui("+Resize")
mygui.SetFont("bold")
phase := mygui.AddText("w700 cBlue", "Getting apps information...")
mygui.SetFont("norm")
msg := mygui.AddText("wp y+10 h500", "This may take minutes")
mygui.Show()

appsinfo := getAppsInfo(msg)

phase.Text := "Parsing apps information..."
msg.Text := ""
appx := getAppx(msg, appsinfo.Output)
mygui.Destroy()

mygui := Gui("+Resize")
lvy := 0, lvh := 500
lv := mygui.Add(
    "ListView",
    ; disable multiple selection
    Format("-Multi x0 y{} w700 h{}", lvy, lvh - lvy),
    ["DisplayName", "Executable", "Folder", "FullName"]
)
listApps(lv, appx, &icons)
lv.ModifyCol(1, 250)
lv.ModifyCol(2, 250)
lv.ModifyCol(3, 300)

mygui.OnEvent("Size", resizeLV)
resizeLV(g, minmax, w, h) {
    global lv
    global lvy, lvh := h
    lv.Move(, , w, lvh - lvy)
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
    fullname := v.GetText(item, 4)

    cm := Menu()
    cm.Add("Copy Executable (AppID)", substClip.Bind(shellapp))
    cm.Add("Open Folder", runsimply.Bind(folder))
    cm.Add("Run Executable", runsimply.Bind(shellapp))
    cm.Add()
    cm.Add("Reload List", reload.Bind(v))
    cm.Add()
    cm.Add("Uninstall Package", uninstall.Bind(fullname, v, item))
    cm.Default := "Run Executable"
    cm.Show(x, y)

    substClip(text, *) {
        A_Clipboard := text
    }

    runsimply(text, *) {
        Run(text)
    }

    reload(lv, *) {
        global lvy, lvh
        global mygui
        global icons
        lv.Move(, lvy := 20, , lvh - lvy)
        msg := mygui.AddText("wp x0 y0", "Reloading...")
        appsinfo := getAppsInfo(msg)
        appx := getAppx(msg, appsinfo.Output)
        msg.Text := ""
        msg.Opt("+Hidden")
        lv.Delete()
        IL_Destroy(icons)
        listApps(lv, appx, &icons)
        lv.Move(, lvy := 0, , lvh - lvy)
        lv.Redraw()
    }

    uninstall(full, lv, item, *) {
        RunWait(Format('pwsh -Command "Remove-AppxPackage -Confirm {}"', full))
        kept := StdoutToVar(Format('pwsh -Command "Get-AppxPackageManifest {}"', full))
        if (kept.ExitCode == 0 and kept.Output == "") {
            lv.Delete(item)
        }
        else {
            MsgBox("Failed to uninstall.`r`nPackage exists.")
        }
    }
}

mygui.Show()
return


getAppsInfo(msg) {
    EnvSet("TERM", "dumb") ; disable colors
    info := StdoutToVar(
        "pwsh -NoProfile -Command " .
        "
        (Join ; treat as one-liner
            [console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
            ;
            Get-AppxPackage
            | Where IsFramework -eq $False
            | ForEach {
                $maa = (Get-AppxPackageManifest $_).Package.Applications.Application;
                If ($maa -And $maa.VisualElements.AppListEntry -ne 'none') {
                    $_.PackageFamilyName + ',' +
                    $_.InstallLocation + ',' +
                    $_.PackageFullName + ',' +
                    $maa.foreach{
                        $_.Id + '|' +
                        $_.VisualElements.Square44x44Logo + '|' +
                        $_.VisualElements.DisplayName + '|'
                    }
                }
            }
        )",
        ,
        "CP65001", ; UTF-8
        msg
    )
    if (info.ExitCode != 0) {
        msg.Text := info.Output
        if (MsgBox(
            "Install PowerShell?",
            "Error getting apps information",
            "YesNo"
        ) == "Yes") {
            RunWait("winget install --id Microsoft.Powershell --source winget")
            MsgBox("Restart", "Perhaps PowerShell has been installed")
            msg.Text := "Restarting"
            return getAppsInfo(msg)
        }
        else {
            ExitApp()
        }
    }
    return info
}

getAppx(msg, lines) {
    appx := Object()
    errs := ""
    for (line in StrSplit(lines, "`n", "`r ")) {
        csv := StrSplit(line, ",", " |")
        if (csv.Length == 0) {
            continue
        }
        else if (csv.Length == 4) {
            packagefamilyname := csv[1]
            app := appx.%packagefamilyname% := Object()
            app.installlocation := csv[2]
            app.packagefullname := csv[3]

            app.triplets := Array()
            triplets := StrSplit(csv[4], "|", " ")
            if (triplets.Length == 0 or Mod(triplets.Length, 3) != 0) {
                errs .= Format("parse error: [{}]`r`n", line)
            }
            Loop (triplets.Length / 3) {
                maybeRes := resolveRes(app, triplets[(A_Index - 1) * 3 + 3], &errs)
                app.triplets.Push({
                    id: triplets[(A_Index - 1) * 3 + 1],
                    icon: triplets[(A_Index - 1) * 3 + 2],
                    name: maybeRes ?? triplets[(A_Index - 1) * 3 + 3]
                })
            }
            for (i in app.triplets) {
                msg.Text := Format(
                    "[{}]=[{}!{}] ({})`r`n",
                    i.name,
                    packagefamilyname,
                    i.id,
                    i.icon
                )
            }
        }
        else {
            errs .= Format("parse error: [{}] length={}`r`n", line, csv.Length)
        }
    }
    if (errs) {
        MsgBox(errs)
    }
    return appx

    ; Translate ms-resource:[/][/]RESOURCENAME
    resolveRes(app, maybeRes, &errs) {
        if (SubStr(maybeRes, 1, 12) != "ms-resource:") {
            return maybeRes
        }
        if (result := loadIndirect(app.packagefullname, maybeRes)) {
            return result
        }
        tweakedRes := RegExReplace(
            maybeRes,
            "^ms-resource:/*",
            "ms-resource:Resources/"
        )
        if (!(result := loadIndirect(app.packagefullname, tweakedRes))) {
            errs .= Format("can't load indirect string: {}`r`n", maybeRes)
        }
        return result
    }

    loadIndirect(pkgfull, res) {
        static out := Buffer(255, 0)
        failed := DllCall(
            "Shlwapi\SHLoadIndirectString",
            "Str", "@{" . pkgfull . "?" . res . "}",
            "Ptr", out,
            "UInt", 253,
            "Ptr*", 0
        )
        return failed ? False : StrGet(out)
    }
}

listApps(lv, appx, &icons) {
    icons := IL_Create(100, 20) ; may be optimized more
    lv.SetImageList(icons)
    list := Array()
    for (pfn, app in appx.OwnProps()) {
        for (i in app.triplets) {
            icon := i.icon ? findIcon(app.installlocation . "\" . i.icon) : ""
            lv.Add(
                "Icon" . IL_Add(icons, icon, 0xFFFFFF, True),
                i.name,
                pfn . "!" . i.id,
                app.installlocation,
                app.packagefullname
            )
        }
    }
    return

    findIcon(path) {
        if (FileExist(path)) {
            return path
        }
        SplitPath(path, , &dir, &ext, &name)
        ; can contain various options like "target-*"
        Loop Files Format("{}\{}.*.{}", dir, name, ext), "F" {
            ; too lazy to find the best one
            return A_LoopFileFullPath
        }
    }
}

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
StdoutToVar(sCmd, sDir := "", sEnc := "CP0", gMsg := False) {
    ; Create 2 buffer-like objects to wrap the handles to take advantage of the __Delete meta-function.
    oHndStdoutRd := { Ptr: 0, __Delete: delete(this) => DllCall("CloseHandle", "Ptr", this) }
    oHndStdoutWr := { Base: oHndStdoutRd }

    If !DllCall("CreatePipe"
        , "PtrP", oHndStdoutRd
        , "PtrP", oHndStdoutWr
        , "Ptr", 0
        , "UInt", 0)
        Throw OSError(, , "Error creating pipe.")
    If !DllCall("SetHandleInformation"
        , "Ptr", oHndStdoutWr
        , "UInt", 1
        , "UInt", 1)
        Throw OSError(, , "Error setting handle information.")

    PI := Buffer(A_PtrSize == 4 ? 16 : 24, 0)
    SI := Buffer(A_PtrSize == 4 ? 68 : 104, 0)
    NumPut("UInt", SI.Size, SI, 0)
    NumPut("UInt", 0x100, SI, A_PtrSize == 4 ? 44 : 60)
    NumPut("Ptr", oHndStdoutWr.Ptr, SI, A_PtrSize == 4 ? 60 : 88)
    NumPut("Ptr", oHndStdoutWr.Ptr, SI, A_PtrSize == 4 ? 64 : 96)

    If !DllCall("CreateProcess"
        , "Ptr", 0
        , "Str", sCmd
        , "Ptr", 0
        , "Ptr", 0
        , "Int", True
        , "UInt", 0x08000000
        , "Ptr", 0
        , "Ptr", sDir ? StrPtr(sDir) : 0
        , "Ptr", SI
        , "Ptr", PI)
        Throw OSError(, , "Error creating process.")

    ; The write pipe must be closed before reading the stdout so we release the object.
    ; The reading pipe will be released automatically on function return.
    oHndStdOutWr := ""

    ; Before reading, we check if the pipe has been written to, so we avoid freezings.
    nAvail := 0, nLen := 0
    While DllCall("PeekNamedPipe"
        , "Ptr", oHndStdoutRd
        , "Ptr", 0
        , "UInt", 0
        , "Ptr", 0
        , "UIntP", &nAvail
        , "Ptr", 0) != 0
    {
        ; If the pipe buffer is empty, sleep and continue checking.
        If !nAvail && Sleep(100)
            Continue
        cBuf := Buffer(nAvail + 1)
        DllCall("ReadFile"
            , "Ptr", oHndStdoutRd
            , "Ptr", cBuf
            , "UInt", nAvail
            , "PtrP", &nLen
            , "Ptr", 0)
        newStr := StrGet(cBuf, nLen, sEnc)
        sOutput .= newStr
        If gMsg
            gMsg.Text := newStr
    }

    ; Get the exit code, close all process handles and return the output object.
    DllCall("GetExitCodeProcess"
        , "Ptr", NumGet(PI, 0, "Ptr")
        , "UIntP", &nExitCode := 0)
    DllCall("CloseHandle", "Ptr", NumGet(PI, 0, "Ptr"))
    DllCall("CloseHandle", "Ptr", NumGet(PI, A_PtrSize, "Ptr"))
    Return { Output: sOutput, ExitCode: nExitCode }
}