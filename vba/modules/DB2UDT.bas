Attribute VB_Name = "DB2UDT"
Option Explicit

'==============================================================================
' Beschreibung   : DB to UDT (type)
'
' Author         : SHartmann
' Version        : V2.0
' Zuletzt ge�ndert: 2026-02-11
'==============================================================================

' ============================================================
'  TIA DB Strukturen in einzelne UDTs (type) wandeln und exportieren
' ============================================================

Public Const DB2UDT_VERSION As String = "v20"

' ----------------------------
' Public entry points
' ----------------------------

Public Sub Generator_Einrichten()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsHome As Worksheet, wsSrc As Worksheet, wsSet As Worksheet, wsRep As Worksheet

    Set wsHome = GetOrCreateSheet(wb, "Start")
    Set wsSrc = GetOrCreateSheet(wb, "DB_Quelle")
    Set wsSet = GetOrCreateSheet(wb, "Einstellungen")
    Set wsRep = GetOrCreateSheet(wb, "Bericht")

    ' Basic layout (no merged cells)
    FormatStart wsHome
    FormatDBQuelle wsSrc
    FormatEinstellungen wsSet
    FormatBericht wsRep

    MsgBox "DB2UDT Generator " & DB2UDT_VERSION & " ist bereit." & vbCrLf & _
           "1) DB_Quelle_Importieren (oder Text in DB_Quelle!A2 einfuegen)" & vbCrLf & _
           "2) UDTs_Erzeugen", vbInformation
End Sub

Public Sub DB_Quelle_Importieren()
    Dim fp As String
    fp = PickFile("DB-Quelldatei auswaehlen", "DB/SCL/TXT (*.db;*.scl;*.txt),*.db;*.scl;*.txt,Alle Dateien (*.*),*.*")
    If fp = "" Then Exit Sub

    Dim txt As String
    txt = ReadTextFileSmart(fp)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, "DB_Quelle")

    ' Unmerge anything to avoid ClearContents errors
    ws.Cells.UnMerge

    ' Clear only used area in column A
    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    If lastRow < 2 Then lastRow = 2

    ws.Range("A2:A" & lastRow).ClearContents

    ' Write text line-by-line into column A (starting A2)
    Dim lines() As String
    lines = Split(NormalizeNewlines(txt), vbCrLf)

    Dim n As Long: n = UBound(lines) - LBound(lines) + 1
    If n <= 0 Then Exit Sub

    Dim maxRows As Long: maxRows = 1048576 - 1 ' rows from A2..
    If n > maxRows Then n = maxRows

    Dim arr() As Variant
    ReDim arr(1 To n, 1 To 1)

    Dim i As Long
    For i = 1 To n
        arr(i, 1) = lines(LBound(lines) + i - 1)
    Next i

    ws.Range("A2").Resize(n, 1).Value = arr

    MsgBox "Importiert: " & CStr(n) & " Zeilen nach DB_Quelle.", vbInformation
End Sub

Public Sub UDTs_Erzeugen()
    Dim wsSrc As Worksheet, wsSet As Worksheet, wsRep As Worksheet
    Set wsSrc = GetOrCreateSheet(ThisWorkbook, "DB_Quelle")
    Set wsSet = GetOrCreateSheet(ThisWorkbook, "Einstellungen")
    Set wsRep = GetOrCreateSheet(ThisWorkbook, "Bericht")

    Dim prefix As String, published As String, ver As String, outFolder As String
    Dim dryRun As Boolean, deltaMode As Boolean, writeImportOrder As Boolean

    prefix = CStr(wsSet.Range("B1").Value)
    If Trim$(prefix) = "" Then prefix = "type_"

    published = UCase$(Trim$(CStr(wsSet.Range("B2").Value)))
    If published = "" Then published = "TRUE"
    If published <> "TRUE" And published <> "FALSE" Then published = "TRUE"

    ver = Trim$(CStr(wsSet.Range("B3").Value))
    If ver = "" Then ver = "0.1"

    outFolder = Trim$(CStr(wsSet.Range("B4").Value))
    outFolder = Replace(outFolder, Chr$(34), "") ' remove quotes if any

    dryRun = StrToBool(wsSet.Range("B5").Value, False)
    deltaMode = StrToBool(wsSet.Range("B6").Value, True)
    writeImportOrder = StrToBool(wsSet.Range("B7").Value, True)

    If outFolder = "" Then
        outFolder = PickFolder("Speicherort fuer UDTs waehlen")
        outFolder = Replace(outFolder, Chr$(34), "")
        If outFolder = "" Then Exit Sub
    End If

    outFolder = EnsureTrailingBackslash(outFolder)
    If Not EnsureFolderExists(outFolder) Then
        MsgBox "Ausgabeordner nicht gefunden und konnte nicht erstellt werden:" & vbCrLf & outFolder, vbCritical
        Exit Sub
    End If

    Dim srcText As String
    srcText = GetSourceTextFromSheet(wsSrc)
    If Trim$(srcText) = "" Then
        MsgBox "DB_Quelle ist leer. Fuege den Quelltext ab A2 ein oder nutze DB_Quelle_Importieren.", vbExclamation
        Exit Sub
    End If

    Dim dbName As String
    dbName = ParseDBName(srcText)
    If dbName = "" Then dbName = "DB"

    Dim decl As String
    decl = ExtractDeclarationPart(srcText)
    If Trim$(decl) = "" Then
        MsgBox "Kein Deklarationsteil gefunden (vor BEGIN).", vbExclamation
        Exit Sub
    End If

    Dim structs As Object
    Set structs = CreateObject("Scripting.Dictionary") ' name -> Collection(member lines)

    Dim rootMembers As Collection
    Set rootMembers = New Collection

    ParseStructs decl, prefix, structs, rootMembers

    If structs.Count = 0 Then
        MsgBox "Keine STRUCT-Definitionen im Deklarationsteil gefunden (vor BEGIN).", vbExclamation
        Exit Sub
    End If

    ' Report init
    ReportInit wsRep, outFolder, prefix, ver, published, dryRun, deltaMode

    Dim written As Long, skipped As Long, unchanged As Long, errors As Long
    written = 0: skipped = 0: unchanged = 0: errors = 0

    ' Write named structs
    Dim key As Variant
    For Each key In structs.keys
        Dim act As String, chg As Boolean, exists As Boolean
        If WriteUDTFile(outFolder, prefix, published, ver, CStr(key), structs(key), dryRun, deltaMode, act, exists, chg) Then
            written = written + 1
        Else
            ' not written either because dry-run or skipped
            If InStr(1, act, "SKIP", vbTextCompare) > 0 Then
                skipped = skipped + 1
                If Not chg And exists Then unchanged = unchanged + 1
            End If
        End If
        ReportAdd wsRep, prefix & CStr(key), outFolder & prefix & CStr(key) & ".udt", act, exists, chg, structs(key).Count
    Next key

    ' Write root (DB) UDT
    Dim actR As String, chgR As Boolean, exR As Boolean
    Call WriteUDTFile(outFolder, prefix, published, ver, dbName, rootMembers, dryRun, deltaMode, actR, exR, chgR)
    If InStr(1, actR, "WRITE", vbTextCompare) > 0 Then written = written + 1
    If InStr(1, actR, "SKIP", vbTextCompare) > 0 Then
        skipped = skipped + 1
        If Not chgR And exR Then unchanged = unchanged + 1
    End If
    ReportAdd wsRep, prefix & dbName, outFolder & prefix & dbName & ".udt", actR, exR, chgR, rootMembers.Count

    ' Import order
    If writeImportOrder Then
        Dim ioAct As String, ioChg As Boolean, ioEx As Boolean
        Dim ioText As String
        ioText = BuildImportOrderText(prefix, structs.keys, dbName)
        Call WriteTextMaybe(outFolder & "IMPORT_ORDER.txt", ioText, dryRun, deltaMode, ioAct, ioEx, ioChg)
        ReportAdd wsRep, "IMPORT_ORDER", outFolder & "IMPORT_ORDER.txt", ioAct, ioEx, ioChg, CountLines(ioText)
    End If

    wsRep.Columns("A:G").AutoFit

    Dim total As Long
    total = structs.Count + 1 + IIf(writeImportOrder, 1, 0)

    Dim msg As String
    msg = "Fertig (" & DB2UDT_VERSION & ")." & vbCrLf & _
          "Gesamt: " & CStr(total) & vbCrLf & _
          "Geschrieben: " & CStr(written) & vbCrLf & _
          "Uebersprungen: " & CStr(skipped) & vbCrLf & _
          "Unveraendert: " & CStr(unchanged) & vbCrLf & _
          "Ausgabe: " & outFolder

    If dryRun Then msg = msg & vbCrLf & vbCrLf & "TESTLAUF ist AN: es wurden keine Dateien geschrieben."

    MsgBox msg, vbInformation
End Sub

' Convenience macros
Public Sub UDTs_Testlauf()
    Dim wsSet As Worksheet: Set wsSet = GetOrCreateSheet(ThisWorkbook, "Einstellungen")
    Dim prev As Variant: prev = wsSet.Range("B5").Value
    wsSet.Range("B5").Value = "TRUE"
    UDTs_Erzeugen
    wsSet.Range("B5").Value = prev
End Sub

' ----------------------------
' Parsing
' ----------------------------

Private Sub ParseStructs(ByVal decl As String, ByVal prefix As String, ByVal structs As Object, ByVal rootMembers As Collection)
    Dim lines() As String
    lines = Split(NormalizeNewlines(decl), vbCrLf)

    Dim stackNames() As String
    Dim sp As Long: sp = 0

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim raw As String: raw = lines(i)
        Dim t As String: t = Trim$(raw)
        If t = "" Then GoTo NextLine

        ' End of a struct block
        If StartsWithCI(t, "END_STRUCT") Then
            If sp > 0 Then sp = sp - 1
            GoTo NextLine
        End If

        ' Root STRUCT (anonymous) start
        If StartsWithCI(t, "STRUCT") And sp = 0 Then
            sp = sp + 1
            ReDim Preserve stackNames(1 To sp)
            stackNames(sp) = "__root__"
            GoTo NextLine
        End If

        ' Named STRUCT start:  Name {..} : Struct   // comment
        Dim sName As String, sCmt As String
        If IsNamedStructStart(raw, sName, sCmt) Then
            If sp = 0 Then
                ' No explicit root found; create implicit root
                sp = 1
                ReDim stackNames(1 To sp)
                stackNames(sp) = "__root__"
            End If

            ' Add a member to parent referencing the new type
            Dim parent As String: parent = stackNames(sp)
            AddMemberLine parent, prefix, sName, sCmt, structs, rootMembers

            ' Push new struct context
            sp = sp + 1
            ReDim Preserve stackNames(1 To sp)
            stackNames(sp) = sName

            If Not structs.exists(sName) Then
                Dim col As Collection: Set col = New Collection
                structs.Add sName, col
            End If
            GoTo NextLine
        End If

        ' Normal member line inside current struct
        If sp > 0 Then
            Dim base As String, cmt As String
            SplitLineComment raw, base, cmt
            base = StripBraces(base)
            base = CollapseSpaces(Trim$(base))

            If base <> "" Then
                If Not StartsWithCI(base, "STRUCT") And Not StartsWithCI(base, "END_STRUCT") Then
                    AddRawMember stackNames(sp), base, cmt, structs, rootMembers
                End If
            End If
        End If

NextLine:
    Next i
End Sub

Private Sub AddMemberLine(ByVal ctx As String, ByVal prefix As String, ByVal name As String, ByVal cmt As String, _
                          ByVal structs As Object, ByVal rootMembers As Collection)
    Dim line As String
    line = name & " : " & Chr$(34) & prefix & name & Chr$(34) & ";"
    line = CollapseSpaces(line)
    AddRawMember ctx, line, cmt, structs, rootMembers
End Sub

Private Sub AddRawMember(ByVal ctx As String, ByVal base As String, ByVal cmt As String, _
                         ByVal structs As Object, ByVal rootMembers As Collection)
    base = CollapseSpaces(Trim$(base))
    If base = "" Then Exit Sub

    Dim outLine As String
    outLine = "      " & base
    If Trim$(cmt) <> "" Then
        outLine = outLine & "   " & Trim$(cmt)
    End If

    If ctx = "__root__" Then
        rootMembers.Add outLine
    Else
        If Not structs.exists(ctx) Then
            Dim col As Collection: Set col = New Collection
            structs.Add ctx, col
        End If
        structs(ctx).Add outLine
    End If
End Sub

Private Function IsNamedStructStart(ByVal line As String, ByRef nameOut As String, ByRef commentOut As String) As Boolean
    Dim base As String, cmt As String
    SplitLineComment line, base, cmt

    base = StripBraces(base)
    base = CollapseSpaces(Trim$(base))

    ' Expect: <Name> : Struct
    Dim p As Long: p = InStr(1, base, ":", vbTextCompare)
    If p <= 0 Then Exit Function

    Dim leftPart As String: leftPart = Trim$(Left$(base, p - 1))
    Dim rightPart As String: rightPart = Trim$(Mid$(base, p + 1))

    If Not StartsWithCI(rightPart, "Struct") Then Exit Function
    If Not IsValidIdentifier(leftPart) Then Exit Function

    nameOut = leftPart
    commentOut = cmt
    IsNamedStructStart = True
End Function

Private Function ExtractDeclarationPart(ByVal src As String) As String
    Dim norm As String: norm = NormalizeNewlines(src)
    Dim lines() As String: lines = Split(norm, vbCrLf)

    Dim out As String: out = ""
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim t As String: t = Trim$(lines(i))
        If UCase$(t) = "BEGIN" Then Exit For
        out = out & lines(i) & vbCrLf
    Next i
    ExtractDeclarationPart = out
End Function

Private Function ParseDBName(ByVal src As String) As String
    Dim norm As String: norm = NormalizeNewlines(src)
    Dim lines() As String: lines = Split(norm, vbCrLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim t As String: t = Trim$(lines(i))
        If StartsWithCI(t, "DATA_BLOCK") Then
            Dim q1 As Long: q1 = InStr(1, t, Chr$(34), vbBinaryCompare)
            If q1 > 0 Then
                Dim q2 As Long: q2 = InStr(q1 + 1, t, Chr$(34), vbBinaryCompare)
                If q2 > q1 Then
                    ParseDBName = Mid$(t, q1 + 1, q2 - q1 - 1)
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' ----------------------------
' Output (delta + dryrun)
' ----------------------------

Private Function WriteUDTFile(ByVal folder As String, ByVal prefix As String, ByVal published As String, ByVal ver As String, _
                              ByVal structName As String, ByVal members As Collection, _
                              ByVal dryRun As Boolean, ByVal deltaMode As Boolean, _
                              ByRef actionOut As String, ByRef existedOut As Boolean, ByRef changedOut As Boolean) As Boolean
    folder = EnsureTrailingBackslash(folder)

    Dim fn As String
    fn = folder & prefix & structName & ".udt"

    Dim sb As String
    sb = BuildUDTText(prefix, published, ver, structName, members)

    WriteUDTFile = WriteTextMaybe(fn, sb, dryRun, deltaMode, actionOut, existedOut, changedOut)
End Function

Private Function BuildUDTText(ByVal prefix As String, ByVal published As String, ByVal ver As String, _
                              ByVal structName As String, ByVal members As Collection) As String
    Dim sb As String
    sb = "TYPE " & Chr$(34) & prefix & structName & Chr$(34) & vbCrLf
    sb = sb & "{ Published := '" & published & "' }" & vbCrLf
    sb = sb & "VERSION : " & ver & vbCrLf
    sb = sb & "   STRUCT" & vbCrLf

    Dim i As Long
    For i = 1 To members.Count
        sb = sb & CStr(members(i)) & vbCrLf
    Next i

    sb = sb & "   END_STRUCT;" & vbCrLf & vbCrLf
    sb = sb & "END_TYPE" & vbCrLf

    BuildUDTText = sb
End Function

Private Function BuildImportOrderText(ByVal prefix As String, ByVal keys As Variant, ByVal rootName As String) As String
    Dim sb As String: sb = ""
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        sb = sb & prefix & CStr(keys(i)) & ".udt" & vbCrLf
    Next i
    sb = sb & prefix & rootName & ".udt" & vbCrLf
    BuildImportOrderText = sb
End Function

Private Function WriteTextMaybe(ByVal path As String, ByVal text As String, ByVal dryRun As Boolean, ByVal deltaMode As Boolean, _
                               ByRef actionOut As String, ByRef existedOut As Boolean, ByRef changedOut As Boolean) As Boolean
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    existedOut = fso.FileExists(path)

    Dim oldText As String: oldText = ""
    If existedOut Then
        On Error Resume Next
        oldText = ReadTextFileSmart(path)
        On Error GoTo 0
    End If

    Dim same As Boolean
    same = False
    If existedOut Then
        same = (CanonicalText(oldText) = CanonicalText(text))
    End If

    changedOut = Not same

    If deltaMode And existedOut And same Then
        actionOut = "UEBERSPRINGEN (unveraendert)"
        WriteTextMaybe = False
        Exit Function
    End If

    If dryRun Then
        If existedOut Then
            actionOut = "TESTLAUF (wuerde ueberschreiben)"
        Else
            actionOut = "TESTLAUF (wuerde erstellen)"
        End If
        WriteTextMaybe = False
        Exit Function
    End If

    ' Ensure folder exists
    Dim folder As String
    folder = fso.GetParentFolderName(path)
    If folder <> "" Then EnsureFolderExists EnsureTrailingBackslash(folder)

    WriteTextFileUTF8 path, text
    If existedOut Then
        actionOut = "SCHREIBEN (ueberschreiben)"
    Else
        actionOut = "SCHREIBEN (neu)"
    End If
    WriteTextMaybe = True
End Function

' ----------------------------
' Report sheet
' ----------------------------

Private Sub ReportInit(ByVal ws As Worksheet, ByVal outFolder As String, ByVal prefix As String, ByVal ver As String, _
                       ByVal published As String, ByVal dryRun As Boolean, ByVal deltaMode As Boolean)
    ws.Cells.UnMerge
    ws.Cells.ClearContents

    ws.Range("A1").Value = "DB2UDT Bericht " & DB2UDT_VERSION
    ws.Range("A2").Value = "Ausgabe"
    ws.Range("B2").Value = outFolder
    ws.Range("A3").Value = "Praefix"
    ws.Range("B3").Value = prefix
    ws.Range("A4").Value = "UDT Version"
    ws.Range("B4").Value = ver
    ws.Range("A5").Value = "Veroeffentlicht"
    ws.Range("B5").Value = published
    ws.Range("A6").Value = "Testlauf"
    ws.Range("B6").Value = IIf(dryRun, "TRUE", "FALSE")
    ws.Range("A7").Value = "Delta-Modus"
    ws.Range("B7").Value = IIf(deltaMode, "TRUE", "FALSE")

    ws.Range("A9").Value = "Name"
    ws.Range("B9").Value = "Datei"
    ws.Range("C9").Value = "Aktion"
    ws.Range("D9").Value = "Existiert"
    ws.Range("E9").Value = "Geaendert"
    ws.Range("F9").Value = "Zeilen"
    ws.Range("G9").Value = "Zeit"

    ws.Range("A9:G9").Font.Bold = True
End Sub

Private Sub ReportAdd(ByVal ws As Worksheet, ByVal name As String, ByVal filePath As String, ByVal action As String, _
                      ByVal existed As Boolean, ByVal changed As Boolean, ByVal lines As Long)
    Dim r As Long
    r = LastUsedRow(ws) + 1
    If r < 10 Then r = 10

    ws.Cells(r, 1).Value = name
    ws.Cells(r, 2).Value = filePath
    ws.Cells(r, 3).Value = action
    ws.Cells(r, 4).Value = IIf(existed, "TRUE", "FALSE")
    ws.Cells(r, 5).Value = IIf(changed, "TRUE", "FALSE")
    ws.Cells(r, 6).Value = lines
    ws.Cells(r, 7).Value = Now
End Sub

' ----------------------------
' Sheet formatting (simple, stable)
' ----------------------------

Private Sub FormatStart(ByVal ws As Worksheet)
    ws.Cells.UnMerge
    ws.Cells.ClearContents

    ws.Range("A1").Value = "DB2UDT Generator " & DB2UDT_VERSION
    ws.Range("A2").Value = "Ablauf"
    ws.Range("A3").Value = "1) Makro 'Generator_Einrichten' einmal ausfuehren"
    ws.Range("A4").Value = "2) Makro 'DB_Quelle_Importieren' ausfuehren (oder Text in DB_Quelle!A2 einfuegen)"
    ws.Range("A5").Value = "3) Einstellungen pruefen"
    ws.Range("A6").Value = "4) Makro 'UDTs_Erzeugen' ausfuehren (oder 'UDTs_Testlauf')"

    ws.Range("A8").Value = "Hinweise"
    ws.Range("A9").Value = "- Testlauf: keine Dateien schreiben, nur Bericht"
    ws.Range("A10").Value = "- Delta-Modus: ueberspringt unveraenderte Dateien"
    ws.Range("A11").Value = "- UDT-Name = Praefix + Struct-Name"

    ws.Columns("A").ColumnWidth = 80
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
End Sub

Private Sub FormatDBQuelle(ByVal ws As Worksheet)
    ws.Cells.UnMerge
    ws.Cells.ClearContents
    ws.Range("A1").Value = "DB-Quelle hier ab A2 einfuegen (eine Zeile pro Zeile)."
    ws.Range("A1").Font.Bold = True
    ws.Columns("A").ColumnWidth = 120
    ws.Range("A2").Select
End Sub

Private Sub FormatEinstellungen(ByVal ws As Worksheet)
    ws.Cells.UnMerge
    ws.Cells.ClearContents

    ws.Range("A1").Value = "Praefix (z.B. type_)"
    ws.Range("A2").Value = "Veroeffentlicht (TRUE/FALSE)"
    ws.Range("A3").Value = "UDT Version (z.B. 0.1)"
    ws.Range("A4").Value = "Ausgabeordner (optional)"
    ws.Range("A5").Value = "Testlauf (TRUE/FALSE)"
    ws.Range("A6").Value = "Delta-Modus (TRUE/FALSE)"
    ws.Range("A7").Value = "IMPORT_ORDER.txt schreiben (TRUE/FALSE)"

    ws.Range("B1").Value = "type_"
    ws.Range("B2").Value = "TRUE"
    ws.Range("B3").Value = "0.1"
    ws.Range("B4").Value = ""
    ws.Range("B5").Value = "FALSE"
    ws.Range("B6").Value = "TRUE"
    ws.Range("B7").Value = "TRUE"

    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B").ColumnWidth = 60
    ws.Range("A1:A7").Font.Bold = True
End Sub

Private Sub FormatBericht(ByVal ws As Worksheet)
    ws.Cells.UnMerge
    ws.Cells.ClearContents
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 70
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 8
    ws.Columns("G").ColumnWidth = 20
End Sub

' ----------------------------
' Sheet helpers
' ----------------------------

Private Function GetSourceTextFromSheet(ByVal ws As Worksheet) As String
    Dim lastRow As Long: lastRow = LastUsedRow(ws)
    If lastRow < 2 Then
        GetSourceTextFromSheet = ""
        Exit Function
    End If

    Dim lastCol As Long: lastCol = LastUsedCol(ws)
    If lastCol < 1 Then lastCol = 1

    Dim r As Long, c As Long
    Dim sb As String: sb = ""

    For r = 2 To lastRow
        Dim rowLine As String: rowLine = ""

        If lastCol = 1 Then
            rowLine = CStr(ws.Cells(r, 1).Value)
        Else
            ' Paste-case: Excel split by TAB into multiple columns; rebuild line
            For c = 1 To lastCol
                Dim v As String: v = CStr(ws.Cells(r, c).Value)
                If v <> "" Then
                    If rowLine = "" Then
                        rowLine = v
                    Else
                        rowLine = rowLine & vbTab & v
                    End If
                End If
            Next c
        End If

        If rowLine <> "" Then sb = sb & rowLine & vbCrLf
    Next r

    GetSourceTextFromSheet = sb
End Function

Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.name = name
    End If
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim r As Range
    On Error Resume Next
    Set r = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If r Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = r.Row
    End If
End Function

Private Function LastUsedCol(ByVal ws As Worksheet) As Long
    Dim r As Range
    On Error Resume Next
    Set r = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If r Is Nothing Then
        LastUsedCol = 1
    Else
        LastUsedCol = r.Column
    End If
End Function

' ----------------------------
' File dialogs
' ----------------------------

Private Function PickFile(ByVal title As String, ByVal filterDescAndSpec As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    fd.title = title
    fd.AllowMultiSelect = False
    fd.Filters.Clear

    ' filterDescAndSpec: "Desc,*.db;*.scl"
    Dim parts() As String
    parts = Split(filterDescAndSpec, ",")
    If UBound(parts) >= 1 Then
        fd.Filters.Add parts(0), parts(1)
        If UBound(parts) >= 3 Then
            fd.Filters.Add parts(2), parts(3)
        End If
    End If

    If fd.Show <> -1 Then
        PickFile = ""
        Exit Function
    End If
    PickFile = fd.SelectedItems(1)
End Function

Private Function PickFolder(ByVal title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = title
    If fd.Show <> -1 Then
        PickFolder = ""
    Else
        PickFolder = fd.SelectedItems(1)
    End If
End Function

' ----------------------------
' Text file IO (smart read / UTF-8 write)
' ----------------------------

Private Function ReadTextFileSmart(ByVal path As String) As String
    ' Uses ADODB.Stream (late-bound) to support UTF-8 BOM and UTF-16LE BOM.
    Dim charset As String
    charset = DetectCharsetByBOM(path)

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.charset = charset
    stm.Open
    stm.LoadFromFile path
    ReadTextFileSmart = stm.ReadText(-1)
    stm.Close
End Function

Private Function DetectCharsetByBOM(ByVal path As String) As String
    On Error GoTo Fallback

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' binary
    stm.Open
    stm.LoadFromFile path

    Dim b As Variant
    b = stm.Read(3)
    stm.Close

    If IsArray(b) Then
        Dim b0 As Long, b1 As Long, b2 As Long
        b0 = CLng(b(0))
        b1 = CLng(b(1))
        b2 = CLng(b(2))

        If b0 = 239 And b1 = 187 And b2 = 191 Then
            DetectCharsetByBOM = "utf-8"
            Exit Function
        End If
        If b0 = 255 And b1 = 254 Then
            DetectCharsetByBOM = "unicode" ' UTF-16LE
            Exit Function
        End If
        If b0 = 254 And b1 = 255 Then
            DetectCharsetByBOM = "unicode" ' best-effort
            Exit Function
        End If
    End If

Fallback:
    DetectCharsetByBOM = "utf-8"
End Function

Private Sub WriteTextFileUTF8(ByVal path As String, ByVal text As String)
    ' Writes UTF-8 using ADODB.Stream
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.charset = "utf-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2 ' overwrite
    stm.Close
End Sub

' ----------------------------
' String utilities
' ----------------------------

Private Function NormalizeNewlines(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbLf, vbCrLf)
    NormalizeNewlines = s
End Function

Private Function CanonicalText(ByVal s As String) As String
    ' Normalize newlines and trim trailing spaces per line to stabilize delta compare.
    s = NormalizeNewlines(s)
    Dim lines() As String: lines = Split(s, vbCrLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        lines(i) = RTrim$(lines(i))
    Next i
    CanonicalText = Join(lines, vbCrLf)
End Function

Private Sub SplitLineComment(ByVal line As String, ByRef base As String, ByRef cmt As String)
    Dim p As Long: p = InStr(1, line, "//", vbBinaryCompare)
    If p > 0 Then
        base = Left$(line, p - 1)
        cmt = Mid$(line, p)
    Else
        base = line
        cmt = ""
    End If
End Sub

Private Function StripBraces(ByVal s As String) As String
    Dim out As String: out = ""
    Dim i As Long, ch As String, depth As Long
    depth = 0

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "{" Then
            depth = depth + 1
        ElseIf ch = "}" Then
            If depth > 0 Then depth = depth - 1
        Else
            If depth = 0 Then out = out & ch
        End If
    Next i

    StripBraces = out
End Function

Private Function CollapseSpaces(ByVal s As String) As String
    s = Replace(s, vbTab, " ")
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    CollapseSpaces = s
End Function

Private Function StartsWithCI(ByVal s As String, ByVal prefix As String) As Boolean
    StartsWithCI = (UCase$(Left$(Trim$(s), Len(prefix))) = UCase$(prefix))
End Function

Private Function IsValidIdentifier(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim ch As String
    ch = Mid$(s, 1, 1)
    If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or ch = "_") Then Exit Function

    Dim i As Long
    For i = 2 To Len(s)
        ch = Mid$(s, i, 1)
        If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_") Then
            Exit Function
        End If
    Next i
    IsValidIdentifier = True
End Function

Private Function StrToBool(ByVal v As Variant, ByVal defaultVal As Boolean) As Boolean
    Dim s As String: s = UCase$(Trim$(CStr(v)))
    If s = "TRUE" Or s = "1" Or s = "YES" Or s = "JA" Then
        StrToBool = True
    ElseIf s = "FALSE" Or s = "0" Or s = "NO" Or s = "NEIN" Then
        StrToBool = False
    Else
        StrToBool = defaultVal
    End If
End Function

Private Function CountLines(ByVal s As String) As Long
    If Trim$(s) = "" Then CountLines = 0: Exit Function
    CountLines = UBound(Split(NormalizeNewlines(s), vbCrLf)) + 1
End Function

' ----------------------------
' Path utilities
' ----------------------------

Private Function EnsureTrailingBackslash(ByVal folder As String) As String
    folder = Trim$(folder)
    If folder = "" Then
        EnsureTrailingBackslash = ""
        Exit Function
    End If
    If Right$(folder, 1) <> "\\" Then folder = folder & "\\"
    EnsureTrailingBackslash = folder
End Function

Private Function EnsureFolderExists(ByVal folder As String) As Boolean
    On Error GoTo Fail
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folder) Then
        EnsureFolderExists = True
        Exit Function
    End If

    Dim parent As String
    parent = fso.GetParentFolderName(folder)
    If parent <> "" Then
        If Not fso.FolderExists(parent) Then
            Call EnsureFolderExists(EnsureTrailingBackslash(parent))
        End If
    End If

    fso.CreateFolder folder
    EnsureFolderExists = True
    Exit Function

Fail:
    EnsureFolderExists = False
End Function