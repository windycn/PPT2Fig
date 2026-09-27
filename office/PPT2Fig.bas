Attribute VB_Name = "PPT2Fig"
Option Explicit

' Import this module into a blank PowerPoint presentation, add ribbon.xml,
' and save it as a PowerPoint Add-In (.ppam). No macro is stored in user decks.

Public Sub PPT2FigRibbonExport(ByVal control As IRibbonControl)
    PPT2FigRun control.Id
End Sub

Public Sub PPT2FigRun(Optional ByVal preset As String = "")
    Dim source As Presentation, work As Presentation
    Dim chosen As Selection, selectedShapes As ShapeRange
    Dim tempDir As String, separator As String, token As String
    Dim requestPath As String, responsePath As String, copyPath As String
    Dim nativePdf As String, jobPath As String
    Dim selectedSlides As String, selectedIds As String, currentPage As Long
    Dim scope As String, pages As String, outFolder As String
    Dim fmt As String, pdfMode As String, pdfPackaging As String
    Dim dpi As String, pixelWidth As String, pixelHeight As String
    Dim selectedBounds As String, crop As String, baseName As String
    Dim namingMode As String, exportStamp As String
    Dim widthPt As Double, heightPt As Double
    Dim i As Long, id As Long, countBefore As Long, responseText As String
    Dim slide As Slide, shp As Shape
    Dim stage As String, failureText As String
    Dim presetScope As String, presetFormat As String
    Dim presetDpi As String, presetPdfMode As String, presetPackaging As String

    On Error GoTo Failed
    stage = "read active presentation"
    If Application.Presentations.Count = 0 Then Err.Raise vbObjectError + 100, , "Open a PowerPoint presentation first."
    Set source = ActivePresentation
    Set chosen = ActiveWindow.Selection
    If source.Slides.Count = 0 Then Err.Raise vbObjectError + 101, , "The presentation has no slides."
    widthPt = source.PageSetup.SlideWidth
    heightPt = source.PageSetup.SlideHeight

    currentPage = 1
    On Error Resume Next
    currentPage = ActiveWindow.View.Slide.SlideIndex
    On Error GoTo Failed
    selectedSlides = ""
    selectedIds = ""
    selectedBounds = ""
    If chosen.Type = ppSelectionSlides Then
        For i = 1 To chosen.SlideRange.Count
            selectedSlides = AppendCsv(selectedSlides, CStr(chosen.SlideRange(i).SlideIndex))
        Next i
        currentPage = chosen.SlideRange(1).SlideIndex
    ElseIf chosen.Type = ppSelectionShapes Then
        Set selectedShapes = chosen.ShapeRange
        For i = 1 To selectedShapes.Count
            selectedIds = AppendCsv(selectedIds, CStr(selectedShapes(i).Id))
        Next i
        selectedBounds = ShapeBounds(selectedShapes)
    End If

    Select Case preset
        Case "PPT2FigCurrent": presetScope = "current"
        Case "PPT2FigSelected", "PPT2FigQuickSelected": presetScope = "selected"
        Case "PPT2FigRange": presetScope = "range"
        Case "PPT2FigAll", "PPT2FigQuickAll": presetScope = "all"
        Case "PPT2FigShapes", "PPT2FigQuickShapes": presetScope = "shapes"
    End Select
    If Left$(preset, 10) = "PPT2FigPNG" Then presetFormat = "png"
    If Left$(preset, 10) = "PPT2FigJPG" Then presetFormat = "jpg"
    If Left$(preset, 11) = "PPT2FigTIFF" Then presetFormat = "tiff"
    If Left$(preset, 10) = "PPT2FigPDF" Then presetFormat = "pdf"
    If Left$(preset, 10) = "PPT2FigEPS" Then presetFormat = "eps"
    If Right$(preset, 3) = "150" Then presetDpi = "150"
    If Right$(preset, 3) = "300" Then presetDpi = "300"
    If Right$(preset, 3) = "600" Then presetDpi = "600"
    If Right$(preset, 4) = "1200" Then presetDpi = "1200"
    If Right$(preset, 4) = "2400" Then presetDpi = "2400"
    If Right$(preset, 6) = "Custom" Then presetDpi = "custom"
    If InStr(1, preset, "Raster", vbBinaryCompare) > 0 Then presetPdfMode = "raster"
    If preset = "PPT2FigPDFSeparate" Then presetPackaging = "separate"
    If presetScope = "selected" And Len(selectedSlides) = 0 Then Err.Raise vbObjectError + 114, , "Select slide thumbnails in PowerPoint first."
    If presetScope = "shapes" And Len(selectedIds) = 0 Then Err.Raise vbObjectError + 115, , "Select one or more shapes in PowerPoint first."

#If Mac Then
    separator = "/"
    tempDir = Environ$("TMPDIR")
#Else
    separator = "\"
    tempDir = Environ$("TEMP")
#End If
    If Right$(tempDir, 1) <> separator Then tempDir = tempDir & separator
    Randomize
    token = Format$(Now, "yyyymmddhhnnss") & CStr(Int(Rnd() * 100000))
    tempDir = tempDir & "PPT2Fig-" & token
    stage = "create temporary directory: " & tempDir
    MkDir tempDir
    requestPath = tempDir & separator & "request-" & token & ".ini"
    responsePath = tempDir & separator & "response-" & token & ".ini"
    nativePdf = tempDir & separator & "native-" & token & ".pdf"
    jobPath = tempDir & separator & "job-" & token & ".ini"
    copyPath = tempDir & separator & "copy-" & token & CopyExtension(source.Name)

    stage = "write request: " & requestPath
    WriteUtf16 requestPath, "[request]" & vbLf & _
        "presentation_name=" & CleanIniValue(source.Name) & vbLf & _
        "slide_count=" & CStr(source.Slides.Count) & vbLf & _
        "current_slide=" & CStr(currentPage) & vbLf & _
        "selected_slides=" & selectedSlides & vbLf & _
        "selected_shape_ids=" & selectedIds & vbLf & _
        "preset_scope=" & presetScope & vbLf & _
        "preset_format=" & presetFormat & vbLf & _
        "preset_dpi=" & presetDpi & vbLf & _
        "preset_pdf_mode=" & presetPdfMode & vbLf & _
        "preset_packaging=" & presetPackaging & vbLf & _
        "selection_bounds=" & selectedBounds & vbLf & _
        "slide_width_pt=" & NumberText(widthPt) & vbLf & _
        "slide_height_pt=" & NumberText(heightPt) & vbLf

    stage = "show settings dialog"
    If Not RunHelper("dialog", requestPath, responsePath) Then Err.Raise vbObjectError + 102, , "Cannot launch the PPT2Fig settings window."
    If Dir$(responsePath) = "" Then GoTo Finished  ' User canceled.
    responseText = ReadUtf16(responsePath)
    scope = IniValue(responseText, "scope")
    pages = IniValue(responseText, "pages")
    fmt = IniValue(responseText, "format")
    pdfMode = IniValue(responseText, "pdf_mode")
    pdfPackaging = IniValue(responseText, "pdf_packaging")
    dpi = IniValue(responseText, "dpi")
    pixelWidth = IniValue(responseText, "pixel_width")
    pixelHeight = IniValue(responseText, "pixel_height")
    outFolder = IniValue(responseText, "output_dir")
    baseName = IniValue(responseText, "base_name")
    namingMode = IniValue(responseText, "naming_mode")
    exportStamp = IniValue(responseText, "export_stamp")
    If Len(baseName) = 0 Then baseName = CleanBaseName(source.Name)
    If Len(namingMode) = 0 Then namingMode = "detailed"
    If Len(pages) = 0 Or Len(outFolder) = 0 Then Err.Raise vbObjectError + 103, , "Export settings are incomplete."
    If scope = "shapes" And Len(selectedIds) = 0 Then Err.Raise vbObjectError + 104, , "No shapes are selected."

    ' SaveCopyAs includes current in-memory edits without saving or modifying the source.
    stage = "save presentation copy: " & copyPath
    source.SaveCopyAs copyPath
    stage = "open presentation copy"
    On Error Resume Next
#If Mac Then
    Set work = Application.Presentations.Open(copyPath, msoFalse, msoFalse, msoTrue)
#Else
    Set work = Application.Presentations.Open(copyPath, msoFalse, msoFalse, msoFalse)
    If work Is Nothing Then Set work = Application.Presentations.Open(copyPath, msoFalse, msoFalse, msoTrue)
#End If
    On Error GoTo Failed
    If work Is Nothing Then Err.Raise vbObjectError + 105, , "Cannot open the temporary presentation copy."

    If scope = "shapes" Then
        Set slide = work.Slides(currentPage)
        countBefore = 0
        For i = slide.Shapes.Count To 1 Step -1
            id = slide.Shapes(i).Id
            If Not CsvContains(selectedIds, id) Then
                slide.Shapes(i).Delete
            Else
                countBefore = countBefore + 1
            End If
        Next i
        If countBefore <> selectedShapes.Count Then Err.Raise vbObjectError + 106, , "Selected shape IDs changed in the temporary copy."
        On Error Resume Next
        slide.DisplayMasterShapes = msoFalse
        slide.FollowMasterBackground = msoFalse
        slide.Background.Fill.Visible = msoFalse
        On Error GoTo Failed
        For i = work.Slides.Count To 1 Step -1
            If i <> currentPage Then work.Slides(i).Delete
        Next i
        crop = CropForBounds(selectedBounds, widthPt, heightPt)
    Else
        For i = work.Slides.Count To 1 Step -1
            If Not CsvContains(pages, i) Then work.Slides(i).Delete
        Next i
        crop = ""
    End If

    If work.Slides.Count = 0 Then Err.Raise vbObjectError + 107, , "No slides were selected for export."
    stage = "export native PDF: " & nativePdf
    ExportNativePdf work, nativePdf
    stage = "wait for native PDF"
    For i = 1 To 500
        If Dir$(nativePdf) <> "" Then Exit For
        DoEvents
    Next i
    If Dir$(nativePdf) = "" Then Err.Raise vbObjectError + 108, , "PowerPoint did not produce the temporary PDF."
    stage = "close presentation copy"
    work.Close
    Set work = Nothing

    stage = "write conversion job"
    WriteUtf16 jobPath, "[export]" & vbLf & _
        "source_pdf=" & nativePdf & vbLf & _
        "output_dir=" & CleanIniValue(outFolder) & vbLf & _
        "base_name=" & baseName & vbLf & _
        "naming_mode=" & namingMode & vbLf & _
        "export_stamp=" & exportStamp & vbLf & _
        "format=" & fmt & vbLf & _
        "dpi=" & dpi & vbLf & _
        "pixel_width=" & pixelWidth & vbLf & _
        "pixel_height=" & pixelHeight & vbLf & _
        "pdf_mode=" & pdfMode & vbLf & _
        "pdf_packaging=" & pdfPackaging & vbLf & _
        "selection=" & LCase$(CStr(scope = "shapes")) & vbLf & _
        "page_numbers=" & pages & vbLf & _
        "selection_crop=" & crop & vbLf
    stage = "convert PDF"
    If Not RunHelper("convert", jobPath, "") Then Err.Raise vbObjectError + 109, , "Conversion failed. Check the output folder and PPT2Fig log."
    MsgBox "Export completed:" & vbCrLf & outFolder, vbInformation, "PPT2Fig"

Finished:
    On Error Resume Next
    If Not work Is Nothing Then work.Close
    DeleteIfExists requestPath
    DeleteIfExists responsePath
    DeleteIfExists copyPath
    DeleteIfExists nativePdf
    DeleteIfExists jobPath
    If Len(tempDir) > 0 Then RmDir tempDir
    Exit Sub
Failed:
    failureText = "Export failed at " & stage & ": " & Err.Description & " (" & CStr(Err.Number) & ")"
    MsgBox failureText, vbExclamation, "PPT2Fig"
    Resume Finished
End Sub

Private Sub ExportNativePdf(ByVal pres As Presentation, ByVal path As String)
    Dim exportError As String
    On Error Resume Next
    Err.Clear
#If Mac Then
    ' The Mac VBA object library does not expose ExportAsFixedFormat.
    pres.SaveAs path, ppSaveAsPDF
#Else
    pres.ExportAsFixedFormat Path:=path, FixedFormatType:=ppFixedFormatTypePDF, _
        Intent:=ppFixedFormatIntentPrint, FrameSlides:=msoFalse, _
        OutputType:=ppPrintOutputSlides, PrintHiddenSlides:=msoTrue
    If Err.Number <> 0 Or Dir$(path) = "" Then
        Err.Clear
        pres.SaveAs path, ppSaveAsPDF
    End If
#End If
    If Err.Number <> 0 Then exportError = Err.Description & " (" & CStr(Err.Number) & ")"
    On Error GoTo 0
    If Len(exportError) > 0 Then Err.Raise vbObjectError + 110, , "PowerPoint PDF export failed: " & exportError
End Sub

Private Function ShapeBounds(ByVal shapes As ShapeRange) As String
    Dim leftPt As Double, topPt As Double, rightPt As Double, bottomPt As Double
    Dim i As Long, s As Shape, effectPad As Double
    leftPt = 1E+30: topPt = 1E+30
    rightPt = -1E+30: bottomPt = -1E+30
    For i = 1 To shapes.Count
        Set s = shapes(i)
        effectPad = 1
        On Error Resume Next
        If s.Shadow.Visible = msoTrue Then effectPad = 12 + Abs(s.Shadow.OffsetX) + Abs(s.Shadow.OffsetY)
        If s.Glow.Radius > effectPad Then effectPad = s.Glow.Radius
        On Error GoTo 0
        If s.Left - effectPad < leftPt Then leftPt = s.Left - effectPad
        If s.Top - effectPad < topPt Then topPt = s.Top - effectPad
        If s.Left + s.Width + effectPad > rightPt Then rightPt = s.Left + s.Width + effectPad
        If s.Top + s.Height + effectPad > bottomPt Then bottomPt = s.Top + s.Height + effectPad
    Next i
    ShapeBounds = NumberText(leftPt) & "," & NumberText(topPt) & "," & _
        NumberText(rightPt - leftPt) & "," & NumberText(bottomPt - topPt)
End Function

Private Function CropForBounds(ByVal bounds As String, ByVal slideWidth As Double, ByVal slideHeight As Double) As String
    Dim part() As String, x As Double, yTop As Double, w As Double, h As Double
    Dim xRight As Double, yBottom As Double
    part = Split(bounds, ",")
    If UBound(part) <> 3 Then Err.Raise vbObjectError + 111, , "Invalid selection bounds."
    x = Val(part(0)): yTop = Val(part(1))
    w = Val(part(2)): h = Val(part(3))
    xRight = x + w: yBottom = yTop + h
    If x < 0 Then x = 0
    If yTop < 0 Then yTop = 0
    If xRight > slideWidth Then xRight = slideWidth
    If yBottom > slideHeight Then yBottom = slideHeight
    w = xRight - x: h = yBottom - yTop
    If w <= 0 Or h <= 0 Then Err.Raise vbObjectError + 112, , "Selection is outside the slide."
    CropForBounds = NumberText(x) & "," & NumberText(slideHeight - yTop - h) & "," & _
        NumberText(w) & "," & NumberText(h)
End Function

Private Function NumberText(ByVal value As Double) As String
    NumberText = Trim$(Str$(value))
End Function

Private Function AppendCsv(ByVal existing As String, ByVal item As String) As String
    If Len(existing) = 0 Then AppendCsv = item Else AppendCsv = existing & "," & item
End Function

Private Function CsvContains(ByVal csv As String, ByVal item As Long) As Boolean
    CsvContains = InStr(1, "," & csv & ",", "," & CStr(item) & ",", vbBinaryCompare) > 0
End Function

Private Function CopyExtension(ByVal fileName As String) As String
    Dim pos As Long, ext As String
    pos = InStrRev(fileName, ".")
    If pos > 0 Then ext = LCase$(Mid$(fileName, pos))
    If ext <> ".pptx" And ext <> ".pptm" And ext <> ".ppt" Then ext = ".pptx"
    CopyExtension = ext
End Function

Private Function CleanBaseName(ByVal fileName As String) As String
    Dim pos As Long, i As Long, c As String, result As String
    pos = InStrRev(fileName, ".")
    If pos > 1 Then fileName = Left$(fileName, pos - 1)
    For i = 1 To Len(fileName)
        c = Mid$(fileName, i, 1)
        If InStr(1, "<>:""/\|?*", c, vbBinaryCompare) = 0 Then result = result & c
    Next i
    If Len(result) = 0 Then result = "presentation"
    CleanBaseName = result
End Function

Private Function CleanIniValue(ByVal value As String) As String
    CleanIniValue = Replace(Replace(value, vbCr, " "), vbLf, " ")
End Function

Private Function IniValue(ByVal body As String, ByVal key As String) As String
    Dim line As Variant, prefix As String
    prefix = key & "="
    For Each line In Split(body, vbLf)
        If Left$(CStr(line), Len(prefix)) = prefix Then
            IniValue = Replace(Mid$(CStr(line), Len(prefix) + 1), vbCr, "")
            Exit Function
        End If
    Next line
End Function

Private Sub WriteUtf16(ByVal path As String, ByVal body As String)
    Dim handle As Integer, i As Long, unit As Long, low As Byte, high As Byte
    handle = FreeFile
    Open path For Binary Access Write As #handle
    low = &HFF: high = &HFE
    Put #handle, , low: Put #handle, , high
    For i = 1 To Len(body)
        unit = AscW(Mid$(body, i, 1))
        If unit < 0 Then unit = unit + 65536
        low = CByte(unit And &HFF&)
        high = CByte((unit \ 256) And &HFF&)
        Put #handle, , low: Put #handle, , high
    Next i
    Close #handle
End Sub

Private Function ReadUtf16(ByVal path As String) As String
    Dim handle As Integer, bytes() As Byte, length As Long, i As Long, unit As Long
    handle = FreeFile
    Open path For Binary Access Read As #handle
    length = LOF(handle)
    If length < 2 Then Close #handle: Exit Function
    ReDim bytes(1 To length)
    Get #handle, , bytes
    Close #handle
    For i = 3 To length - 1 Step 2
        unit = CLng(bytes(i)) + CLng(bytes(i + 1)) * 256
        If unit > 32767 Then unit = unit - 65536
        ReadUtf16 = ReadUtf16 & ChrW$(unit)
    Next i
End Function

Private Sub DeleteIfExists(ByVal path As String)
    If Len(path) > 0 Then If Dir$(path) <> "" Then Kill path
End Sub

Private Function RunHelper(ByVal mode As String, ByVal first As String, ByVal second As String) As Boolean
#If Mac Then
    Dim response As String
    response = AppleScriptTask("PPT2FigBridge.scpt", "runPPT2Fig", mode & vbTab & first & vbTab & second)
    RunHelper = (response = "OK")
#Else
    Dim executable As String, command As String, shell As Object
    executable = Environ$("LOCALAPPDATA") & "\PPT2Fig\ppt2fig.exe"
    If Dir$(executable) = "" Then Err.Raise vbObjectError + 113, , "PPT2Fig helper not found: " & executable
    command = """" & executable & """ " & mode & " " & """" & first & """"
    If Len(second) > 0 Then command = command & " " & """" & second & """"
    Set shell = CreateObject("WScript.Shell")
    RunHelper = (shell.Run(command, 1, True) = 0)
#End If
End Function
