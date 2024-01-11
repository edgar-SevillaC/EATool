Attribute VB_Name = "RequirementsReport"

' Author: Edgar Sevilla
'
'
' The code of the package is dual-license. This means that you can decide which license you wish to use when using the beamer package. The two options are:
'     a) You can use the GNU General Public License, Version 2 or any later version published by the Free Software Foundation.
'     b) You can use the LaTeX Project Public License, version 1.3c or (at your option) any later version.

Option Explicit

Private Const cFIRST_ROW_IN_SHEET As Long = 7
Private Const cSHEET_NAME As String = "RequirementsReport"

'Counters
Dim TotalSWARequirements As Long
Dim TotalSWARequirementsCov As Long
Dim TotalSWARequirementsNoCov As Long

Dim TotalSWARequirementsASIL As Long
Dim TotalSWARequirementsASILCov As Long
Dim TotalSWARequirementsASILNoCov As Long

Dim TotalSWARequirementsSecurity As Long
Dim TotalSWARequirementsSecurityCov As Long
Dim TotalSWARequirementsSecurityNoCov As Long


Dim PkgNumber As Long
Dim PkgSWARequirements As Long
Dim PkgSWARequirementsCov As Long
Dim PkgSWARequirementsNoCov As Long

Dim PkgSWARequirementsASIL As Long
Dim PkgSWARequirementsASILCov As Long
Dim PkgSWARequirementsASILNoCov As Long

Dim PkgSWARequirementsSecurity As Long
Dim PkgSWARequirementsSecurityCov As Long
Dim PkgSWARequirementsSecurityNoCov As Long

Dim incorrectLinksList As ArrayList

Sub RequirementsReport_SelectRootPackage_btn_Click()

    If Not EaRepository Is Nothing Then
    
        Dim treeSelectedType
        Dim rootPackage As EA.Package
        
        treeSelectedType = EaRepository.GetTreeSelectedItemType()
        
        If treeSelectedType = otPackage Then
            Set rootPackage = EaRepository.GetTreeSelectedObject()
            Worksheets("Home").Range("txtbx_SpecificPackage").Value = rootPackage.PackageGUID
            Worksheets("Home").Range("txtbx_SpecificPackageName").Value = rootPackage.Name
            
            MsgBox "Package Selected: " & Chr(10) & rootPackage.Name & Chr(10) & rootPackage.PackageGUID, vbInformation
        Else
            MsgBox "Wrong selection in EA model. Please check", vbExclamation, "Sorry!"
        End If
    
    Else
        MsgBox "Please load first a project", vbExclamation, "Sorry!"
    End If
    
End Sub

Sub RequirementsReport_RUN_btn_Click()
    
    OptimizedMode True
    Dim startTime
    Dim elapsedTime
    ' Check if project is already open
    If Not EaRepository Is Nothing Then
    
        startTime = Time()
        
        Dim specificPackageGuid As String
        Set incorrectLinksList = New ArrayList
        specificPackageGuid = Worksheets("Home").Range("txtbx_SpecificPackage").Value
        ReqTrc_GeneratePartialQuery4TraceConnectors
        Requirements_TraceabilityReport specificPackageGuid
        ActiveSheet.Range("E7").Select
        ActiveWindow.ScrollRow = 7
        ActiveWindow.ScrollColumn = 5 'Column E
        elapsedTime = Format(Now() - startTime, "hh:mm:ss")
        RequirementsReportWriteFileOutputs
        MsgBox ("Process Done" & Chr(10) & "Elapsed time: " & elapsedTime)
    Else
        MsgBox "Please load first a project", vbExclamation, "Sorry!"
    End If
    OptimizedMode False
End Sub

Sub RequirementsReport_PtcExtract_btn_Click()

    Dim startTime
    Dim elapsedTime
    
    startTime = Time()
    
    Requirements_ExtToolExport
    
    elapsedTime = Format(Now() - startTime, "hh:mm:ss")
    MsgBox ("Process Done" & Chr(10) & "Elapsed time: " & elapsedTime)

End Sub


Sub RequirementsReport_PtcRead_btn_Click()

    OptimizedMode True
    Dim startTime
    Dim elapsedTime
    
    startTime = Time()
    
    Requirements_ReadFromExtTool
    
    'WorksheetUnlock (False)
    'GenerateCharts
    'WorksheetUnlock (True)
    
    elapsedTime = Format(Now() - startTime, "hh:mm:ss")
    MsgBox ("Process Done" & Chr(10) & "Elapsed time: " & elapsedTime)
    OptimizedMode False
End Sub




Private Function Requirements_TraceabilityReport(strRootPackageGuid As String)

    
    Dim rootPackage As EA.Package
    Dim strOutput As String
    Dim LR As Long
    cleanContentReport

    
    'Debug.Print ("SWA Requirements Coverage -- START script ")
    TotalSWARequirements = 0
    TotalSWARequirementsCov = 0
    TotalSWARequirementsNoCov = 0
    TotalSWARequirementsASIL = 0
    TotalSWARequirementsASILCov = 0
    TotalSWARequirementsASILNoCov = 0
    TotalSWARequirementsSecurity = 0
    TotalSWARequirementsSecurityCov = 0
    TotalSWARequirementsSecurityNoCov = 0
    PkgNumber = 0

    
    Set rootPackage = EaRepository.GetPackageByGuid(strRootPackageGuid)
    
    If Not (rootPackage Is Nothing) Then
    
        DumpPackage rootPackage
        
        strOutput = ",,TOTAL (" & rootPackage.Name & ")," & TotalSWARequirements & "," & TotalSWARequirementsCov & "," & TotalSWARequirementsNoCov & "," & _
                                                          TotalSWARequirementsASIL & "," & TotalSWARequirementsASILCov & "," & TotalSWARequirementsASILNoCov & "," & _
                                                          TotalSWARequirementsSecurity & "," & TotalSWARequirementsSecurityCov & "," & TotalSWARequirementsSecurityNoCov
        xlsW_writeLineSColumn cSHEET_NAME, strOutput, 3
        
        LR = xlsW_getLastRow(cSHEET_NAME, 3) + 1
        
        ActiveSheet.Cells.Cells(LR, 6) = "=SUM(F" & cFIRST_ROW_IN_SHEET & ":F" & LR - 1 & ")"
        ActiveSheet.Cells.Cells(LR, 7) = "=SUM(G" & cFIRST_ROW_IN_SHEET & ":G" & LR - 1 & ")"
        ActiveSheet.Cells.Cells(LR, 8) = "=SUM(H" & cFIRST_ROW_IN_SHEET & ":H" & LR - 1 & ")"
        ActiveSheet.Cells.Cells(LR, 9) = "=SUM(I" & cFIRST_ROW_IN_SHEET & ":I" & LR - 1 & ")"
        ActiveSheet.Cells.Cells(LR, 10) = "=SUM(J" & cFIRST_ROW_IN_SHEET & ":J" & LR - 1 & ")"
        ActiveSheet.Cells.Cells(LR, 11) = "=SUM(K" & cFIRST_ROW_IN_SHEET & ":K" & LR - 1 & ")"
        ActiveSheet.Cells.Cells(LR, 12) = "=SUM(L" & cFIRST_ROW_IN_SHEET & ":L" & LR - 1 & ")"
        ActiveSheet.Cells.Cells(LR, 13) = "=SUM(M" & cFIRST_ROW_IN_SHEET & ":M" & LR - 1 & ")"
        ActiveSheet.Cells.Cells(LR, 14) = "=SUM(N" & cFIRST_ROW_IN_SHEET & ":N" & LR - 1 & ")"
        
    
        'Debug.Print ("============")
        'Debug.Print ("Total Number of SW Requirements :: Covered :: Non Covered")
        'Debug.Print (TotalSWARequirements & " (" & TotalSWARequirementsCov & " / " & TotalSWARequirementsNoCov & ")")
        'Debug.Print ("Total Number of SW Requirements (ASIL) :: Covered :: Non Covered")
        'Debug.Print (TotalSWARequirementsASIL & " (" & TotalSWARequirementsASILCov & " / " & TotalSWARequirementsASILNoCov & ")")
        'Debug.Print ("Total Number of SW Requirements (Security) :: Covered :: Non Covered")
        'Debug.Print (TotalSWARequirementsSecurity & " (" & TotalSWARequirementsSecurityCov & " / " & TotalSWARequirementsSecurityNoCov & ")")
        'Debug.Print ("============")
        'Debug.Print ("Done!")
        
    Else
        MsgBox "Please Check Package GUID", vbExclamation, "Sorry!"
    End If

    
End Function


Private Function DumpPackage(thePackage)
    
    ' Cast thePackage to EA.Package so we get intellisense
    Dim strOutput As String
    Dim currentPackage As EA.Package
    Dim elementsInModel As EA.Collection
    Dim childElement_swa As EA.element
    
    Set currentPackage = thePackage
    

   
    'Get SRS ID
    Dim strTemp As String
    Dim srtSRSId As String
    Dim j As Long
    srtSRSId = ""
    strTemp = Right(currentPackage.Name, 10) ' Extract last 10 characters
    For j = 1 To Len(strTemp)
        If Mid(strTemp, j, 1) >= "0" And Mid(strTemp, j, 1) <= "9" Then
            j = j - 1
            srtSRSId = Right(strTemp, Len(strTemp) - j)
            Exit For
        End If
    Next

    If srtSRSId <> "" Then
    
        PkgSWARequirements = 0
        PkgSWARequirementsCov = 0
        PkgSWARequirementsNoCov = 0
        
        PkgSWARequirementsASIL = 0
        PkgSWARequirementsASILCov = 0
        PkgSWARequirementsASILNoCov = 0
    
        PkgSWARequirementsSecurity = 0
        PkgSWARequirementsSecurityCov = 0
        PkgSWARequirementsSecurityNoCov = 0
    
        'Debug.Print ("##### SRS: " & srtSRSId & " #####")
        
        Dim query As String
        query = "SELECT *                                                 " & Chr(10) & _
                "FROM t_object, t_package                                  " & Chr(10) & _
                "WHERE                                                    " & Chr(10) & _
                "      object_type = 'Requirement'                    AND " & Chr(10) & _
                "      t_package.Name = '" & currentPackage.Name & "' AND " & Chr(10) & _
                "      t_object.Package_ID = t_package.Package_ID"
        
        Set elementsInModel = EaRepository.GetElementSet(query, 2)
        
        For Each childElement_swa In elementsInModel
            'Debug.Print childElement_swa.Name
            Dim bHasTraceResult As Boolean
            bHasTraceResult = ReqTrc_VerifyRequirementTraceability(childElement_swa)
            UpdateStatistics childElement_swa, bHasTraceResult
            
            If bHasTraceResult = False Then
                incorrectLinksList.Add childElement_swa.Name & " - " & currentPackage.Name
            End If
            
        Next
        
        
        PkgNumber = PkgNumber + 1
        strOutput = PkgNumber & "," & srtSRSId & "," & currentPackage.Name & "," & _
                    PkgSWARequirements & "," & PkgSWARequirementsCov & "," & PkgSWARequirementsNoCov & "," & _
                    PkgSWARequirementsASIL & "," & PkgSWARequirementsASILCov & "," & PkgSWARequirementsASILNoCov & "," & _
                    PkgSWARequirementsSecurity & "," & PkgSWARequirementsSecurityCov & "," & PkgSWARequirementsSecurityNoCov
                    
                    
        xlsW_writeLineSColumn cSHEET_NAME, strOutput, 3
        
        'dbgPrint_fx ("Package Number of SW Requirements :: Covered :: Non Covered ")
        'dbgPrint_fx (PkgSWARequirements & " ( " & PkgSWARequirementsCov & " :: " & PkgSWARequirementsNoCov & " )")
        'dbgPrint_fx ("Package Number of SW Requirements (ASIL) :: Covered :: Non Covered ")
        'dbgPrint_fx (PkgSWARequirementsASIL & " ( " & PkgSWARequirementsASILCov & " :: " & PkgSWARequirementsASILNoCov & " )")
        'dbgPrint_fx ("Package Number of SW Requirements (Security) :: Covered :: Non Covered ")
        'dbgPrint_fx (PkgSWARequirementsSecurity & " ( " & PkgSWARequirementsSecurityCov & " :: " & PkgSWARequirementsSecurityNoCov & " )")
        'dbgPrint_fx ("============")
    End If
            
            
    ' Recursively process any child packages
    Dim childPackage As EA.Package
    For Each childPackage In currentPackage.Packages
        
        DumpPackage childPackage

    Next
        
End Function



Private Function UpdateStatistics(theElement, bHasTraceResult As Boolean)
    
    Dim currentElement As EA.element
    Dim Safetytag As EA.TaggedValue
    Dim SecurityTag As EA.TaggedValue
    
    Set currentElement = theElement
    
    TotalSWARequirements = TotalSWARequirements + 1
    
    Set Safetytag = currentElement.TaggedValues.GetByName("Safety")
    Set SecurityTag = currentElement.TaggedValues.GetByName("Security")

    If Not Safetytag Is Nothing Then
        If InStr(Safetytag.Value, "ASIL") Then
            
            TotalSWARequirementsASIL = TotalSWARequirementsASIL + 1
            PkgSWARequirementsASIL = PkgSWARequirementsASIL + 1
            
            If bHasTraceResult = True Then
                PkgSWARequirementsASILCov = PkgSWARequirementsASILCov + 1
                TotalSWARequirementsASILCov = TotalSWARequirementsASILCov + 1
            Else
                PkgSWARequirementsASILNoCov = PkgSWARequirementsASILNoCov + 1
                TotalSWARequirementsASILNoCov = TotalSWARequirementsASILNoCov + 1
            End If
        
        End If
    
    End If
    
    If Not SecurityTag Is Nothing Then
        If SecurityTag.Value = "Yes" Then
            
            TotalSWARequirementsSecurity = TotalSWARequirementsSecurity + 1
            PkgSWARequirementsSecurity = PkgSWARequirementsSecurity + 1
            
            If bHasTraceResult = True Then
                PkgSWARequirementsSecurityCov = PkgSWARequirementsSecurityCov + 1
                TotalSWARequirementsSecurityCov = TotalSWARequirementsSecurityCov + 1
            Else
                PkgSWARequirementsSecurityNoCov = PkgSWARequirementsSecurityNoCov + 1
                TotalSWARequirementsSecurityNoCov = TotalSWARequirementsSecurityNoCov + 1
            End If
        
        End If
    
    End If
    
    
    PkgSWARequirements = PkgSWARequirements + 1
    If bHasTraceResult = True Then
        PkgSWARequirementsCov = PkgSWARequirementsCov + 1
        TotalSWARequirementsCov = TotalSWARequirementsCov + 1
    Else
        PkgSWARequirementsNoCov = PkgSWARequirementsNoCov + 1
        TotalSWARequirementsNoCov = TotalSWARequirementsNoCov + 1
    End If
    
    
End Function


Sub GenerateCharts()

    Dim LR As Long
    LR = xlsW_getLastRow(cSHEET_NAME, 3) + 1
    
    Dim shp As Shape
    
    For Each shp In ActiveSheet.Shapes
        If InStr(shp.Name, "Chart") Then
            shp.Delete
        End If
    Next shp
    
    Dim PercentageCovered As Double
    Dim PercentageCoveredAbs As Double
    Dim xColor As Long
    
    'ExtTool Requirements Diff
    '=============================
    ActiveSheet.Range("F" & LR & ":G" & LR).Select
    ActiveSheet.Shapes.AddChart2(251, xlPie).Select
    ActiveChart.SetSourceData Source:=Range("RequirementsReport!$F$" & LR & ",RequirementsReport!$G$" & LR) '=RequirementsReport!$D$63,RequirementsReport!$J$63,RequirementsReport!$K$63
    ActiveChart.Parent.Width = 240
    ActiveChart.Parent.Height = 180
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Unsync ExtTool Requirements"
    ActiveChart.Parent.Left = Range("S7").Left
    ActiveChart.Parent.Top = Range("S7").Top
    ActiveChart.ChartType = xlDoughnut
    ActiveChart.FullSeriesCollection(1).XValues = "{" & Chr(34) & "Sync in EA" & Chr(34) & "," & Chr(34) & "Not Sync" & Chr(34) & "}"
    
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(1).Select
    With selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
    
    
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(2).Select
    PercentageCovered = ActiveSheet.Range("P" & LR).Value / ActiveSheet.Range("O" & LR).Value
    PercentageCoveredAbs = Abs(PercentageCovered)
    If PercentageCoveredAbs > 0.05 Then
        xColor = RGB(255, 0, 0) 'Red
    ElseIf PercentageCoveredAbs > 0.04 Then
        xColor = RGB(255, 192, 0) ' Orange
    ElseIf PercentageCoveredAbs > 0.025 Then
        xColor = RGB(255, 255, 0) 'Yellow
    Else
        xColor = RGB(0, 176, 80)
    End If
    
    With selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = xColor
        .Transparency = 0
        .Solid
    End With
    
    ActiveChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 90, 80, 70, 30).Select
    selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Format(PercentageCovered * 100, "#.0") & " %"
    With selection.ShapeRange(1).TextFrame2.TextRange.Font
        .Bold = msoTrue
        .Size = 16
        '.ForeColor.RGB = xColor
    End With
    
    
    'Total Requirements
    '=============================
    ActiveSheet.Range("G" & LR & ":H" & LR).Select
    ActiveSheet.Shapes.AddChart2(251, xlPie).Select
    ActiveChart.Parent.Width = 240
    ActiveChart.Parent.Height = 180
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Requirements"
    ActiveChart.SetSourceData Source:=Range("RequirementsReport!$G$" & LR & ":$H$" & LR)
    ActiveChart.Parent.Left = Range("S21").Left
    ActiveChart.Parent.Top = Range("S21").Top


    ActiveChart.ChartType = xlDoughnut
    ActiveChart.FullSeriesCollection(1).XValues = "=RequirementsReport!$G$6:$H$6"
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(2).Select
    With selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
    
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(1).Select

    
    If ActiveSheet.Range("F" & LR).Value > 0 Then
        PercentageCovered = ActiveSheet.Range("G" & LR).Value / ActiveSheet.Range("F" & LR).Value
    Else
        PercentageCovered = 0
    End If
    
    If PercentageCovered < 0.3 Then
        xColor = RGB(255, 0, 0) 'Red
    ElseIf PercentageCovered < 0.5 Then
        xColor = RGB(255, 192, 0) ' Orange
    ElseIf PercentageCovered < 0.8 Then
        xColor = RGB(255, 255, 0) 'Yellow
    Else
        xColor = RGB(0, 176, 80)
    End If
    
    With selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = xColor
        .Transparency = 0
        .Solid
    End With
    
    ActiveChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 90, 80, 70, 30).Select
    selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Format(PercentageCovered * 100, "#.0") & " %"
    With selection.ShapeRange(1).TextFrame2.TextRange.Font
        .Bold = msoTrue
        .Size = 16
        '.ForeColor.RGB = xColor
    End With

    'FuSa Requirements
    '=============================
    ActiveSheet.Range("J" & LR & ":K" & LR).Select
    ActiveSheet.Shapes.AddChart2(251, xlPie).Select
    ActiveChart.Parent.Width = 240
    ActiveChart.Parent.Height = 180
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "ASIL Requirements"
    ActiveChart.SetSourceData Source:=Range("RequirementsReport!$J$" & LR & ":$K$" & LR)
    ActiveChart.Parent.Left = Range("S35").Left
    ActiveChart.Parent.Top = Range("S35").Top


    ActiveChart.ChartType = xlDoughnut
    ActiveChart.FullSeriesCollection(1).XValues = "=RequirementsReport!$J$6:$K$6"
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(2).Select
    With selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
    
    
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(1).Select
    If ActiveSheet.Range("G" & LR).Value > 0 Then
        PercentageCovered = ActiveSheet.Range("J" & LR).Value / ActiveSheet.Range("I" & LR).Value
    Else
        PercentageCovered = 0
    End If
    If PercentageCovered < 0.3 Then
        xColor = RGB(255, 0, 0) 'Red
    ElseIf PercentageCovered < 0.5 Then
        xColor = RGB(255, 192, 0) ' Orange
    ElseIf PercentageCovered < 0.8 Then
        xColor = RGB(255, 255, 0) 'Yellow
    Else
        xColor = RGB(0, 176, 80)
    End If
    
    With selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = xColor
        .Transparency = 0
        .Solid
    End With
    
    ActiveChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 90, 80, 70, 30).Select
    selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Format(PercentageCovered * 100, "#.0") & " %"
    With selection.ShapeRange(1).TextFrame2.TextRange.Font
        .Bold = msoTrue
        .Size = 16
        '.ForeColor.RGB = xColor
    End With

    'Security Requirements
    '=============================
    ActiveSheet.Range("M" & LR & ":N" & LR).Select
    ActiveSheet.Shapes.AddChart2(251, xlPie).Select
    ActiveChart.Parent.Width = 240
    ActiveChart.Parent.Height = 180
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Security Requirements"
    ActiveChart.SetSourceData Source:=Range("RequirementsReport!$M$" & LR & ":$N$" & LR)
    ActiveChart.Parent.Left = Range("S49").Left
    ActiveChart.Parent.Top = Range("S49").Top


    ActiveChart.ChartType = xlDoughnut
    ActiveChart.FullSeriesCollection(1).XValues = "=RequirementsReport!$M$6:$N$6"
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(2).Select
    With selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0
        .Solid
    End With
    
    
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(1).Select
    
    If ActiveSheet.Range("L" & LR).Value > 0 Then
        PercentageCovered = ActiveSheet.Range("M" & LR).Value / ActiveSheet.Range("L" & LR).Value
    Else
        PercentageCovered = 0
    End If
    If PercentageCovered < 0.3 Then
        xColor = RGB(255, 0, 0) 'Red
    ElseIf PercentageCovered < 0.5 Then
        xColor = RGB(255, 192, 0) ' Orange
    ElseIf PercentageCovered < 0.8 Then
        xColor = RGB(255, 255, 0) 'Yellow
    Else
        xColor = RGB(0, 176, 80)
    End If
    
    With selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = xColor
        .Transparency = 0
        .Solid
    End With
    
    ActiveChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 90, 80, 70, 30).Select
    selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Format(PercentageCovered * 100, "#.0") & " %"
    With selection.ShapeRange(1).TextFrame2.TextRange.Font
        .Bold = msoTrue
        .Size = 16
        '.ForeColor.RGB = xColor
    End With

    ActiveSheet.Cells(7, 2).Select
End Sub

Private Function cleanContentReport()

    'Clean workbook
    Dim LR As Long
    LR = xlsW_getLastRow(cSHEET_NAME, 3) + 1
    
    If LR < cFIRST_ROW_IN_SHEET Then
        LR = cFIRST_ROW_IN_SHEET
    End If
    
    xlsW_cleanRowContent cSHEET_NAME, "C" & cFIRST_ROW_IN_SHEET & ":S" & LR

    WorksheetUnlock (False)
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        If InStr(shp.Name, "Chart") Then
            shp.Delete
        End If
    Next shp
    WorksheetUnlock (True)


End Function







Private Function RequirementsReportWriteFileOutputs()

    Dim FSO As New FileSystemObject
    Dim FileIncorrectLinks
    Dim strDebug As String
    Dim Item As Variant

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    
    If incorrectLinksList.Count > 0 Then
    
        Set FileIncorrectLinks = FSO.CreateTextFile(ActiveWorkbook.path & "\RequirementsReport_IncorrectLinks.txt")

        strDebug = "Links Removed (" & incorrectLinksList.Count & ")"
        
        FileIncorrectLinks.Write strDebug & Chr(10)
        
        For Each Item In incorrectLinksList
            FileIncorrectLinks.Write "   " & Item & Chr(10)
        Next Item
        
        FileIncorrectLinks.Close
    End If

    
End Function
