Attribute VB_Name = "RequirementsExtTool"

' Author: Edgar Sevilla
'
'
' The code of the package is dual-license. This means that you can decide which license you wish to use when using the beamer package. The two options are:
'     a) You can use the GNU General Public License, Version 2 or any later version published by the Free Software Foundation.
'     b) You can use the LaTeX Project Public License, version 1.3c or (at your option) any later version.

Option Explicit
Sub Requirements_ExtToolExport()

    Dim path As String
    Dim lngErrorCode As Long
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    Dim sFileName As String
    
    
    sFileName = MainConfig.ExtTool_BatchFile
    
    path = Chr(34) & ActiveWorkbook.path & "\" & sFileName & Chr(34)
    lngErrorCode = wsh.Run(path, windowStyle, waitOnReturn)

    If lngErrorCode <> 0 Then
        MsgBox "Uh oh! Something went wrong with the batch file!"
        Exit Sub
    End If

    Application.Wait (Now + TimeValue("0:00:05"))
    
    
End Sub

Sub Requirements_ReadFromExtTool()

    Dim path As String
    Dim LR As Long
    Dim i As Long
    Dim sExtToolExportFileName As String
    Dim ExtToolExportFound As Boolean
    Dim wb As Workbook
    
    sExtToolExportFileName = MainConfig.ExtTool_ExportFile
    
    For Each wb In Workbooks
        If wb.Name = sExtToolExportFileName Then
    
            Workbooks(sExtToolExportFileName).Close SaveChanges:=False

            Exit For
        End If
    Next wb
    
    
    ExtToolExportFound = False
    
    For Each wb In Workbooks
        If wb.Name = sExtToolExportFileName Then
            ExtToolExportFound = True
            Exit For
        End If
    Next wb
    
    If ExtToolExportFound = False Then
        path = ActiveWorkbook.path & "\" & sExtToolExportFileName
        Set wb = Workbooks.Open(path)
        If Not (wb Is Nothing) Then
            ExtToolExportFound = True
        End If
    End If
    
    If ExtToolExportFound = True Then
      'Clean
      ThisWorkbook.Worksheets("ExtTool_Requirements").Cells.Clear
      
      LR = Workbooks(sExtToolExportFileName).Worksheets("Sheet0").Cells(Rows.Count, 1).End(xlUp).Row
      
      Workbooks(sExtToolExportFileName).Worksheets("Sheet0").Range("A1:H" & LR).Copy _
      ThisWorkbook.Worksheets("ExtTool_Requirements").Range("A1")
    
      Workbooks(sExtToolExportFileName).Close SaveChanges:=False
      
      Dim LR2
      LR2 = ActiveWorkbook.ActiveSheet.Cells(Rows.Count, 3).End(xlUp).Row
      
      For i = 7 To LR2
          ActiveSheet.Cells(i, 15) = "=COUNTIF(ExtTool_Requirements!$A$1:$A$" & LR & ",D" & i & ")"
          ActiveSheet.Cells(i, 16) = "=Q" & i & "-F" & i
          
          Dim fml As String
          Dim j As Integer
          fml = "="
          
          For j = 0 To MainConfig.MatureRequirementStates.Count - 1
          
            fml = fml & "COUNTIFS(ExtTool_Requirements!$A$1:$A$" & LR & ",D" & i & ",ExtTool_Requirements!$B$1:$B$" & LR & "," & Chr(34) & MainConfig.MatureRequirementStates(j) & Chr(34) & ") + " & Chr(10)
          
          Next
          fml = Left(fml, Len(fml) - 3)
          'Debug.Print (fml)
          
          ActiveSheet.Cells.Cells(i, 17) = fml
          
          
          fml = "="
          
          For j = 0 To MainConfig.InmatureRequirementStates.Count - 1
          
            fml = fml & "COUNTIFS(ExtTool_Requirements!$A$1:$A$" & LR & ",D" & i & ",ExtTool_Requirements!$B$1:$B$" & LR & "," & Chr(34) & MainConfig.InmatureRequirementStates(j) & Chr(34) & ") + " & Chr(10)
          
          Next
          fml = Left(fml, Len(fml) - 3)
          'Debug.Print (fml)
          
          ActiveSheet.Cells.Cells(i, 18) = fml
          
          
          fml = "="
          
          For j = 0 To MainConfig.MatureRequirementStates.Count - 1
          
            fml = fml & "COUNTIFS(ExtTool_Requirements!$A$1:$A$" & LR & ",D" & i & ",ExtTool_Requirements!$B$1:$B$" & LR & "," & Chr(34) & MainConfig.MatureRequirementStates(j) & Chr(34) & ",ExtTool_Requirements!$H$1:$H$" & LR & "," & Chr(34) & "ASIL A" & Chr(34) & ") + " & Chr(10)
            fml = fml & "COUNTIFS(ExtTool_Requirements!$A$1:$A$" & LR & ",D" & i & ",ExtTool_Requirements!$B$1:$B$" & LR & "," & Chr(34) & MainConfig.MatureRequirementStates(j) & Chr(34) & ",ExtTool_Requirements!$H$1:$H$" & LR & "," & Chr(34) & "ASIL B" & Chr(34) & ") + " & Chr(10)
            fml = fml & "COUNTIFS(ExtTool_Requirements!$A$1:$A$" & LR & ",D" & i & ",ExtTool_Requirements!$B$1:$B$" & LR & "," & Chr(34) & MainConfig.MatureRequirementStates(j) & Chr(34) & ",ExtTool_Requirements!$H$1:$H$" & LR & "," & Chr(34) & "ASIL B (B)" & Chr(34) & ") + " & Chr(10)
            fml = fml & "COUNTIFS(ExtTool_Requirements!$A$1:$A$" & LR & ",D" & i & ",ExtTool_Requirements!$B$1:$B$" & LR & "," & Chr(34) & MainConfig.MatureRequirementStates(j) & Chr(34) & ",ExtTool_Requirements!$H$1:$H$" & LR & "," & Chr(34) & "ASIL B (D)" & Chr(34) & ") + " & Chr(10)
          Next
          fml = Left(fml, Len(fml) - 3)
          'Debug.Print (fml)
          
          ActiveSheet.Cells.Cells(i, 19) = fml
          
      Next
      LR2 = LR2 + 1
      ActiveSheet.Cells.Cells(LR2, 15) = "=SUM(O7:O" & LR2 - 1 & ")"
      ActiveSheet.Cells.Cells(LR2, 16) = "=Q" & LR2 & "-F" & LR2
      ActiveSheet.Cells.Cells(LR2, 17) = "=SUM(Q7:Q" & LR2 - 1 & ")"
      ActiveSheet.Cells.Cells(LR2, 18) = "=SUM(R7:R" & LR2 - 1 & ")"
      ActiveSheet.Cells.Cells(LR2, 19) = "=SUM(S7:S" & LR2 - 1 & ")"
    End If
    
End Sub

