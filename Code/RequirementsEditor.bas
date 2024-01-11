Attribute VB_Name = "RequirementsEditor"

' Author: Edgar Sevilla
'
'
' The code of the package is dual-license. This means that you can decide which license you wish to use when using the beamer package. The two options are:
'     a) You can use the GNU General Public License, Version 2 or any later version published by the Free Software Foundation.
'     b) You can use the LaTeX Project Public License, version 1.3c or (at your option) any later version.

Option Explicit

Private Const cFIRST_ROW_IN_SHEET As Long = 8
Private Const cFIRST_COL_OPT_IN_SHEET As Long = 9
Private Const cSHEET_NAME As String = "RequirementsCreator"

Dim tagsList As ArrayList

Sub ReqMaker_READ_btn_Click()
    
    OptimizedMode True
    ' Check if project is already open
    If Not EaRepository Is Nothing Then
    
        'ActiveSheet.Cells(3, 8) = Format(Now, "dd/mm/yyyy HH:mm:ss")
        
    
        If xlsW_readCell(cSHEET_NAME, 4, 7) <> "" Then
            ReqMaker_ReadRequirements
        Else
            MsgBox "Please indicate Package name", vbExclamation, "Sorry!"
        End If
    
        'ActiveSheet.Cells(4, 8) = Format(Now, "dd/mm/yyyy HH:mm:ss")
        MsgBox ("Process Done")
    Else
        MsgBox "Please load first a project", vbExclamation, "Sorry!"
    End If
    OptimizedMode False
End Sub

Sub ReqMaker_WRITE_btn_Click()
    
    OptimizedMode True
    ' Check if project is already open
    If Not EaRepository Is Nothing Then
    
        'ActiveSheet.Cells(3, 8) = Format(Now, "dd/mm/yyyy HH:mm:ss")
        
        ReqMaker_WriteRequirements
    
        'ActiveSheet.Cells(4, 8) = Format(Now, "dd/mm/yyyy HH:mm:ss")
        MsgBox ("Process Done")
    Else
        MsgBox "Please load first a project", vbExclamation, "Sorry!"
    End If
    OptimizedMode False
End Sub

Private Function ReqMaker_ReadRequirements()

    
    Dim modelPackage As EA.Package
    'Clean workbook
    Dim LR As Long
    Dim LC As Long
    LR = xlsW_getLastRow(cSHEET_NAME, 3)
    
    If LR < cFIRST_ROW_IN_SHEET Then
        LR = cFIRST_ROW_IN_SHEET
    End If
    
    xlsW_cleanRowContent cSHEET_NAME, "C" & cFIRST_ROW_IN_SHEET & ":V" & LR
 
    Set tagsList = New ArrayList
    
    LC = xlsW_getLastColumn(cSHEET_NAME, 7)
    If LC < cFIRST_COL_OPT_IN_SHEET Then
        LC = cFIRST_COL_OPT_IN_SHEET
    End If
    
    Dim j As Long
    
    For j = cFIRST_COL_OPT_IN_SHEET To LC
        tagsList.Add xlsW_readCell(cSHEET_NAME, cFIRST_ROW_IN_SHEET - 1, j)
    Next
 
    Dim RequrirementsInPackage As EA.Collection
    Dim RequirementElement As EA.element
    Dim packageName As String
    Dim query As String
    packageName = xlsW_readCell(cSHEET_NAME, 4, 7)
    
    
    query = _
        "SELECT *                                        " & Chr(10) & _
        "FROM t_object, t_package                        " & Chr(10) & _
        "WHERE                                           " & Chr(10) & _
        "     object_type = 'Requirement' AND            " & Chr(10) & _
        "     t_package.Name = '" & packageName & "' AND " & Chr(10) & _
        "     t_object.Package_ID = t_package.Package_ID "
        
    Set RequrirementsInPackage = EaRepository.GetElementSet(query, 2)
    
    For Each RequirementElement In RequrirementsInPackage
                
        Dim tag As EA.TaggedValue
        Dim strOutput As String
        Dim TagsInfo As String
        
        TagsInfo = ""
        For j = 0 To tagsList.Count - 1
            
            Set tag = RequirementElement.TaggedValues.GetByName(tagsList(j))
            If Not tag Is Nothing Then
                TagsInfo = TagsInfo & Replace(tag.Value, ",", ";")
            End If
            
            TagsInfo = TagsInfo & ","
        Next
        
        
        strOutput = _
            RequirementElement.Name & "," & _
            "," & _
            RequirementElement.Status & "," & _
            RequirementElement.ElementGUID & "," & _
            TagsInfo

    
        strOutput = "=ROW()-7," & strOutput
        
        xlsW_writeLineSColumn cSHEET_NAME, strOutput, 3
        
        'Write the notes
        writeCellOnFile_Fx cSHEET_NAME, RequirementElement.Notes, xlsW_getLastRow(cSHEET_NAME, 3) & ",5"
        
    Next
    
    LR = xlsW_getLastRow(cSHEET_NAME, 3)
    ActiveSheet.Range("A8:A" & LR).EntireRow.RowHeight = 14.4

End Function


Private Function ReqMaker_WriteRequirements()

    Dim eaRequirement As EA.element
    Dim LR As Long
    Dim LC As Long
    LR = xlsW_getLastRow(cSHEET_NAME, 3)
    
    If LR < cFIRST_ROW_IN_SHEET Then
        LR = cFIRST_ROW_IN_SHEET
    End If
 
    Set tagsList = New ArrayList
    
    LC = xlsW_getLastColumn(cSHEET_NAME, 7)
    If LC < cFIRST_COL_OPT_IN_SHEET Then
        LC = cFIRST_COL_OPT_IN_SHEET
    End If
    
    Dim j As Long
    
    For j = cFIRST_COL_OPT_IN_SHEET To LC
        tagsList.Add xlsW_readCell(cSHEET_NAME, cFIRST_ROW_IN_SHEET - 1, j)
    Next
 
    Dim tag As EA.TaggedValue
    Dim reqGuid As String
    Dim reqID As String
    Dim changeRequiredFlag As String
    Dim i As Long
    For i = cFIRST_ROW_IN_SHEET To LR
        
        changeRequiredFlag = xlsW_readCell(cSHEET_NAME, i, 8)
        If changeRequiredFlag = "x" Then
        
            reqGuid = xlsW_readCell(cSHEET_NAME, i, 7)
            reqID = xlsW_readCell(cSHEET_NAME, i, 4)
            If reqGuid <> "" Then
                'Reading from EA model
                Set eaRequirement = EaRepository.GetElementByGuid(reqGuid)
                If eaRequirement.Name = reqID Then
                    eaRequirement.Status = x
                    eaRequirement.Notes = x
                    
                    For j = 0 To tagsList.Count - 1
                        
                        Set tag = eaRequirement.TaggedValues.GetByName(tagsList(j))
                        If Not tag Is Nothing Then
                            tag.Value = x
                        End If
                        
                        tag.Update
                    Next
                    
                    eaRequirement.Update
                End If
            Else
                'Reading from ExtTool
                'Tbd activity
            End If
        
        End If
        
    Next

End Function

