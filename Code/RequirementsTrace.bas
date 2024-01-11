Attribute VB_Name = "RequirementsTrace"

' Author: Edgar Sevilla
'
'
' The code of the package is dual-license. This means that you can decide which license you wish to use when using the beamer package. The two options are:
'     a) You can use the GNU General Public License, Version 2 or any later version published by the Free Software Foundation.
'     b) You can use the LaTeX Project Public License, version 1.3c or (at your option) any later version.

Option Explicit

Private Const cFIRST_ROW_IN_SHEET As Long = 10
Private Const cFIRST_COLUMN_IN_SHEET As Long = 9
Private Const cROW_SWA_ELEMENTS As Long = 9
Private Const cSHEET_NAME As String = "RequirementsTrace"

Dim LoadedTraceableElementsList As ArrayList
Dim LoadedRequirementsList As ArrayList
Dim removedLinksList As ArrayList
Dim newLinksList As ArrayList

Private Function ResetUserInterface()
    ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Columns.EntireColumn.Hidden = False
    'ActiveSheet.Rows.EntireRow.Hidden = False

End Function


Sub RequirementsTrace_Read_btn_Click()

    OptimizedMode True
    Dim startTime
    Dim elapsedTime
    
    startTime = Time()
    writeCellOnFile_Fx cSHEET_NAME, MainConfig.TraceMkr_OptionalFieldName, "9,5"
    ResetUserInterface
    
    If Not EaRepository Is Nothing Then
        ReqTrc_GeneratePartialQuery4TraceConnectors
        RequirementsTrace_CleanWorkSheet
        RequirementsTrace_LoadSwaElements
        RequirementsTrace_ReadRequirements
        RequirementsTrace_MarkTraces 0, 0
        'ActiveSheet.UsedRange.SpecialCells (xlCellTypeLastCell)
        ActiveSheet.Range("D10").Select
        ActiveWindow.ScrollRow = 10
        ActiveWindow.ScrollColumn = 4 'Column D
        elapsedTime = Format(Now() - startTime, "hh:mm:ss")
        MsgBox ("Process Done" & Chr(10) & "Elapsed time: " & elapsedTime)
    Else
        MsgBox "Please load first a project", vbExclamation, "Sorry!"
    End If
    
    OptimizedMode False
    
End Sub

Sub RequirementsTrace_Write_btn_Click()
    
    Dim startTime
    Dim elapsedTime
    OptimizedMode True
    startTime = Time()
    ResetUserInterface
    
    If Not EaRepository Is Nothing Then
        
        RequirementsTrace_SetLinks

        elapsedTime = Format(Now() - startTime, "hh:mm:ss")
        MsgBox ("Process Done" & Chr(10) & "Elapsed time: " & elapsedTime)
    Else
        MsgBox "Please load first a project", vbExclamation, "Sorry!"
    End If
    
    OptimizedMode False
End Sub

Sub RequirementsTrace_AddElement_btn_Click()
    
    OptimizedMode True
    ResetUserInterface
    
    If Not EaRepository Is Nothing Then
    
        Dim treeSelectedType
        Dim swaElement As EA.element
        Dim swaPackage As EA.Package
        Dim elementsInModel As EA.Collection
        Dim query As String
        Dim foundF As Long
        
        treeSelectedType = EaRepository.GetTreeSelectedItemType()
        If Not LoadedRequirementsList Is Nothing Then
        
            If treeSelectedType = otElement Then
                Set swaElement = EaRepository.GetTreeSelectedObject()
                If MainConfig.IsTraceableElementValid(swaElement.Type, swaElement.stereotype) <> True Then
                    Set swaElement = Nothing
                End If
            
            ElseIf treeSelectedType = otPackage Then
                Set swaPackage = EaRepository.GetTreeSelectedObject()
                
                query = "SELECT *                                             " & Chr(10) & _
                "FROM t_object, t_package                                     " & Chr(10) & _
                "WHERE (                                                      " & Chr(10)
                
                Dim totalTraceableElements As Long
                Dim swa_stereotype
                Dim i As Long
                totalTraceableElements = MainConfig.TraceableElements.Count 'GetTotalTraceableElements
                For i = 0 To totalTraceableElements - 1
                    swa_stereotype = Right(MainConfig.TraceableElements(i), Len(MainConfig.TraceableElements(i)) - InStr(MainConfig.TraceableElements(i), ":"))
                    If i <> (totalTraceableElements - 1) Then
                        query = query & _
                        "       t_object.Stereotype = '" & swa_stereotype & "' OR               " & Chr(10)
                    Else
                        query = query & _
                        "       t_object.Stereotype = '" & swa_stereotype & "'                  " & Chr(10)
                    End If
                Next
               
                query = query & _
                "      ) AND                                                  " & Chr(10) & _
                "      t_package.Package_ID = " & swaPackage.PackageID & " AND         " & Chr(10) & _
                "      t_object.Package_ID = t_package.Package_ID"
                
                'Debug.Print (query)
                Set elementsInModel = EaRepository.GetElementSet(query, 2)
            
            Else
                MsgBox "Wrong selection in EA model. Please check", vbExclamation, "Sorry!"
            End If
            
            
            If (Not swaElement Is Nothing) Or (Not elementsInModel Is Nothing) Then
            
                Dim LC As Long
                Dim LC2 As String
                
                
                LC = xlsW_getLastColumn(cSHEET_NAME, 7)
                If LC < cFIRST_COLUMN_IN_SHEET Then
                    LC = cFIRST_COLUMN_IN_SHEET
                Else
                    LC = LC + 1
                End If
                
                If (Not elementsInModel Is Nothing) Then
                    For Each swaElement In elementsInModel
                        
                        foundF = lookForTraceableElement(swaElement.ElementId)
                        If foundF = -1 Then
                            writeCellOnFile_Fx cSHEET_NAME, swaElement.ElementId, "7," & LC
                            writeCellOnFile_Fx cSHEET_NAME, swaElement.ElementGUID, "8," & LC
                            writeCellOnFile_Fx cSHEET_NAME, swaElement.Name, "9," & LC
                            LC = LC + 1
                            
                            LoadedTraceableElementsList.Add swaElement.ElementId
    
                            query = xlsW_readCell(cSHEET_NAME, 3, 6)
                            
                            If query <> "*" Then
                                RequirementsTrace_MarkTraces swaElement.ElementId, LC
                            End If
                        End If
                        'MsgBox "Package Selected: " & swaElement.Name, vbExclamation, "!!"
                    Next
                ElseIf (Not swaElement Is Nothing) Then
                    foundF = lookForTraceableElement(swaElement.ElementId)
                    If foundF = -1 Then
                    
                        writeCellOnFile_Fx cSHEET_NAME, swaElement.ElementId, "7," & LC
                        writeCellOnFile_Fx cSHEET_NAME, swaElement.ElementGUID, "8," & LC
                        writeCellOnFile_Fx cSHEET_NAME, swaElement.Name, "9," & LC
                        
                        LoadedTraceableElementsList.Add swaElement.ElementId
                        
                        query = xlsW_readCell(cSHEET_NAME, 3, 6)
                        
                        If query <> "*" Then
                            RequirementsTrace_MarkTraces swaElement.ElementId, LC
                        End If
                    Else
                        MsgBox "Element [" & swaElement.Name & "] is already in the target elements", vbExclamation, "Sorry!"
                    End If
                Else
                
                End If
            
            End If
            
            
        Else
            MsgBox "Please Run the 'Read' option first", vbExclamation, "Sorry!"
        End If
    
    Else
        MsgBox "Please load first a project", vbExclamation, "Sorry!"
    End If
    
    OptimizedMode False
    
End Sub

Sub RequirementsTrace_RemoveElement_btn_Click()

    OptimizedMode True
    ResetUserInterface
        Dim SelectedRows As String
        Dim Rng As Range
        Set Rng = selection
        SelectedRows = Rng.Address
        'Debug.Print SelectedRows
        
        Dim ElementTragetID As String
        
        Dim ColumnLetterStr As String
        Dim ColumnLetterEnd As String
        Dim ColumnNumberStr As Long
        Dim ColumnNumberEnd As Long
        Dim RowNumberStr As Long
        Dim RowNumberEnd As Long
        Dim LR As Long
        Dim LC As Long
        LR = xlsW_getLastRow(cSHEET_NAME, 3)
        LC = xlsW_getLastColumn(cSHEET_NAME, 7)
        
        If InStr(SelectedRows, ":") > 0 Then
        
            'Debug.Print SelectedRows
            RowNumberStr = Split(Split(SelectedRows, ":")(0), "$")(2)
            RowNumberEnd = Split(Split(SelectedRows, ":")(1), "$")(2)
            ColumnLetterStr = Split(Split(SelectedRows, ":")(0), "$")(1)
            ColumnLetterEnd = Split(Split(SelectedRows, ":")(1), "$")(1)
            
            ColumnNumberStr = Range(ColumnLetterStr & 1).Column
            ColumnNumberEnd = Range(ColumnLetterEnd & 1).Column
                
            If (RowNumberStr = cROW_SWA_ELEMENTS And RowNumberEnd = cROW_SWA_ELEMENTS) And _
                ColumnNumberStr >= cFIRST_COLUMN_IN_SHEET And _
                ColumnNumberStr <= LC Then
            
                If ColumnNumberEnd > LC Then
                    ColumnNumberEnd = LC
                    ColumnLetterEnd = xlsW_getColumnLetter(LC)
                End If

                Dim i As Long
                For i = ColumnNumberStr To ColumnNumberEnd
                    ElementTragetID = xlsW_readCell(cSHEET_NAME, 7, i)
                    LoadedTraceableElementsList.Remove CLng(ElementTragetID)
                Next
                
                'Delete selection
                'Range(ColumnLetterStr & 7 & ":" & ColumnLetterEnd & LR).Select
                'Range(ColumnLetterStr & 7 & ":" & ColumnLetterEnd & LR).ClearContents
                
                'Copy remaining data
                If ColumnNumberEnd < LC Then
                    ColumnLetterStr = xlsW_getColumnLetter(ColumnNumberEnd + 1)
                    ColumnLetterEnd = xlsW_getColumnLetter(LC)
                    Range(ColumnLetterStr & 7 & ":" & ColumnLetterEnd & LR).Select
                    Range(ColumnLetterStr & 7 & ":" & ColumnLetterEnd & LR).Copy
                
                
                    'Paste data
                    ColumnLetterStr = xlsW_getColumnLetter(ColumnNumberStr)
                    Range(ColumnLetterStr & 7).PasteSpecial xlPasteValues
                End If
                
                'delete last column
                ColumnLetterStr = xlsW_getColumnLetter(LC - (ColumnNumberEnd - ColumnNumberStr))
                Range(ColumnLetterStr & 7 & ":" & ColumnLetterEnd & LR).Select
                Range(ColumnLetterStr & 7 & ":" & ColumnLetterEnd & LR).ClearContents
                
                ColumnLetterStr = xlsW_getColumnLetter(ColumnNumberStr)
                Range(ColumnLetterStr & 9).Select
                
            End If
            

        
        Else
        
        
            RowNumberStr = Split(SelectedRows, "$")(2)
            ColumnLetterStr = Split(SelectedRows, "$")(1)
            ColumnNumberStr = Range(ColumnLetterStr & 1).Column
            'Debug.Print ColumnNumberStr & ":" & RowNumberStr
            
            If RowNumberStr = cROW_SWA_ELEMENTS And _
               ColumnNumberStr >= cFIRST_COLUMN_IN_SHEET And _
               ColumnNumberStr <= LC Then
                
                ElementTragetID = xlsW_readCell(cSHEET_NAME, 7, ColumnNumberStr)
                
                LoadedTraceableElementsList.Remove CLng(ElementTragetID)
                
                'Delete selection
                'Range(ColumnLetterStr & 7 & ":" & ColumnLetterStr & LR).Select
                'Range(ColumnLetterStr & 7 & ":" & ColumnLetterStr & LR).ClearContents
                
                'Copy remaining data
                If ColumnNumberStr < LC Then
                    ColumnLetterStr = xlsW_getColumnLetter(ColumnNumberStr + 1)
                    ColumnLetterEnd = xlsW_getColumnLetter(LC)
                    Range(ColumnLetterStr & 7 & ":" & ColumnLetterEnd & LR).Select
                    Range(ColumnLetterStr & 7 & ":" & ColumnLetterEnd & LR).Copy
                    
                    'Paste data
                    ColumnLetterStr = xlsW_getColumnLetter(ColumnNumberStr)
                    Range(ColumnLetterStr & 7).PasteSpecial xlPasteValues
                End If
                
                'delete last column
                ColumnLetterEnd = xlsW_getColumnLetter(LC)
                Range(ColumnLetterEnd & 7 & ":" & ColumnLetterEnd & LR).Select
                Range(ColumnLetterEnd & 7 & ":" & ColumnLetterEnd & LR).ClearContents
                
                Range(ColumnLetterStr & 9).Select
            
            End If
            
        End If
        'If RowIndex = x And colIndex >= cFIRST_COLUMN_IN_SHEET Then
        
        'End If
        
        'LoadedTraceableElementsList.Add swaElement.ElementId
    
    
    OptimizedMode False

End Sub


Private Function RequirementsTrace_ReadRequirements()

    Dim packageName As String
    Dim i As Long

    
 
    Set LoadedRequirementsList = New ArrayList
    Set removedLinksList = New ArrayList
    Set newLinksList = New ArrayList
    
    Dim elementsInModel As EA.Collection
    Dim elementsInModel2 As EA.Collection
    Dim query As String
    Dim childElement_swa As EA.element
    Dim childElement_swa2 As EA.element
        
    packageName = xlsW_readCell(cSHEET_NAME, 4, 6)
    
    query = "SELECT *                                          " & Chr(10) & _
            "FROM t_object, t_package                          " & Chr(10) & _
            "WHERE object_type = 'Requirement' AND             " & Chr(10) & _
            "      t_package.Name = '" & packageName & "' AND  " & Chr(10) & _
            "      t_object.Package_ID = t_package.Package_ID"
            'Debug.Print (query)
    Set elementsInModel = EaRepository.GetElementSet(query, 2)
    
    i = cFIRST_ROW_IN_SHEET
    For Each childElement_swa In elementsInModel
                              
        Dim tag As EA.TaggedValue
        Dim OptionalTagValue As String
        Dim RequirementTextValue As String
        Dim strOutput As String
    
        OptionalTagValue = ""
        If MainConfig.TraceMkr_ShowOptionalField = True Then
            Set tag = childElement_swa.TaggedValues.GetByName(MainConfig.TraceMkr_OptionalFieldName)
            If Not tag Is Nothing Then
                OptionalTagValue = tag.Value
            End If
        End If
    
        RequirementTextValue = ""
        If MainConfig.TraceMkr_ShowRequirementText = True Then
            RequirementTextValue = Misc_RemoveJunkFromString(childElement_swa.Notes)
        End If
    
        strOutput = _
            childElement_swa.Name & "," & _
            OptionalTagValue & "," & _
            RequirementTextValue & "," & _
            childElement_swa.ElementGUID & ","

        strOutput = "=ROW()-9," & strOutput
        xlsW_writeLineSColumn cSHEET_NAME, strOutput, 3
        LoadedRequirementsList.Add childElement_swa.ElementId
    
        Dim x As String
        Dim LR As Long
        LR = xlsW_getLastRow(cSHEET_NAME, 3) + 1
        Range("H" & cFIRST_ROW_IN_SHEET & ":H" & LR).Font.Color = vbWhite
        x = "=IF(COUNTA(I" & i & ":ZZ" & i & ")>0,1,0)"
        writeCellOnFile_Fx cSHEET_NAME, x, i & ",8"
        i = i + 1
        
    Next
    
End Function

Public Function RequirementsTrace_LoadSwaElements()

    Dim SrsPackage As EA.Package


 
    Set LoadedTraceableElementsList = New ArrayList

    Dim query As String
    Dim childElement_swa As EA.element
    
    Dim k
    Dim arraySplit
    Dim strOutput
    
    query = xlsW_readCell(cSHEET_NAME, 3, 6)
    
    If query <> "" And query <> "*" Then
        
        If InStr(query, ":*sql*: ") > 0 Then
            query = Right(query, Len(query) - 8)
        Else
            'All elements from Package Name
            query = _
                "SELECT                                                      " & Chr(10) & _
                "    t_object.ea_guid,                                       " & Chr(10) & _
                "    t_object.Name,                                          " & Chr(10) & _
                "    t_object.Object_ID                                      " & Chr(10) & _
                "FROM t_object, t_package                                    " & Chr(10) & _
                "WHERE                                                       " & Chr(10) & _
                "    (                                                       " & Chr(10)
                
                Dim totalTraceableElements As Long
                Dim swa_stereotype
                totalTraceableElements = MainConfig.TraceableElements.Count 'GetTotalTraceableElements
                For k = 0 To totalTraceableElements - 1
                    swa_stereotype = Right(MainConfig.TraceableElements(k), Len(MainConfig.TraceableElements(k)) - InStr(MainConfig.TraceableElements(k), ":"))
                    If k <> (totalTraceableElements - 1) Then
                        query = query & _
                        "       t_object.Stereotype = '" & swa_stereotype & "' OR               " & Chr(10)
                    Else
                        query = query & _
                        "       t_object.Stereotype = '" & swa_stereotype & "'                  " & Chr(10)
                    End If
                Next
                
                'Add t_object.Classifier = 0 to remove instances of element
                '"         t_object.Classifier = 0) OR                        " & Chr(10) & _

                query = query & _
                "    ) AND                                                   " & Chr(10) & _
                "    t_package.Name = '" & query & "' AND                    " & Chr(10) & _
                "    t_object.Package_ID = t_package.Package_ID              " & Chr(10) & _
                "ORDER By t_object.Name     "
                'Debug.Print (query)
        End If
        
        
        k = cFIRST_COLUMN_IN_SHEET
        Dim xmloutput As String
        Dim ObjectGuid As String
        Dim ObjectId As String
        Dim ObjectName As String
        Dim xmlRows
        
        Dim i As Long
        xmloutput = EaRepository.sqlQuery(query)
        'Debug.Print (query)
        
        xmlRows = Split(xmloutput, "<Row><ea_guid>")
        If UBound(xmlRows) > 0 Then
            For i = 1 To UBound(xmlRows)
                ObjectGuid = Left(xmlRows(i), InStr(xmlRows(i), "}"))
                ObjectId = Right(xmlRows(i), Len(xmlRows(i)) - (InStr(xmlRows(i), "<Object_ID>") + Len("<Object_ID>") - 1))
                ObjectId = Left(ObjectId, InStr(ObjectId, "</Object_ID>") - 1)
                ObjectName = Right(xmlRows(i), Len(xmlRows(i)) - (InStr(xmlRows(i), "<Name>") + Len("<Name>") - 1))
                ObjectName = Left(ObjectName, InStr(ObjectName, "</Name>") - 1)
                
                If InStr(ObjectGuid, "{") Then
                
                    writeCellOnFile_Fx cSHEET_NAME, ObjectId, "7," & k
                    writeCellOnFile_Fx cSHEET_NAME, ObjectGuid, "8," & k
                    writeCellOnFile_Fx cSHEET_NAME, ObjectName, "9," & k
                    k = k + 1
                    'Debug.Print ObjectName
                    
                    LoadedTraceableElementsList.Add ObjectId
                End If
            Next
        Else
            
        End If
    End If
    
    'For i = 0 To LoadedTraceableElementsList.Count - 1
    '    Debug.Print " + (" & i + 1 & ") " & LoadedTraceableElementsList(i)
    'Next
    
End Function


Public Function RequirementsTrace_MarkTraces(loadedTraceableElement As Long, index As Long)

    Dim i As Long
    Dim k As Long
    Dim j As Long
    Dim LinkFlg As Boolean
    Dim query As String
    Dim xmloutput As String
    Dim xmlRows
    Dim xmlRows2
    Dim reqObjId As Long
    Dim swaObjId As Long
    Dim x As String
    Dim ObjectGuid As String
    Dim ObjectId As String
    Dim ObjectName As String
    Dim LC As Long
    Dim LR As Long
    Dim tracesList As ArrayList
    LR = xlsW_getLastRow(cSHEET_NAME, 3) + 1
    
    If LR < cFIRST_ROW_IN_SHEET Then
        LR = cFIRST_ROW_IN_SHEET
    End If
    
    Range("I" & cFIRST_ROW_IN_SHEET & ":ZZ" & LR).Font.Bold = True
    Range("I" & cFIRST_ROW_IN_SHEET & ":ZZ" & LR).Font.Color = RGB(0, 176, 80)
    Range("I" & cFIRST_ROW_IN_SHEET & ":ZZ" & LR).Font.Name = "Wingdings"
    Range("I" & cFIRST_ROW_IN_SHEET & ":ZZ" & LR).ShrinkToFit = False
        
    
    For i = 0 To LoadedRequirementsList.Count - 1
        reqObjId = LoadedRequirementsList(i)
       
        If loadedTraceableElement = 0 Then
       
            Set tracesList = ReqTrc_GetRequirementTraces(reqObjId)
            For k = 0 To tracesList.Count - 1
                
                swaObjId = tracesList(k)
                j = lookForTraceableElement(swaObjId)
                If j <> -1 Then
                    'LinkFlg = True
                    'Debug.Print " Link found [i] = " & i & ", [k] = " & k
                    writeCellOnFile_Fx cSHEET_NAME, Chr(252), i + 10 & "," & j + 9
                Else
                    query = xlsW_readCell(cSHEET_NAME, 3, 6)
                    If query = "*" Then
                        'Add the element
                        Dim swaElement As EA.element
                        Set swaElement = EaRepository.GetElementByID(swaObjId)

                        LC = xlsW_getLastColumn(cSHEET_NAME, 7)
                        If LC < cFIRST_COLUMN_IN_SHEET Then
                            LC = cFIRST_COLUMN_IN_SHEET
                        Else
                            LC = LC + 1
                        End If
                    
                        writeCellOnFile_Fx cSHEET_NAME, CStr(swaObjId), "7," & LC
                        writeCellOnFile_Fx cSHEET_NAME, swaElement.ElementGUID, "8," & LC
                        writeCellOnFile_Fx cSHEET_NAME, swaElement.Name, "9," & LC
                    
                        LoadedTraceableElementsList.Add swaObjId
                        writeCellOnFile_Fx cSHEET_NAME, Chr(252), i + 10 & "," & LC
                    End If
                End If
            Next
               
        Else
            Debug.Print " [Else] If loadedTraceableElement = 0 "
        End If
    Next

End Function

Function RequirementsTrace_SetLinks()

    Dim i As Long
    Dim j As Long
    Dim LR As Long
    Dim LC As Long
    Dim result As Boolean
    Dim ElementSourceGUID As String
    Dim ElementTargetGUID As String
    Dim theElementSource As EA.element
    Dim theElementTarget As EA.element
    Dim LinkFlag As String
    Dim strOutput As String
    
    LR = xlsW_getLastRow(cSHEET_NAME, 3)
    LC = xlsW_getLastColumn(cSHEET_NAME, 9)
    If LC < cFIRST_COLUMN_IN_SHEET Then
        LC = cFIRST_COLUMN_IN_SHEET
    End If
    
    If removedLinksList Is Nothing Then
        Set removedLinksList = New ArrayList
    End If
    
    If newLinksList Is Nothing Then
        Set newLinksList = New ArrayList
    End If
    
    If LR >= cFIRST_ROW_IN_SHEET And LC >= cFIRST_COLUMN_IN_SHEET Then
        
        
        For i = cFIRST_ROW_IN_SHEET To LR
        
            For j = cFIRST_COLUMN_IN_SHEET To LC
        
                LinkFlag = xlsW_readCell(cSHEET_NAME, i, j)
                
                'Create Link?
                If LinkFlag = "x" Then
                    'Debug.Print " Link required in " & i & ":" & j
                    
                    'Create connection
                    ElementSourceGUID = xlsW_readCell(cSHEET_NAME, i, 7)
                    ElementTargetGUID = xlsW_readCell(cSHEET_NAME, 8, j)
                    
                    'Check entries in the Sheet
                    'Create links bewtwen existent elements
                    If ElementSourceGUID <> Empty And _
                       ElementTargetGUID <> Empty Then
                    
                        Set theElementSource = EaRepository.GetElementByGuid(ElementSourceGUID)
                        Set theElementTarget = EaRepository.GetElementByGuid(ElementTargetGUID)
                        
                        If ((theElementSource Is Nothing) Or (theElementTarget Is Nothing)) Then
                            MsgBox ("Malformed data -> Please check :: GUID was not found [" & i & ":" & j & "]")
                            'change color indicating operation not completed
                            'Range("I" & cFIRST_ROW_IN_SHEET & ":ZZ" & LR).Font.Color = RGB(0, 176, 80)
                        Else
                            
                            result = ConnCreator_createLink(theElementSource, _
                                                            theElementTarget, _
                                                            MainConfig.GetTraceConnectorType(0), _
                                                            MainConfig.GetTraceConnectorStereotypeFull(0))
                            
                            If result = True Then
                                'change symbol on successful operation
                                writeCellOnFile_Fx cSHEET_NAME, Chr(252), i & "," & j
                                
                                'Store information about the links that were created
                                strOutput = theElementSource.Name & " (" & theElementSource.ElementGUID & ") <->" & _
                                            theElementTarget.Name & " (" & theElementTarget.ElementGUID & ")   |  " & _
                                            MainConfig.GetTraceConnectorType(0) & ":" & MainConfig.GetTraceConnectorStereotypeFull(0)
                                newLinksList.Add strOutput
                            End If
                            
                        End If
                    
                    End If
                    
                
                'Remove Link
                ElseIf LinkFlag = "o" Then
                    'Remove connection
                    
                    ElementSourceGUID = xlsW_readCell(cSHEET_NAME, i, 7)
                    ElementTargetGUID = xlsW_readCell(cSHEET_NAME, 8, j)
                    
                    If ElementSourceGUID <> Empty And _
                       ElementTargetGUID <> Empty Then
                       
                        Set theElementSource = EaRepository.GetElementByGuid(ElementSourceGUID)
                        Set theElementTarget = EaRepository.GetElementByGuid(ElementTargetGUID)
                        
                        If ((theElementSource Is Nothing) Or (theElementTarget Is Nothing)) Then
                            MsgBox ("Malformed data -> Please check :: GUID was not found [" & i & ":" & j & "]")
                            'change color indicating operation not completed
                            'Range("I" & cFIRST_ROW_IN_SHEET & ":ZZ" & LR).Font.Color = RGB(0, 176, 80)
                        Else
                            result = ConnRemover_removeLink(theElementSource, theElementTarget, MainConfig.GetTraceConnectorType(0), MainConfig.GetTraceConnectorStereotypeFull(0))
                            
                            If result = True Then
                                'change symbol on successful operation
                                writeCellOnFile_Fx cSHEET_NAME, "", i & "," & j
                                
                                'Store information about the links that were removed
                                strOutput = theElementSource.Name & " (" & theElementSource.ElementGUID & ") <->" & _
                                            theElementTarget.Name & " (" & theElementTarget.ElementGUID & ")   |  " & _
                                            MainConfig.GetTraceConnectorType(0) & ":" & MainConfig.GetTraceConnectorStereotypeFull(0)
                                removedLinksList.Add strOutput
                            
                            End If
                            
                        End If

                    End If
                    
                Else
                    'nothing
                End If
        
            Next
        
        Next
        
        TraceWriteFileOutputs
    
    End If

End Function

Private Function TraceWriteFileOutputs()

    Dim FSO As New FileSystemObject
    Dim FileLinksRemoved
    Dim FileLinksAdded
    Dim strDebug As String
    Dim Item As Variant

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    
    If removedLinksList.Count > 0 Then
    
        Set FileLinksRemoved = FSO.CreateTextFile(ActiveWorkbook.path & "\RequirementsTrace_LinksRemoved.txt")

        strDebug = "Links Removed (" & removedLinksList.Count & ")"
        
        FileLinksRemoved.Write strDebug & Chr(10)
        For Each Item In removedLinksList
            FileLinksRemoved.Write "   " & Item & Chr(10)
        Next Item
        
        FileLinksRemoved.Close
    End If
    
    
    If newLinksList.Count > 0 Then
    
        Set FileLinksAdded = FSO.CreateTextFile(ActiveWorkbook.path & "\RequirementsTrace_LinksAdded.txt")
        
        strDebug = "Links Added (" & newLinksList.Count & ")"
        
        FileLinksAdded.Write strDebug & Chr(10)
        For Each Item In newLinksList
            FileLinksAdded.Write "   " & Item & Chr(10)
        Next Item
        
        FileLinksAdded.Close
    End If
    
End Function


Sub RequirementsTrace_ExportTraceInfo()

    OptimizedMode True
    
    Dim FSO As New FileSystemObject
    Dim FileExport
    Dim i As Long
    Dim j As Long
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set FileExport = FSO.CreateTextFile(ActiveWorkbook.path & "\RequirementsTrace_Matrix.txt")

    Dim LC
    LC = xlsW_getLastColumn(cSHEET_NAME, 7)
    If LC < cFIRST_COLUMN_IN_SHEET Then
        LC = cFIRST_COLUMN_IN_SHEET
    End If
    Dim LR As Long
    LR = xlsW_getLastRow(cSHEET_NAME, 3)
    
    If LR < cFIRST_ROW_IN_SHEET Then
        LR = cFIRST_ROW_IN_SHEET
    End If

    Dim ReqName
    Dim ReqCat
    Dim reqGuid
    Dim SwaElemName
    Dim SwaElemGuid
    Dim SwaElemId
    
    FileExport.Write "[Target]" & Chr(10)
    FileExport.Write xlsW_readCell(cSHEET_NAME, 3, 6) & Chr(10)
    FileExport.Write Chr(10)
    
    FileExport.Write "[SRS Package]" & Chr(10)
    FileExport.Write xlsW_readCell(cSHEET_NAME, 4, 6) & Chr(10)
    FileExport.Write Chr(10)
    
    FileExport.Write "[Requirements]" & Chr(10)
    For i = 10 To LR
        ReqName = xlsW_readCell(cSHEET_NAME, i, 4)
        ReqCat = xlsW_readCell(cSHEET_NAME, i, 5)
        reqGuid = xlsW_readCell(cSHEET_NAME, i, 7)
        FileExport.Write ReqName & "," & ReqCat & ",," & reqGuid & Chr(10)
    Next
    
    FileExport.Write Chr(10)
    FileExport.Write "[SWA Elements]" & Chr(10)
    For i = 9 To LC
        SwaElemName = xlsW_readCell(cSHEET_NAME, 9, i)
        SwaElemGuid = xlsW_readCell(cSHEET_NAME, 8, i)
        SwaElemId = xlsW_readCell(cSHEET_NAME, 7, i)
        FileExport.Write SwaElemName & "," & SwaElemGuid & "," & SwaElemId & Chr(10)
    Next
    
    FileExport.Write Chr(10)
    FileExport.Write "[Traces]" & Chr(10)
        Dim Line As String
        For i = cFIRST_ROW_IN_SHEET To LR
            Line = ""
            For j = cFIRST_COLUMN_IN_SHEET To LC
                Dim val
                
                val = xlsW_readCell(cSHEET_NAME, i, j)
                If val = "" Then
                    val = " "
                End If
                Line = Line & val & ","
            
            Next
            FileExport.Write Line & Chr(10)
        Next
    
    FileExport.Close

    OptimizedMode False
End Sub


Sub RequirementsTrace_ImportTraceInfo()
    OptimizedMode True
    
    Dim FSO As New FileSystemObject
    Dim FileImport
    Dim i As Long
    Dim k As Long
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set FileImport = FSO.OpenTextFile(ActiveWorkbook.path & "\RequirementsTrace_Matrix.txt", ForReading)
    
    Dim LR As Long
    LR = xlsW_getLastRow(cSHEET_NAME, 3) + 1
    
    If LR < cFIRST_ROW_IN_SHEET Then
        LR = cFIRST_ROW_IN_SHEET
    End If
    
    xlsW_cleanRowContent cSHEET_NAME, "C" & cFIRST_ROW_IN_SHEET & ":ZI" & LR
    
    Dim LC As Long
    Dim LC2 As String
    
    LC = xlsW_getLastColumn(cSHEET_NAME, 7)
    If LC < cFIRST_COLUMN_IN_SHEET Then
        LC = cFIRST_COLUMN_IN_SHEET
    End If
    LC2 = xlsW_getColumnLetter(LC)
    
    xlsW_cleanRowContent cSHEET_NAME, "I7:" & LC2 & "9"
    
    Dim PrevtxtLine As String
    Dim txtLine As String
    Dim tCommand As String
    Do Until FileImport.AtEndOfStream
        
        txtLine = FileImport.ReadLine
        'Debug.Print "TL:" & txtLine
        If txtLine = "[Target]" Then
            tCommand = "Target"
            txtLine = FileImport.ReadLine
        ElseIf txtLine = "[SRS Package]" Then
            tCommand = "SRS Package"
            txtLine = FileImport.ReadLine
        ElseIf txtLine = "[Requirements]" Then
            tCommand = "Requirements"
            txtLine = FileImport.ReadLine
            i = 0
        ElseIf txtLine = "[SWA Elements]" Then
            tCommand = "SWA Elements"
            txtLine = FileImport.ReadLine
            k = 9
        ElseIf txtLine = "[Traces]" Then
            tCommand = "Traces"
            txtLine = FileImport.ReadLine
        ElseIf txtLine = "" Then
            tCommand = ""
        Else
            'do nothing
            
        End If
        
        Select Case tCommand
        
            Case "Target"
                writeCellOnFile_Fx cSHEET_NAME, txtLine, "3,6"
            Case "SRS Package"
                writeCellOnFile_Fx cSHEET_NAME, txtLine, "4,6"
            Case "Requirements"
                txtLine = "=ROW()-9," & txtLine
                xlsW_writeLineSColumn cSHEET_NAME, txtLine, 3
                txtLine = "=IF(COUNTA(I" & i + 10 & ":ZZ" & i + 10 & ")>0,1,0)"
                writeCellOnFile_Fx cSHEET_NAME, txtLine, i + 10 & ", 8"
                i = i + 1
            Case "SWA Elements"
                Dim arrayVal
                arrayVal = Split(txtLine, ",")
                writeCellOnFile_Fx cSHEET_NAME, CStr(arrayVal(0)), "9," & k
                writeCellOnFile_Fx cSHEET_NAME, CStr(arrayVal(1)), "8," & k
                writeCellOnFile_Fx cSHEET_NAME, CStr(arrayVal(2)), "7," & k
                k = k + 1
            Case "Traces"
                xlsW_writeLineSColumn cSHEET_NAME, txtLine, 9
            Case Else
                'Do nothing
        End Select
        
    Loop

    FileImport.Close
    OptimizedMode False
End Sub


Private Function lookForTraceableElement(swaObjId As Long) As Long
    Dim k As Long
    Dim bFound As Boolean
    bFound = False
    For k = 0 To LoadedTraceableElementsList.Count - 1
        If swaObjId = LoadedTraceableElementsList(k) Then
            bFound = True
            Exit For
        End If
    Next
    
    If bFound = True Then
        lookForTraceableElement = k
    Else
        lookForTraceableElement = -1
    End If
    
End Function




Public Function RequirementsTrace_CleanWorkSheet()

    'Clean workbook
    Dim LR As Long
    LR = xlsW_getLastRow(cSHEET_NAME, 3) + 1
    
    If LR < cFIRST_ROW_IN_SHEET Then
        LR = cFIRST_ROW_IN_SHEET
    End If
    
    xlsW_cleanRowContent cSHEET_NAME, "C" & cFIRST_ROW_IN_SHEET & ":ZI" & LR
    
    
    Dim LC As Long
    Dim LC2 As String
    
    LC = xlsW_getLastColumn(cSHEET_NAME, 7)
    If LC < cFIRST_COLUMN_IN_SHEET Then
        LC = cFIRST_COLUMN_IN_SHEET
    End If
    LC2 = xlsW_getColumnLetter(LC)
    
    xlsW_cleanRowContent cSHEET_NAME, "I7:" & LC2 & "9"
    
End Function
