Attribute VB_Name = "CommonEA"

' Author: Edgar Sevilla
'
'
' The code of the package is dual-license. This means that you can decide which license you wish to use when using the beamer package. The two options are:
'     a) You can use the GNU General Public License, Version 2 or any later version published by the Free Software Foundation.
'     b) You can use the LaTeX Project Public License, version 1.3c or (at your option) any later version.



Option Explicit
' Common EA lib


Public Function EA_ExtractPath(Parent_Id As Long)

    Dim path As String
    Dim query As String
    Dim xmloutput As String
    Dim Rows
    Dim ParentID As String
    Dim packageName As String
    Dim i
    
    query = "Select t_package.Name, t_package.Parent_ID From t_package where t_package.Package_ID = " & Parent_Id
    xmloutput = EaRepository.sqlQuery(query)
    Rows = Split(xmloutput, "<Row><Name>")
    
    For i = 1 To UBound(Rows)
        ParentID = Right(Rows(i), Len(Rows(i)) - (InStr(Rows(i), "<Parent_ID>") + 10))
        ParentID = Left(ParentID, InStr(ParentID, "</Parent_ID>") - 1)
        Parent_Id = CLng(ParentID)
        packageName = Left(Rows(i), (InStr(Rows(i), "<")) - 1)
    Next
    
    If ParentID <> 0 Then
        packageName = EA_ExtractPath(Parent_Id) & "." & packageName
    End If
    
    EA_ExtractPath = packageName
    
End Function



Public Function EA_GetTagValue(Element_Id As Long, tagName As String) As String

    Dim query As String
    Dim xmloutput As String
    Dim xmlRows
    
    query = _
        "SELECT                                                         " & Chr(10) & _
            "t_object.Name,                                             " & Chr(10) & _
            "t_objectproperties.Property,                               " & Chr(10) & _
            "t_objectproperties.Value                                   " & Chr(10) & _
        "FROM t_object                                                  " & Chr(10) & _
        "LEFT JOIN t_objectproperties ON                                " & Chr(10) & _
            "    (t_objectproperties.Object_ID = t_object.Object_ID AND " & Chr(10) & _
            "     t_objectproperties.Property LIKE 'Doc ID')            " & Chr(10) & _
        "WHERE                                                          " & Chr(10) & _
            "t_object.Object_Type = 'Requirement' AND                   " & Chr(10) & _
            "t_objectproperties.Property <> '' AND                      " & Chr(10) & _
            "t_object.Object_ID = 2647166"

    xmloutput = EaRepository.sqlQuery(query)
    xmlRows = Split(xmloutput, "<Row><Name>")
    
    Dim ObjectName As String
    Dim ObjectPropertyName As String
    Dim ObjectPropertyValue As String
    
    If UBound(xmlRows) > 1 Then
        ObjectName = Right(xmlRows(i), Len(xmlRows(i)) - (InStr(xmlRows(i), "<Name>") + Len("<Name>") - 1))
        ObjectName = Left(ObjectName, InStr(ObjectName, "</Name>") - 1)
        ObjectPropertyName = Right(xmlRows(i), Len(xmlRows(i)) - (InStr(xmlRows(i), "<Name>") + Len("<Name>") - 1))
        ObjectPropertyName = Left(ObjectName, InStr(ObjectName, "</Name>") - 1)
        ObjectPropertyValue = Right(xmlRows(i), Len(xmlRows(i)) - (InStr(xmlRows(i), "<Object_ID>") + Len("<Object_ID>") - 1))
        ObjectPropertyValue = Left(ObjectId, InStr(ObjectId, "</Object_ID>") - 1)
        
    Else
        ObjectPropertyValue = ""
    End If
    
    EA_GetTagValue = ObjectPropertyValue

End Function

