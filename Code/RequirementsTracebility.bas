Attribute VB_Name = "RequirementsTracebility"

' Author: Edgar Sevilla
'
'
' The code of the package is dual-license. This means that you can decide which license you wish to use when using the beamer package. The two options are:
'     a) You can use the GNU General Public License, Version 2 or any later version published by the Free Software Foundation.
'     b) You can use the LaTeX Project Public License, version 1.3c or (at your option) any later version.


Option Explicit

Dim ReqTrc_queryTraceConnectors As String



Public Function ReqTrc_VerifyRequirementTraceability(theElement) As Boolean
    
    ' Iterate through all elements and add them to the list
    Dim currentRequirement As EA.element
    Dim currentConnector As EA.connector
    Dim swaElement As EA.element
    Dim bFoundFlg As Boolean
    Dim i As Long
    Dim Guids
    Dim xmlWithGuids As String
    Dim query As String
    
    
    Set currentRequirement = theElement
    bFoundFlg = False
    
    If MainConfig.IsRequirementValid(currentRequirement.Status) = True Then
        
        If MainConfig.IsValidConnector4TraceabilityCheckEnabled() = False Then
            query = _
                 "SELECT t_connector.ea_guid                                       " & Chr(10) & _
                 "FROM t_connector                                                 " & Chr(10) & _
                 "WHERE                                                            " & Chr(10) & _
                 "    t_connector.Connector_Type like '*' AND  " & Chr(10) & _
                 "    t_connector.Start_Object_ID = " & currentRequirement.ElementId
                                  
        Else
            query = _
                 "SELECT t_connector.ea_guid                                               " & Chr(10) & _
                 "FROM t_connector                                                         " & Chr(10) & _
                 "WHERE                                                                    " & Chr(10) & _
                 "    t_connector.Start_Object_ID = " & currentRequirement.ElementId & " AND  " & Chr(10)
                 
            query = query & ReqTrc_queryTraceConnectors

        End If
        xmlWithGuids = EaRepository.sqlQuery(query)
        'Debug.Print (query)
        Guids = Split(xmlWithGuids, "<Row><ea_guid>")
        
        For i = 1 To UBound(Guids)
            'Guids(i) = Replace(Guids(i), "</ea_guid></Row>", "")
            Guids(i) = Left(Guids(i), InStr(Guids(i), "}"))
            If InStr(Guids(i), "{") Then
            
                Set currentConnector = EaRepository.GetConnectorByGuid(Guids(i))
                
                If MainConfig.IsTraceConnectorValid(currentConnector.Type, currentConnector.stereotype, currentConnector.FQStereotype) = True Then
                    If MainConfig.IsValidTraceableElementCheckEnabled() = True Then
                        Set swaElement = EaRepository.GetElementByID(currentConnector.SupplierID)
                        If Not swaElement Is Nothing Then
                            If MainConfig.IsTraceableElementValid(swaElement.Type, swaElement.stereotype) = True And _
                               MainConfig.IsTraceableElementStateValid(swaElement.Status) = True Then
                                bFoundFlg = True
                                Exit For
                            End If
                        End If
                    Else
                        bFoundFlg = True
                        Exit For
                    End If
                Else
                    'Debug.Print ("Invalid Connector " & currentConnector.FQStereotype)
                End If
            End If
        Next
    End If
    
    ReqTrc_VerifyRequirementTraceability = bFoundFlg
    
End Function


Public Function ReqTrc_GetRequirementTraces(reqObjId) As ArrayList
    
    Dim tracesList As ArrayList
    Dim swaElement As EA.element
    Dim i As Long
    Dim xmloutput As String
    Dim xmlRows
    Dim query As String
    Dim Guid As String
    Dim currentConnector As EA.connector
    
    Set tracesList = New ArrayList
        

    
    If MainConfig.IsValidConnector4TraceabilityCheckEnabled() = False Then
        query = _
             "SELECT t_connector.ea_guid       " & Chr(10) & _
             "FROM t_connector                                        " & Chr(10) & _
             "WHERE                                                   " & Chr(10) & _
             "    t_connector.Connector_Type like '*'            AND  " & Chr(10) & _
             "    t_connector.Start_Object_ID = " & reqObjId
             
             
                              
    Else
        query = _
             "SELECT t_connector.ea_guid                              " & Chr(10) & _
             "FROM t_connector                                        " & Chr(10) & _
             "WHERE                                                   " & Chr(10) & _
             "    t_connector.Start_Object_ID = " & reqObjId & " AND  " & Chr(10)
             
        query = query & ReqTrc_queryTraceConnectors

    End If
    
    'Debug.Print (query)
    xmloutput = EaRepository.sqlQuery(query)

    xmlRows = Split(xmloutput, "<Row><ea_guid>")
    
    For i = 1 To UBound(xmlRows)
        Guid = Left(xmlRows(i), InStr(xmlRows(i), "}"))
        
        Set currentConnector = EaRepository.GetConnectorByGuid(Guid)
        
        If MainConfig.IsTraceConnectorValid(currentConnector.Type, currentConnector.stereotype, currentConnector.FQStereotype) = True Then
            If MainConfig.IsValidTraceableElementCheckEnabled() = True Then
                
                Set swaElement = EaRepository.GetElementByID(currentConnector.SupplierID)
            
                If Not swaElement Is Nothing Then
                    If MainConfig.IsTraceableElementValid(swaElement.Type, swaElement.stereotype) = True And _
                       MainConfig.IsTraceableElementStateValid(swaElement.Status) = True Then
                        tracesList.Add currentConnector.SupplierID
                    End If
                End If
            Else
                tracesList.Add currentConnector.SupplierID
            End If
        Else
            'Debug.Print ("Invalid Connector " & currentConnector.FQStereotype)
        End If
    Next

    
    Set ReqTrc_GetRequirementTraces = tracesList
    
End Function


'=====================================================


Public Function ReqTrc_GeneratePartialQuery4TraceConnectors()

    Dim i As Long
    Dim connector As String
    Dim stereotype As String
    Dim connectorlist As String
    Dim stereotypelist As String
    
    connectorlist = ""
    stereotypelist = ""
    For i = 0 To MainConfig.TraceConnectors.Count - 1
        connector = MainConfig.GetTraceConnectorType(i)
        
        If i <> (MainConfig.TraceConnectors.Count - 1) Then
            connectorlist = connectorlist & _
                    "'" & connector & "', "
        Else
            connectorlist = connectorlist & _
                    "'" & connector & "'"
        End If
    Next
    
    
    For i = 0 To MainConfig.TraceConnectors.Count - 1
        stereotype = MainConfig.GetTraceConnectorStereotype(i)
        
        If i <> (MainConfig.TraceConnectors.Count - 1) Then
            If stereotype <> "empty" And stereotype <> "error" Then
                stereotypelist = stereotypelist & _
                    "'" & stereotype & "', "
            End If
        Else
            If stereotype <> "empty" And stereotype <> "error" Then
                stereotypelist = stereotypelist & _
                    "'" & stereotype & "'"
            End If
        End If
    Next

    ReqTrc_queryTraceConnectors = "    t_connector.Connector_Type IN (" & _
                                  connectorlist
    
    If stereotypelist <> "" Then
        ReqTrc_queryTraceConnectors = ReqTrc_queryTraceConnectors & _
        ")    AND     " & Chr(10) & _
        "    t_connector.Stereotype IN (" & _
        stereotypelist & _
        ")"
        
    Else
        ReqTrc_queryTraceConnectors = ReqTrc_queryTraceConnectors & _
        ")"
    End If
    
    'Debug.Print (ReqTrc_queryTraceConnectors)

End Function
