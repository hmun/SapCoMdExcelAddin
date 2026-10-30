' Copyright 2023 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPActivityType

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            log.Debug("New - " & "checking connection")
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            log.Error("New - Exception=" & ex.ToString)
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPActivityType")
        End Try
    End Sub

    Private Sub addToStrucDic(pArrayName As String, pRfcStructureMetadata As RfcStructureMetadata, ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        If pStrucDic.ContainsKey(pArrayName) Then
            pStrucDic.Remove(pArrayName)
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Private Sub addToFieldDic(pArrayName As String, pRfcStructureMetadata As RfcParameterMetadata, ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata))
        If pFieldDic.ContainsKey(pArrayName) Then
            pFieldDic.Remove(pArrayName)
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Public Sub getMeta_CreateMultiple(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"LANGUAGE"}
        Dim aImports As String() = {"COAREA", "TESTRUN"}
        Dim aTables As String() = {"ACTIVITYTYPELIST", "RETURN", "EXTENSIONIN"}
        Try
            log.Debug("getMeta_CreateMultiple - " & "creating Function BAPI_ACTTYPE_CREATEMULTIPLE")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_ACTTYPE_CREATEMULTIPLE")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_CreateMultiple - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPActivityType")
        Finally
            log.Debug("getMeta_CreateMultiple - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Sub getMeta_ChangeMultiple(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"LANGUAGE"}
        Dim aImports As String() = {"COAREA", "TESTRUN"}
        Dim aTables As String() = {"ACTIVITYTYPELIST", "RETURN", "EXTENSIONIN"}
        Try
            log.Debug("getMeta_ChangeMultiple - " & "creating Function BAPI_ACTTYPE_CHANGEMULTIPLE")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_ACTTYPE_CHANGEMULTIPLE")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_ChangeMultiple - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPActivityType")
        Finally
            log.Debug("getMeta_ChangeMultiple - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function createMultiple(pData As TSAP_ATData, Optional pOKMsg As String = "OK") As String
        createMultiple = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_ACTTYPE_CREATEMULTIPLE")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            Dim oACTIVITYTYPELIST As IRfcTable = oRfcFunction.GetTable("ACTIVITYTYPELIST")
            oRETURN.Clear()
            oACTIVITYTYPELIST.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the data values
            pData.aDataDic.to_IRfcTable(pKey:="ACTIVITYTYPELIST", pIRfcTable:=oACTIVITYTYPELIST)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean
            aErr = False
            For i As Integer = 0 To oRETURN.Count - 1
                createMultiple = createMultiple & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            createMultiple = If(createMultiple = "", pOKMsg, If(aErr = False, pOKMsg & createMultiple, "Error" & createMultiple))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPActivityType")
            createMultiple = "Error: Exception in createMultiple"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function changeMultiple(pData As TSAP_ATData, Optional pOKMsg As String = "OK") As String
        changeMultiple = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_ACTTYPE_CHANGEMULTIPLE")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            Dim oACTIVITYTYPELIST As IRfcTable = oRfcFunction.GetTable("ACTIVITYTYPELIST")
            oRETURN.Clear()
            oACTIVITYTYPELIST.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the data values
            pData.aDataDic.to_IRfcTable(pKey:="ACTIVITYTYPELIST", pIRfcTable:=oACTIVITYTYPELIST)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean
            aErr = False
            For i As Integer = 0 To oRETURN.Count - 1
                changeMultiple = changeMultiple & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            changeMultiple = If(changeMultiple = "", pOKMsg, If(aErr = False, pOKMsg & changeMultiple, "Error" & changeMultiple))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPActivityType")
            changeMultiple = "Error: Exception in changeMultiple"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
