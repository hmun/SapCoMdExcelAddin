' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector
Public Class SAPCostElement

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon
    Private aIntPar As SAPCommon.TStr

    Sub New(aSapCon As SapCon, ByRef pIntPar As SAPCommon.TStr)
        aIntPar = pIntPar
        Try
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostElement")
        End Try
    End Sub

    Public Function createMultiple(pData As TSAP_CEData, Optional pOKMsg As String = "OK") As String
        createMultiple = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTELEM_CREATEMULTIPLE")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            Dim oCOSTELEMENTLIST As IRfcTable = oRfcFunction.GetTable("COSTELEMENTLIST")
            oRETURN.Clear()
            oCOSTELEMENTLIST.Clear()

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
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            For Each aKvP In pData.aData.aTDataDic
                Dim oCOSTELEMENTLISTAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "COSTELEMENTLIST"
                            If Not oCOSTELEMENTLISTAppended Then
                                oCOSTELEMENTLIST.Append()
                                oCOSTELEMENTLISTAppended = True
                            End If
                            oCOSTELEMENTLIST.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
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
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostElement")
            createMultiple = "Error: Exception in createMultiple"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function changeMultiple(pData As TSAP_CEData, Optional pOKMsg As String = "Success") As String
        changeMultiple = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTELEM_CHANGEMULTIPLE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            Dim oCOSTELEMENTLIST As IRfcTable = oRfcFunction.GetTable("COSTELEMENTLIST")
            oRETURN.Clear()
            oCOSTELEMENTLIST.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Fieldname <> "COSTELEMCLASS" Then
                    If aTStrRec.Strucname <> "" Then
                        oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                        oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                    Else
                        oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                    End If
                End If
            Next
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            For Each aKvP In pData.aData.aTDataDic
                Dim oCOSTELEMENTLISTAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "COSTELEMENTLIST"
                            If Not oCOSTELEMENTLISTAppended Then
                                oCOSTELEMENTLIST.Append()
                                oCOSTELEMENTLISTAppended = True
                            End If
                            oCOSTELEMENTLIST.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
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
                aSAPBapiTranctionCommit.commit()
            End If
            changeMultiple = If(changeMultiple = "", pOKMsg, If(aErr = False, pOKMsg & changeMultiple, "Error" & changeMultiple))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostElement")
            changeMultiple = "Error: Exception in changeMultiple"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
