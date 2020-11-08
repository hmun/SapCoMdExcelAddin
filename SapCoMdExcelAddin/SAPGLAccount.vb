' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPGLAccount

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
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPGLAccount")
        End Try
    End Sub

    Public Function createMultiple(pData As TSAP_GL_ACCData, Optional pOKMsg As String = "OK") As String
        createMultiple = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("GL_ACCT_MASTER_SAVE_RFC")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            Dim oACCOUNT_COA As IRfcStructure = oRfcFunction.GetStructure("ACCOUNT_COA")
            Dim oACCOUNT_NAMES As IRfcTable = oRfcFunction.GetTable("ACCOUNT_NAMES")
            Dim oACCOUNT_KEYWORDS As IRfcTable = oRfcFunction.GetTable("ACCOUNT_KEYWORDS")
            Dim oACCOUNT_CCODES As IRfcTable = oRfcFunction.GetTable("ACCOUNT_CCODES")
            oRETURN.Clear()
            oACCOUNT_NAMES.Clear()
            oACCOUNT_KEYWORDS.Clear()
            oACCOUNT_CCODES.Clear()

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
                Dim aNameArray() As String
                Dim oACCOUNT_COAAppended As Boolean = False
                Dim oACCOUNT_NAMESAppended As Boolean = False
                Dim oACCOUNT_KEYWORDSAppended As Boolean = False
                Dim oACCOUNT_CCODESAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "ACCOUNT_COA"
                            ' these are nested structured so we need to get the structure from the row/structure
                            If InStr(aTStrRec.Fieldname, "-") <> 0 Then
                                aNameArray = Split(aTStrRec.Fieldname, "-")
                                oStruc = oACCOUNT_COA.GetStructure(aNameArray(0))
                                oStruc.SetValue(aNameArray(1), aTStrRec.formated)
                            Else
                                oACCOUNT_COA.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                            End If
                            oACCOUNT_COA.SetValue("ACTION", "I")
                        Case "ACCOUNT_NAMES"
                            If Not oACCOUNT_NAMESAppended Then
                                oACCOUNT_NAMES.Append()
                                oACCOUNT_NAMES.SetValue("ACTION", "I")
                                oACCOUNT_NAMESAppended = True
                            End If
                            ' these are nested structured so we need to get the structure from the row/structure
                            If InStr(aTStrRec.Fieldname, "-") <> 0 Then
                                aNameArray = Split(aTStrRec.Fieldname, "-")
                                oStruc = oACCOUNT_NAMES.GetStructure(aNameArray(0))
                                oStruc.SetValue(aNameArray(1), aTStrRec.formated)
                            Else
                                oACCOUNT_NAMES.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                            End If
                        Case "ACCOUNT_KEYWORDS"
                            If Not oACCOUNT_KEYWORDSAppended Then
                                oACCOUNT_KEYWORDS.Append()
                                oACCOUNT_KEYWORDS.SetValue("ACTION", "I")
                                oACCOUNT_KEYWORDSAppended = True
                            End If
                            ' these are nested structured so we need to get the structure from the row/structure
                            If InStr(aTStrRec.Fieldname, "-") <> 0 Then
                                aNameArray = Split(aTStrRec.Fieldname, "-")
                                oStruc = oACCOUNT_KEYWORDS.GetStructure(aNameArray(0))
                                oStruc.SetValue(aNameArray(1), aTStrRec.formated)
                            Else
                                oACCOUNT_KEYWORDS.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                            End If
                        Case "ACCOUNT_CCODES"
                            If Not oACCOUNT_CCODESAppended Then
                                oACCOUNT_CCODES.Append()
                                oACCOUNT_CCODES.SetValue("ACTION", "I")
                                oACCOUNT_CCODESAppended = True
                            End If
                            ' these are nested structured so we need to get the structure from the row/structure
                            If InStr(aTStrRec.Fieldname, "-") <> 0 Then
                                aNameArray = Split(aTStrRec.Fieldname, "-")
                                oStruc = oACCOUNT_CCODES.GetStructure(aNameArray(0))
                                oStruc.SetValue(aNameArray(1), aTStrRec.formated)
                            Else
                                oACCOUNT_CCODES.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                            End If
                    End Select
                Next
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
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
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPGLAccount")
            createMultiple = "Error: Exception in createMultiple"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
