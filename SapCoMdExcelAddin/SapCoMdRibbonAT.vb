' Copyright 2023 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SapCoMdRibbonAT

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapCoMdRibbon_Mat getGenParametrs - " & "reading Parameter")
        aWB = Globals.SapCoMdExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid Sap CO Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO Md")
            getGenParameters = False
            Exit Function
        End Try
        aName = "SAPCoMdActivityType"
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid Sap CO Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO Md")
            getGenParameters = False
            Exit Function
        End If
        i = 2
        pPar = New SAPCommon.TStr
        Do While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
            pPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 4).value), pFORMAT:=CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop
        getGenParameters = True
    End Function

    Private Function getIntParameters(ByRef pIntPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = Globals.SapCoMdExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid Sap CO Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO Md")
            getIntParameters = False
            Exit Function
        End Try
        i = 2
        pIntPar = New SAPCommon.TStr
        Do
            pIntPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters check - we should know what we are doing
        getIntParameters = True
    End Function

    Public Sub exec(ByRef pSapCon As SapCon, Optional pMode As String = "Create")

        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        Dim aData As Collection

        Dim aRetStr As String
        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        Dim aCoArea As String = aPar.value("", "COAREA")
        Dim aLanguage As String = aPar.value("LANGUAGE", "LANGU")
        Dim aTest As Boolean = If(aPar.value("", "TESTRUN") = "X", True, False)
        If String.IsNullOrEmpty(aCoArea) Or String.IsNullOrEmpty(aLanguage) Then
            MsgBox("Please fill all obligatory fields in the parameter sheet!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO Md")
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim aSAPActivityType As New SAPActivityType(pSapCon)

        aWB = Globals.SapCoMdExcelAddin.Application.ActiveWorkbook
        Dim aDwsName As String = If(aIntPar.value("AT_WS", "DATA") <> "", aIntPar.value("AT_WS", "DATA"), "ActivityType")
        Try
            aDws = aWB.Worksheets(aDwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Activity Type Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO Md")
            Exit Sub
        End Try
        ' Read the Items
        Try
            log.Debug("SapCoMdRibbonAT.exec - " & "processing data - disabling events, screen update, cursor")
            aDws.Activate()
            Dim aItems As New TData(aIntPar)
            Dim aItem As New TDataRec
            Dim aKey As String
            Dim j As UInt64
            Dim jMax As UInt64 = 0
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("AT_LOFF", "DATA") <> "", CInt(aIntPar.value("AT_LOFF", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("AT_DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("AT_DBG", "DUMPOBJNR")), 0)
            Dim aMsgClmn As String = If(aIntPar.value("AT_COL", "DATAMSG") <> "", aIntPar.value("AT_COL", "DATAMSG"), "INT-MSG")
            Dim aMsgClmnNr As Integer = 0
            Dim aPostClmn As String = If(aIntPar.value("AT_COL", "DATAPOST") <> "", aIntPar.value("AT_COL", "DATAPOST"), "INT-POST")
            Dim aPostClmnNr As Integer = 0
            Dim aOKMsg As String = If(aIntPar.value("AT_RET", "OKMSG") <> "", aIntPar.value("AT_RET", "OKMSG"), "OK")

            Globals.SapCoMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoMdExcelAddin.Application.EnableEvents = False
            Globals.SapCoMdExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aLOff + 1
            ' determine the last column and create the fieldlist
            parseHeaderLine(aDws, jMax, aMsgClmn, aMsgClmnNr, aPostClmn, aPostClmnNr, pHdrLine:=aLOff - 3)
            Dim aPost As String = ""
            aData = New Collection
            j = 1
            Do
                aObjNr += 1
                ' Cost centers are handled line by line or in packages.
                If Left(CStr(aDws.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    If aPostClmnNr <> 0 Then
                        aPost = CStr(aDws.Cells(i, aPostClmnNr).value)
                    End If
                    aKey = CStr(i)
                    aItems.ws_parse_line_simple(aDws, aLOff, i, jMax)
                    ' aItem = aItems.aTDataDic(aKey)
                    ' if the posting indicator is set, or this is the last line -> call the sap BAPI
                    If String.IsNullOrEmpty(CStr(aDws.Cells(i + 1, 1).value)) Or aPost.ToUpper = "X" Then
                        Dim aTSAP_ATData As New TSAP_ATData(aPar, aIntPar, aSAPActivityType, pMode)
                        If aTSAP_ATData.fillHeader(aItems) And aTSAP_ATData.fillData(aItems) Then
                            ' check if we should dump this document
                            If aObjNr = aDumpObjNr Then
                                log.Debug("SapCoMdRibbonAT.exec - " & "dumping Object Nr " & CStr(aObjNr))
                                aTSAP_ATData.dumpHeader()
                                aTSAP_ATData.dumpData()
                            End If
                            ' post the object here
                            If pMode = "Create" Then
                                log.Debug("SapCoMdRibbonAT.exec - " & "calling aSAPActivityTypet.createMultiple")
                                aRetStr = aSAPActivityType.createMultiple(aTSAP_ATData, aOKMsg)
                                log.Debug("SapCoMdRibbonAT.exec - " & "aSAPActivityType.createMultiple returned, aRetStr=" & aRetStr)
                                ' message has to be written in all lines that where processed in items
                                For Each aKey In aItems.aTDataDic.Keys
                                    aDws.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                                Next
                            ElseIf pMode = "Change" Then
                                log.Debug("SapCoMdRibbonAT.exec - " & "calling aSAPActivityTypet.changeMultiple")
                                aRetStr = aSAPActivityType.changeMultiple(aTSAP_ATData, aOKMsg)
                                log.Debug("SapCoMdRibbonAT.exec - " & "aSAPActivityType.changeMultiple returned, aRetStr=" & aRetStr)
                                ' message has to be written in all lines that where processed in items
                                For Each aKey In aItems.aTDataDic.Keys
                                    aDws.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                                Next
                            End If
                        Else
                            log.Warn("SapCoMdRibbonAT.exec - " & "filling Header or Data in aTSAP_ATData failed!")
                            aDws.Cells(i, aMsgClmnNr) = "Filling Header or Data in aTSAP_ATData failed!"
                        End If
                        aItems = New TData(aIntPar)
                    End If
                Else
                    aDws.Cells(i, aMsgClmnNr + 1).Value = "ignored - already processed"
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).Value) <> ""
            log.Debug("SapCoMdRibbonAT.exec - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapCoMdExcelAddin.Application.EnableEvents = True
            Globals.SapCoMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapCoMdExcelAddin.Application.EnableEvents = True
            Globals.SapCoMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapCoMdRibbonAT.exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO Md")
            log.Error("SapCoMdRibbonAT.exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Sub parseHeaderLine(ByRef pWs As Excel.Worksheet, ByRef pMaxJ As Integer,
                                Optional pMsgClmn As String = "", Optional ByRef pMsgClmnNr As Integer = 0,
                                Optional pPostClmn As String = "", Optional ByRef pPostClmnNr As Integer = 0,
                                Optional pHdrLine As Integer = 1)
        pMaxJ = 0
        Do
            pMaxJ += 1
            If Not String.IsNullOrEmpty(pMsgClmn) And CStr(pWs.Cells(pHdrLine, pMaxJ).value) = pMsgClmn Then
                pMsgClmnNr = pMaxJ
            ElseIf CStr(pWs.Cells(pHdrLine, pMaxJ).value) = pPostClmn Then
                pPostClmnNr = pMaxJ
            End If
        Loop While CStr(pWs.Cells(pHdrLine, pMaxJ + 1).value) <> ""
    End Sub

End Class
