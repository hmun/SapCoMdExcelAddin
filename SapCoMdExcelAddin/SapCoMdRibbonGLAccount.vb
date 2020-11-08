' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SapCoMdRibbonGLAccount

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapCoMdRibbonGLAccount getGenParametrs - " & "reading Parameter")
        aWB = Globals.SapCoMdExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP CO Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Co Md")
            getGenParameters = False
            Exit Function
        End Try
        aName = "SAPCoMdGLAccount"
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid SAP CO Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Co Md")
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
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SAP CO Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Co Md")
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
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim aSAPGLAccount As New SAPGLAccount(pSapCon, aIntPar)

        aWB = Globals.SapCoMdExcelAddin.Application.ActiveWorkbook
        Dim aDwsName As String = If(aIntPar.value("GL_WS", "DATA") <> "", aIntPar.value("GL_WS", "DATA"), "GLAccount")
        Try
            aDws = aWB.Worksheets(aDwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP GLAccount Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Co Md")
            Exit Sub
        End Try
        ' Read the Items
        Try
            log.Debug("SapCoMdRibbonGLAccount.exec - " & "processing data - disabling events, screen update, cursor")
            aDws.Activate()
            Dim aItems As New TData(aIntPar)
            Dim aItem As New TDataRec
            Dim aKey As String
            Dim j As UInt64
            Dim jMax As UInt64 = 0
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("GL_LOFF", "DATA") <> "", CInt(aIntPar.value("GL_LOFF", "DATA")), 5)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("GL_DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("GL_DBG", "DUMPOBJNR")), 0)
            Dim aMsgClmn As String = If(aIntPar.value("GL_COL", "DATAMSG") <> "", aIntPar.value("GL_COL", "DATAMSG"), "INT-MSG")
            Dim aMsgClmnNr As Integer = 0
            Dim aPostClmn As String = If(aIntPar.value("GL_COL", "DATAPOST") <> "", aIntPar.value("GL_COL", "DATAPOST"), "INT-POST")
            Dim aPostClmnNr As Integer = 0
            Dim aOKMsg As String = If(aIntPar.value("GL_RET", "OKMSG") <> "", aIntPar.value("GL_RET", "OKMSG"), "OK")
            Globals.SapCoMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoMdExcelAddin.Application.EnableEvents = False
            Globals.SapCoMdExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aLOff + 1
            ' determine the last column and create the fieldlist
            Do
                jMax += 1
                If CStr(aDws.Cells(1, jMax).value) = aMsgClmn Then
                    aMsgClmnNr = jMax
                ElseIf CStr(aDws.Cells(1, jMax).value) = aPostClmn Then
                    aPostClmnNr = jMax
                End If
            Loop While CStr(aDws.Cells(aLOff - 3, jMax + 1).value) <> ""
            Dim aPost As String = ""
            aData = New Collection
            j = 1
            Do
                aObjNr += 1
                ' GLAccounts are handled in packages.
                If Left(CStr(aDws.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    For j = 1 To jMax
                        If CStr(aDws.Cells(1, j).value) <> "N/A" And CStr(aDws.Cells(1, j).value) <> "" And CStr(aDws.Cells(1, j).value) <> aMsgClmn Then
                            If aPostClmnNr <> 0 Then
                                aPost = CStr(aDws.Cells(i, aPostClmnNr).value)
                            End If
                            aKey = Str(i)
                            aItems.addValue(aKey, CStr(aDws.Cells(aLOff - 3, j).value), CStr(aDws.Cells(i, j).value),
                                        CStr(aDws.Cells(aLOff - 2, j).value), CStr(aDws.Cells(aLOff - 1, j).value), pEmty:=False,
                                        pEmptyChar:="")
                        End If
                    Next
                    ' if the posting indicator is set, or this is the last line -> call the sap BAPI
                    If String.IsNullOrEmpty(CStr(aDws.Cells(i + 1, 1).value)) Or aPost.ToUpper = "X" Then
                        Dim aTSAP_GL_ACCData As New TSAP_GL_ACCData(aPar, aIntPar)
                        If aTSAP_GL_ACCData.fillHeader(aItems) And aTSAP_GL_ACCData.fillData(aItems) Then
                            ' check if we should dump this document
                            If aObjNr = aDumpObjNr Then
                                log.Debug("SapCoMdRibbonGLAccount.exec - " & "dumping Object Nr " & CStr(aObjNr))
                                aTSAP_GL_ACCData.dumpHeader()
                                aTSAP_GL_ACCData.dumpData()
                            End If
                            ' post the object here
                            If pMode = "Create" Then
                                log.Debug("SapCoMdRibbonGLAccount.exec - " & "calling aSAPGLAccount.createMultiple")
                                aRetStr = aSAPGLAccount.createMultiple(aTSAP_GL_ACCData, aOKMsg)
                                log.Debug("SapCoMdRibbonGLAccount.exec - " & "aSAPGLAccount.createMultiple returned, aRetStr=" & aRetStr)
                                ' message has to be written in all lines that where processed in items
                                For Each aKey In aItems.aTDataDic.Keys
                                    aDws.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                                Next
                            ElseIf pMode = "Change" Then
                                ' log.Debug("SapCoMdRibbonGLAccount.exec - " & "calling aSAPGLAccount.actChangeMultiple")
                                ' aRetStr = aSAPGLAccount.actChangeMultiple(aTSAP_CCData)
                                ' log.Debug("SapCoMdRibbonGLAccount.exec - " & "aSAPGLAccount.actChangeMultiple returned, aRetStr=" & aRetStr)
                                ' aDws.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                            End If
                        Else
                            log.Warn("SapCoMdRibbonGLAccount.exec - " & "Filling Header or Data in aTSAP_GL_ACCData failed!")
                            aDws.Cells(i, aMsgClmnNr) = "Filling Header or Data in aTSAP_GL_ACCData failed!"
                        End If
                        aItems = New TData(aIntPar)
                    End If
                Else
                    aDws.Cells(i, aMsgClmnNr + 1).Value = "ignored - already processed"
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).Value) <> ""

            log.Debug("SapCoMdRibbonGLAccount.exec - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapCoMdExcelAddin.Application.EnableEvents = True
            Globals.SapCoMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapCoMdExcelAddin.Application.EnableEvents = True
            Globals.SapCoMdExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoMdExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapCoMdRibbonGLAccount.exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO Md")
            log.Error("SapCoMdRibbonGLAccount.exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

End Class
