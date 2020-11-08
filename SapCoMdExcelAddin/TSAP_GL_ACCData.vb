Public Class TSAP_GL_ACCData

    Public aHdrRec As TDataRec
    Public aData As TData

    Private Hdr_Fields() As String = {"TESTMODE", "NO_SAVE_AT_WARNING", "NO_AUTHORITY_CHECK"}
    Private Coa_Fields() As String = {"KEYY-KTOPL", "KEYY-SAKNR", "DATA-KTOKS", "DATA-XBILK", "DATA-GVTYP", "DATA-FUNC_AREA", "DATA-MUSTR", "DATA-VBUND", "DATA-BILKT", "DATA-XLOEV", "DATA-XSPEA", "DATA-XSPEB", "DATA-XSPEP", "INFO-ERDAT", "INFO-ERNAM", "INFO-SAKAN", "ACTION"}
    Private Nam_Fields() As String = {"KEYY-KTOPL", "KEYY-SAKNR", "KEYY-SPRAS", "DATA-TXT20", "DATA-TXT50", "ACTION"}
    Private Key_Fields() As String = {"KTOPL", "SAKNR", "SPRAS", "SCHLW", "ACTION"}
    Private Cco_Fields() As String = {"KEYY-BUKRS", "KEYY-SAKNR", "DATA-WAERS", "DATA-XSALH", "DATA-KDFSL", "DATA-BEWGP", "DATA-MWSKZ", "DATA-XMWNO", "DATA-MITKZ", "DATA-ALTKT", "DATA-WMETH", "DATA-INFKY", "DATA-TOGRU", "DATA-XOPVW", "DATA-XKRES", "DATA-ZUAWA", "DATA-BEGRU", "DATA-BUSAB", "DATA-FSTAG", "DATA-XINTB", "DATA-XNKON", "DATA-XMITK", "DATA-FDLEV", "DATA-XGKON", "DATA-FIPOS", "DATA-HBKID", "DATA-HKTID", "DATA-VZSKZ", "DATA-ZINDT", "DATA-ZINRT", "DATA-DATLZ", "DATA-RECID", "DATA-XSPEB", "DATA-XLOEB", "DATA-XLGCLR", "DATA-MCAKEY", "INFO-ERDAT", "INFO-ERNAM", "ACTION"}

    Private aAccPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Const sCoa As String = "ACCOUNT_COA"
    Private Const sNam As String = "ACCOUNT_NAMES"
    Private Const sKey As String = "ACCOUNT_KEYWORDS"
    Private Const sCco As String = "ACCOUNT_CCODES"

    Public Sub New(ByRef pAccPar As SAPCommon.TStr, ByRef pIntPar As SAPCommon.TStr)
        aAccPar = pAccPar
        aIntPar = pIntPar
    End Sub

    Public Function fillHeader(pData As TData) As Boolean
        aHdrRec = New TDataRec
        Dim aPostRec As New TDataRec
        Dim aKvb As KeyValuePair(Of String, SAPCommon.TStrRec)
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNewHdrRec As New TDataRec
        aPostRec = pData.getPostingRecord()
        For Each aKvb In aAccPar.getData()
            aTStrRec = aKvb.Value
            If valid_Hdr_Field(aTStrRec) Then
                aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmptyChar:="")
            End If
        Next
        ' First fill the value from the paramters and then overwrite then from the posting record
        If Not IsNothing(aPostRec) Then
            For Each aTStrRec In aPostRec.aTDataRecCol
                If valid_Hdr_Field(aTStrRec) Then
                    aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
                End If
            Next
        End If
        aHdrRec = aNewHdrRec
        fillHeader = True
    End Function

    Public Function fillData(pData As TData) As Boolean
        Dim aKvB As KeyValuePair(Of String, TDataRec)
        Dim aTDataRec As TDataRec
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aCnt As UInt64
        aData = New TData(aIntPar)
        fillData = True
        aCnt = 1
        For Each aKvB In pData.aTDataDic
            aTDataRec = aKvB.Value
            ' add the valid gl-account fields
            For Each aTStrRec In aTDataRec.aTDataRecCol
                If valid_Coa_Field(aTStrRec) Then
                    aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sCoa)
                End If
                If valid_Nam_Field(aTStrRec) Then
                    aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sNam)
                End If
                If valid_Key_Field(aTStrRec) Then
                    aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sKey)
                End If
                If valid_Cco_Field(aTStrRec) Then
                    aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sCco)
                End If
            Next
            aCnt += 1
            Next
    End Function

    Public Function valid_Hdr_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Hdr_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("HD", aStrucName) Or String.IsNullOrEmpty(pTStrRec.Strucname) Then
            valid_Hdr_Field = isInArray(pTStrRec.Fieldname, Hdr_Fields)
        End If
    End Function

    Public Function valid_Nam_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Nam_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("NAM", aStrucName) Then
            valid_Nam_Field = isInArray(pTStrRec.Fieldname, Nam_Fields)
        End If
    End Function

    Public Function valid_Cco_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Cco_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("CCO", aStrucName) Then
            valid_Cco_Field = isInArray(pTStrRec.Fieldname, Cco_Fields)
        End If
    End Function

    Public Function valid_Key_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Key_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("KEY", aStrucName) Then
            valid_Key_Field = isInArray(pTStrRec.Fieldname, Key_Fields)
        End If
    End Function

    Public Function valid_Coa_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Coa_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("COA", aStrucName) Then
            valid_Coa_Field = isInArray(pTStrRec.Fieldname, Coa_Fields)
        End If
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        isInArray = (UBound(Filter(pArray, pString)) > -1)
    End Function

    Public Sub dumpHeader()
        Dim dumpHd As String = If(aIntPar.value("DBG", "DUMPHEADER") <> "", aIntPar.value("DBG", "DUMPHEADER"), "")
        If dumpHd <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapCoMdExcelAddin.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpHd)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpHeader - " & "No " & dumpHd & " Sheet in current workbook.")
                MsgBox("No " & dumpHd & " Sheet in current workbook. Check the DBG-DUMPHEADR Parameter",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
                Exit Sub
            End Try
            log.Debug("dumpHeader - " & "dumping to " & dumpHd)
            ' clear the Header
            If CStr(aDWS.Cells(1, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If
            ' dump the Header
            Dim aTStrRec As New SAPCommon.TStrRec
            Dim aFieldArray() As String = {}
            Dim aValueArray() As String = {}
            For Each aTStrRec In aHdrRec.aTDataRecCol
                Array.Resize(aFieldArray, aFieldArray.Length + 1)
                aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                Array.Resize(aValueArray, aValueArray.Length + 1)
                aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
            Next
            aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1, aFieldArray.Length))
            aRange.Value = aFieldArray
            aRange = aDWS.Range(aDWS.Cells(2, 1), aDWS.Cells(2, aValueArray.Length))
            aRange.Value = aValueArray
        End If
    End Sub

    Public Sub dumpData()
        Dim dumpDt As String = If(aIntPar.value("DBG", "DUMPDATA") <> "", aIntPar.value("DBG", "DUMPDATA"), "")
        If dumpDt <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapCoMdExcelAddin.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpDt)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpData - " & "No " & dumpDt & " Sheet in current workbook.")
                MsgBox("No " & dumpDt & " Sheet in current workbook. Check the DBG-DUMPDATA Parameter",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
                Exit Sub
            End Try
            log.Debug("dumpData - " & "dumping to " & dumpDt)
            ' clear the Data
            If CStr(aDWS.Cells(5, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(5, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If

            Dim aKvB As KeyValuePair(Of String, TDataRec)
            Dim aDataRec As New TDataRec
            Dim i As Int64
            Dim aTStrRec As New SAPCommon.TStrRec
            i = 6
            For Each aKvB In aData.aTDataDic
                aDataRec = aKvB.Value
                Dim aFieldArray() As String = {}
                Dim aValueArray() As String = {}
                For Each aTStrRec In aDataRec.aTDataRecCol
                    Array.Resize(aFieldArray, aFieldArray.Length + 1)
                    aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                    Array.Resize(aValueArray, aValueArray.Length + 1)
                    aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
                Next
                aRange = aDWS.Range(aDWS.Cells(i, 1), aDWS.Cells(i, aFieldArray.Length))
                aRange.Value = aFieldArray
                aRange = aDWS.Range(aDWS.Cells(i + 1, 1), aDWS.Cells(i + 1, aValueArray.Length))
                aRange.Value = aValueArray
                i += 2
            Next
        End If
    End Sub

End Class
