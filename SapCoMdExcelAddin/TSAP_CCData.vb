Public Class TSAP_CCData

    Public aHdrRec As TDataRec
    Public aCurRec As TDataRec
    Public aData As TData
    Public aAmounts As Dictionary(Of String, TData)

    Private Hdr_Fields() As String = {"CONTROLLINGAREA", "TESTRUN", "MASTER_DATA_INACTIVE", "LANGU"}
    Private CCE_Fields() As String = {"COSTCENTER", "VALID_FROM", "VALID_TO", "PERSON_IN_CHARGE", "DEPARTMENT", "COSTCENTER_TYPE", "COSTCTR_HIER_GRP", "COMP_CODE", "BUS_AREA", "CURRENCY", "CURRENCY_ISO", "PROFIT_CTR", "RECORD_QUANTITY", "LOCK_IND_ACTUAL_PRIMARY_COSTS", "LOCK_IND_PLAN_PRIMARY_COSTS", "LOCK_IND_ACT_SECONDARY_COSTS", "LOCK_IND_PLAN_SECONDARY_COSTS", "LOCK_IND_ACTUAL_REVENUES", "LOCK_IND_PLAN_REVENUES", "LOCK_IND_COMMITMENT_UPDATE", "CONDITION_TABLE_USAGE", "APPLICATION", "CSTG_SHEET", "ACTY_INDEP_TEMPLATE", "ACTY_DEP_TEMPLATE", "ADDR_TITLE", "ADDR_NAME1", "ADDR_NAME2", "ADDR_NAME3", "ADDR_NAME4", "ADDR_STREET", "ADDR_CITY", "ADDR_DISTRICT", "ADDR_COUNTRY", "ADDR_COUNTRY_ISO", "ADDR_TAXJURCODE", "ADDR_PO_BOX", "ADDR_POSTL_CODE", "ADDR_POBX_PCD", "ADDR_REGION", "TELCO_LANGU", "TELCO_LANGU_ISO", "TELCO_TELEPHONE", "TELCO_TELEPHONE2", "TELCO_TELEBOX", "TELCO_TELEX", "TELCO_FAX_NUMBER", "TELCO_TELETEX", "TELCO_PRINTER", "TELCO_DATA_LINE", "JV_VENTURE", "JV_REC_IND", "JV_EQUITY_TYP", "JV_OTYPE", "JV_JIBCL", "JV_JIBSA", "NAME", "DESCRIPT", "FUNC_AREA", "ACTY_DEP_TEMPLATE_ALLOC_CC", "ACTY_INDEP_TEMPLATE_ALLOC_CC", "FUNC_AREA_LONG", "PERSON_IN_CHARGE_USER", "LOGSYSTEM", "ACTY_DEP_TEMPLATE_SK", "ACTY_INDEP_TEMPLATE_SK"}

    Private aPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Const sCce As String = "COSTCENTERLIST"

    Public Sub New(ByRef pPar As SAPCommon.TStr, ByRef pIntPar As SAPCommon.TStr)
        aPar = pPar
        aIntPar = pIntPar
    End Sub

    Public Function fillHeader(pData As TData) As Boolean
        aHdrRec = New TDataRec
        Dim aPostRec As New TDataRec
        Dim aKvb As KeyValuePair(Of String, SAPCommon.TStrRec)
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNewHdrRec As New TDataRec
        aPostRec = pData.getFirstRecord()
        For Each aKvb In aPar.getData()
            aTStrRec = aKvb.Value
            If valid_Hdr_Field(aTStrRec) Then
                aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmptyChar:="")
            End If
        Next
        ' First fill the value from the paramters and then overwrite then from the posting record
        If Not IsNothing(aPostRec) Then
            For Each aTStrRec In aPostRec.aTDataRecCol
                If valid_Hdr_Field(aTStrRec) Then
                    If aTStrRec.Strucname = "HD" Then
                        aNewHdrRec.setValues("-" & aTStrRec.Fieldname, aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
                    Else
                        aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
                    End If
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
                If valid_CCE_Field(aTStrRec) Then
                    aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sCce)
                End If
            Next
            aCnt += 1
        Next
    End Function

    Public Function valid_Hdr_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Hdr_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("HD", aStrucName) Or isInArray("LANGUAGE", aStrucName) Or String.IsNullOrEmpty(pTStrRec.Strucname) Then
            valid_Hdr_Field = isInArray(pTStrRec.Fieldname, Hdr_Fields)
        End If
    End Function

    Public Function valid_Cce_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Cce_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("COSTCENTERLIST", aStrucName) Or isInArray("CCE", aStrucName) Then
            valid_Cce_Field = isInArray(pTStrRec.Fieldname, CCE_Fields)
        End If
    End Function

    Public Function valid_Ext_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        Dim aValExtString As String = If(aIntPar.value("STR", "VALEXT") <> "", aIntPar.value("STR", "VALEXT"), "")
        valid_Ext_Field = False
        aStrucName = Split(aValExtString, ",")
        If isInArray(pTStrRec.Strucname, aStrucName) Then
            valid_Ext_Field = True
        End If
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        isInArray = (UBound(Filter(pArray, pString)) > -1)
    End Function

    Public Function getCC() As String
        Dim aTStrRec As SAPCommon.TStrRec
        getCC = ""
        For Each aTStrRec In aHdrRec.aTDataRecCol
            If aTStrRec.Fieldname = "COSTCENTER" Then
                getCC = aTStrRec.Value
                Exit Function
            End If
        Next
    End Function

    Public Sub dumpHeader()
        Dim dumpHd As String = If(aIntPar.value("CC_DBG", "DUMPHEADER") <> "", aIntPar.value("CC_DBG", "DUMPHEADER"), "")
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
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO Md")
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
        Dim dumpDt As String = If(aIntPar.value("CC_DBG", "DUMPDATA") <> "", aIntPar.value("CC_DBG", "DUMPDATA"), "")
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
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO Md")
                Exit Sub
            End Try
            log.Debug("dumpData - " & "dumping to " & dumpDt)
            ' clear the Data
            If CStr(aDWS.Cells(5, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(5, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If

            Dim aKvB As KeyValuePair(Of String, TDataRec)
            Dim aKvB_Am As KeyValuePair(Of String, TDataRec)
            Dim aData_Am As New TData(aIntPar)
            Dim aDataRec As New TDataRec
            Dim aDataRec_Am As New TDataRec
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
            Next
        End If
    End Sub

End Class
