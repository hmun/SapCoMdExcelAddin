Public Class TSAP_CEData

    Public aHdrRec As TDataRec
    Public aData As TData

    Private Hdr_Fields() As String = {"COAREA", "COSTELEMCLASS", "TESTRUN", "LANGU"}
    Private Cel_Fields() As String = {"COST_ELEM", "VALID_FROM", "VALID_TO", "CELEM_CATEGORY", "CELEM_ATTRIBUTE", "RECORD_QUANTITY", "UNIT_OF_MEASURE", "UNIT_OF_MEASURE_ISO", "DEFAULT_COSTCENTER", "DEFAULT_ORDER", "JV_REC_IND", "NAME", "DESCRIPT", "FUNC_AREA", "FUNC_AREA_LONG"}

    Private aAccPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Const sCel As String = "COSTELEMENTLIST"

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
                If valid_Cel_Field(aTStrRec) Then
                    aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sCel)
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

    Public Function valid_Cel_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Cel_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("CEL", aStrucName) Then
            valid_Cel_Field = isInArray(pTStrRec.Fieldname, Cel_Fields)
        End If
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        isInArray = (UBound(Filter(pArray, pString)) > -1)
    End Function

    Public Sub dumpHeader()
        Dim dumpHd As String = If(aIntPar.value("CE_DBG", "DUMPHEADER") <> "", aIntPar.value("CE_DBG", "DUMPHEADER"), "")
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
        Dim dumpDt As String = If(aIntPar.value("CE_DBG", "DUMPDATA") <> "", aIntPar.value("CE_DBG", "DUMPDATA"), "")
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
