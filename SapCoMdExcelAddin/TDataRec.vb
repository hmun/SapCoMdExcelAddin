Public Class TDataRec

    Public aTDataRecCol As Collection
    Public Sub New()
        aTDataRecCol = New Collection
    End Sub

    Public Sub setValues(pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String,
                         Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set")
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNameArray() As String
        Dim aKey As String
        Dim aSTRUCNAME As String = ""
        Dim aFIELDNAME As String = ""
        ' do not add empty values
        If Not pEmty And pVALUE = pEmptyChar Then
            Exit Sub
        End If
        If InStr(pNAME, "-") <> 0 Then
            aNameArray = Split(pNAME, "-")
            aSTRUCNAME = aNameArray(0)
            For i As Integer = 1 To aNameArray.Length - 1
                aFIELDNAME = If(String.IsNullOrEmpty(aFIELDNAME), aNameArray(i), aFIELDNAME & "-" & aNameArray(i))
            Next
        Else
            aSTRUCNAME = ""
            aFIELDNAME = pNAME
        End If
        aKey = pNAME
        If aTDataRecCol.Contains(aKey) Then
            aTStrRec = aTDataRecCol(aKey)
            Select Case pOperation
                Case "add"
                    aTStrRec.addValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
                Case "sub"
                    aTStrRec.subValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
                Case "mul"
                    aTStrRec.mulValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
                Case "div"
                    aTStrRec.divValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
                Case Else
                    aTStrRec.setValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
            End Select
        Else
            aTStrRec = New SAPCommon.TStrRec
            aTStrRec.setValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
            aTDataRecCol.Add(aTStrRec, aKey)
        End If
    End Sub

    Public Sub setValues(pTDataRec As TDataRec, Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set")
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pTDataRec.aTDataRecCol
            setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmty, pEmptyChar, pOperation)
        Next
    End Sub

    Public Sub addValues(pTDataRec As TDataRec, Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#")
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pTDataRec.aTDataRecCol
            If aTStrRec.Currency <> "" Then
                setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmty, pEmptyChar, pOperation:="add")
            Else
                setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmty, pEmptyChar, pOperation:="set")
            End If
        Next
    End Sub

    Public Function getPost(ByRef pPar As SAPCommon.TStr) As String
        Dim aClmn As String = If(pPar.value("COL", "DATAPOST") <> "", pPar.value("COL", "DATAPOST"), "INT-POST")
        Dim aTStrRec As SAPCommon.TStrRec
        getPost = ""
        If aTDataRecCol.Contains(aClmn) Then
            aTStrRec = aTDataRecCol(aClmn)
            getPost = aTStrRec.Value
        End If
    End Function

    Public Function getAccType(ByRef pPar As SAPCommon.TStr) As String
        Dim aClmn As String = If(pPar.value("COL", "DATAACCTYPE") <> "", pPar.value("COL", "DATAACCTYPE"), "INT-ACCTYPE")
        Dim aTStrRec As SAPCommon.TStrRec
        getAccType = ""
        If aTDataRecCol.Contains(aClmn) Then
            aTStrRec = aTDataRecCol(aClmn)
            getAccType = aTStrRec.Value
        End If
    End Function

    Public Function getAccount(ByRef pPar As SAPCommon.TStr) As String
        Dim aClmn As String = If(pPar.value("COL", "DATAACCOUNT") <> "", pPar.value("COL", "DATAACCOUNT"), "INT-ACCOUNT")
        Dim aTStrRec As SAPCommon.TStrRec
        getAccount = ""
        If aTDataRecCol.Contains(aClmn) Then
            aTStrRec = aTDataRecCol(aClmn)
            getAccount = aTStrRec.Value
        End If
    End Function

    Public Function getAccTStrRec(ByRef pPar As SAPCommon.TStr) As SAPCommon.TStrRec
        Dim aClmn As String = If(pPar.value("COL", "DATAACCOUNT") <> "", pPar.value("COL", "DATAACCOUNT"), "INT-ACCOUNT")
        Dim aTStrRec As SAPCommon.TStrRec
        getAccTStrRec = New SAPCommon.TStrRec
        If aTDataRecCol.Contains(aClmn) Then
            getAccTStrRec = aTDataRecCol(aClmn)
        End If
    End Function


    Public Function getIsPa(ByRef pPar As SAPCommon.TStr) As Boolean
        Dim aClmn As String = If(pPar.value("COL", "DATAISPA") <> "", pPar.value("COL", "DATAISPA"), "INT-ISPA")
        Dim aTStrRec As SAPCommon.TStrRec
        getIsPa = False
        If aTDataRecCol.Contains(aClmn) Then
            aTStrRec = aTDataRecCol(aClmn)
            If aTStrRec.Value = "Y" Or aTStrRec.Value = "y" Or aTStrRec.Value = "X" Or aTStrRec.Value = "x" Then
                getIsPa = True
            End If
        End If
    End Function

End Class
