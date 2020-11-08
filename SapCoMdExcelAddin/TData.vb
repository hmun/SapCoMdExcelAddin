Public Class TData

    Public aTDataDic As Dictionary(Of String, TDataRec)
    Private aPar As SAPCommon.TStr

    Public Sub New(ByRef pPar As SAPCommon.TStr)
        aTDataDic = New Dictionary(Of String, TDataRec)
        aPar = pPar
    End Sub

    Public Sub addValue(pKey As String, pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String,
                        Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set")
        Dim aTDataRec As TDataRec
        If aTDataDic.ContainsKey(pKey) Then
            aTDataRec = aTDataDic(pKey)
            aTDataRec.setValues(pNAME, pVALUE, pCURRENCY, pFORMAT, pEmty, pEmptyChar, pOperation)
        Else
            aTDataRec = New TDataRec
            aTDataRec.setValues(pNAME, pVALUE, pCURRENCY, pFORMAT, pEmty, pEmptyChar, pOperation)
            aTDataDic.Add(pKey, aTDataRec)
        End If
    End Sub

    Public Sub addValue(pKey As String, pTStrRec As SAPCommon.TStrRec,
                        Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set",
                        Optional pNewStrucname As String = "")
        Dim aTDataRec As TDataRec
        Dim aName As String
        If pNewStrucname <> "" Then
            aName = pNewStrucname & "-" & pTStrRec.Fieldname
        Else
            aName = pTStrRec.Strucname & "-" & pTStrRec.Fieldname
        End If
        If aTDataDic.ContainsKey(pKey) Then
            aTDataRec = aTDataDic(pKey)
            aTDataRec.setValues(aName, pTStrRec.Value, pTStrRec.Currency, pTStrRec.Format, pEmty, pEmptyChar, pOperation)
        Else
            aTDataRec = New TDataRec
            aTDataRec.setValues(aName, pTStrRec.Value, pTStrRec.Currency, pTStrRec.Format, pEmty, pEmptyChar, pOperation)
            aTDataDic.Add(pKey, aTDataRec)
        End If
    End Sub

    Public Sub delData(pKey As String)
        aTDataDic.Remove(pKey)
    End Sub

    Public Function getPostingRecord() As TDataRec
        Dim aTDataRec As TDataRec = Nothing
        Dim aKvb As KeyValuePair(Of String, TDataRec)
        For Each aKvb In aTDataDic
            aTDataRec = aKvb.Value
            If aTDataRec.getPost(aPar) <> "" Then
                getPostingRecord = aTDataRec
                Exit Function
            End If
        Next
        getPostingRecord = Nothing
    End Function

    Public Function getFirstRecord() As TDataRec
        Dim aTDataRec As TDataRec = Nothing
        Dim aKvb As KeyValuePair(Of String, TDataRec)
        aKvb = aTDataDic.ElementAt(0)
        getFirstRecord = Nothing
        If Not IsNothing(aKvb) Then
            getFirstRecord = aKvb.Value
        End If
    End Function

End Class
