' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports Microsoft.Office.Tools.Ribbon
Imports SAP.Middleware.Connector

Public Class SapCoMdRibbon
    Private aSapCon
    Private aSapGeneral
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private aCoAre As String

    Private Function checkCon() As Integer
        Dim aSapConRet As Integer
        Dim aSapVersionRet As Integer
        checkCon = False
        log.Debug("checkCon - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            Exit Function
        End If
        log.Debug("checkCon - " & "checking Connection")
        aSapConRet = 0
        If aSapCon Is Nothing Then
            Try
                aSapCon = New SapCon()
            Catch ex As SystemException
                log.Warn("checkCon-New SapCon - )" & ex.ToString)
            End Try
        End If
        Try
            aSapConRet = aSapCon.checkCon()
        Catch ex As SystemException
            log.Warn("checkCon-aSapCon.checkCon - )" & ex.ToString)
        End Try
        If aSapConRet = 0 Then
            log.Debug("checkCon - " & "checking version in SAP")
            Try
                aSapVersionRet = aSapGeneral.checkVersionInSAP(aSapCon)
            Catch ex As SystemException
                log.Warn("checkCon - )" & ex.ToString)
            End Try
            log.Debug("checkCon - " & "aSapVersionRet=" & CStr(aSapVersionRet))
            If aSapVersionRet = True Then
                log.Debug("checkCon - " & "checkCon = True")
                checkCon = True
            Else
                log.Debug("checkCon - " & "connection check failed")
            End If
        End If
    End Function

    Private Sub ButtonLogoff_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogoff.Click
        log.Debug("ButtonLogoff_Click - " & "starting logoff")
        If Not aSapCon Is Nothing Then
            log.Debug("ButtonLogoff_Click - " & "calling aSapCon.SAPlogoff()")
            aSapCon.SAPlogoff()
            aSapCon = Nothing
        End If
        log.Debug("ButtonLogoff_Click - " & "exit")
    End Sub

    Private Sub ButtonLogon_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogon.Click
        Dim aConRet As Integer

        log.Debug("ButtonLogon_Click - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            log.Debug("ButtonLogon_Click - " & "Version check failed")
            Exit Sub
        End If
        log.Debug("ButtonLogon_Click - " & "creating SapCon")
        If aSapCon Is Nothing Then
            aSapCon = New SapCon()
        End If
        log.Debug("ButtonLogon_Click - " & "calling SapCon.checkCon()")
        aConRet = aSapCon.checkCon()
        If aConRet = 0 Then
            log.Debug("ButtonLogon_Click - " & "connection successfull")
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap Accounting")
        Else
            log.Debug("ButtonLogon_Click - " & "connection failed")
            aSapCon = Nothing
        End If
    End Sub

    Private Sub SapCoMdRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        aSapGeneral = New SapGeneral
    End Sub

    Private Sub ButtonCCCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCCCreate.Click
        Dim aSapCoMdRibbonCC As New SapCoMdRibbonCC
        If checkCon() = True Then
            aSapCoMdRibbonCC.exec(pSapCon:=aSapCon, pMode:="Create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonCCCreate_Click")
        End If
    End Sub

    Private Sub ButtonCCChange_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCCChange.Click
        Dim aSapCoMdRibbonCC As New SapCoMdRibbonCC
        If checkCon() = True Then
            aSapCoMdRibbonCC.exec(pSapCon:=aSapCon, pMode:="Change")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonCCChange_Click")
        End If
    End Sub

    Private Sub ButtonGLAccCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonGLAccCreate.Click
        Dim aSapCoMdRibbonGLAccount As New SapCoMdRibbonGLAccount
        If checkCon() = True Then
            aSapCoMdRibbonGLAccount.exec(pSapCon:=aSapCon, pMode:="Create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonGLAccCreate_Click")
        End If
    End Sub

    Private Sub ButtonCECreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCECreate.Click
        Dim aSapCoMdRibbonCE As New SapCoMdRibbonCE
        If checkCon() = True Then
            aSapCoMdRibbonCE.exec(pSapCon:=aSapCon, pMode:="Create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonCECreate_Click")
        End If
    End Sub

    Private Sub ButtonCEChange_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCEChange.Click
        Dim aSapCoMdRibbonCE As New SapCoMdRibbonCE
        If checkCon() = True Then
            aSapCoMdRibbonCE.exec(pSapCon:=aSapCon, pMode:="Change")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonCEChange_Click")
        End If
    End Sub
End Class