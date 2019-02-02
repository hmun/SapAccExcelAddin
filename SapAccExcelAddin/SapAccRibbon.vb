' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports Microsoft.Office.Tools.Ribbon
Imports SAP.Middleware.Connector

Public Class SapAccRibbon
    Private aSapCon
    Private aSapGeneral
    Const CP = 48 'column of post indicator
    Const CD = 49 'column of first header value
    Const CM = 58 'column of return message

    Private Sub ButtonCheckAccDoc_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCheckAccDoc.Click
        Dim aSapConRet As Integer
        If Not aSapGeneral.checkVersion() Then
            Exit Sub
        End If
        aSapConRet = 0
        If aSapCon Is Nothing Then
            aSapCon = New SapCon
        End If
        aSapConRet = aSapCon.checkCon()
        If aSapConRet = 0 Then
            SAP_AccDoc_execute(pTest:=True)
        Else
            aSapCon = Nothing
        End If
    End Sub
    Private Sub ButtonPostAccDoc_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostAccDoc.Click
        Dim aSapConRet As Integer
        If Not aSapGeneral.checkVersion() Then
            Exit Sub
        End If
        aSapConRet = 0
        If aSapCon Is Nothing Then
            aSapCon = New SapCon
        End If
        aSapConRet = aSapCon.checkCon()
        If aSapConRet = 0 Then
            SAP_AccDoc_execute(pTest:=False)
        Else
            aSapCon = Nothing
        End If
    End Sub

    Private Sub ButtonLogoff_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogoff.Click
        If Not aSapCon Is Nothing Then
            aSapCon.SAPlogoff()
            aSapCon = Nothing
        End If
    End Sub

    Private Sub ButtonLogon_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogon.Click
        Dim aConRet As Integer

        If Not aSapGeneral.checkVersion() Then
            Exit Sub
        End If
        If aSapCon Is Nothing Then
            aSapCon = New SapCon
        End If
        aConRet = aSapCon.checkCon()
        If aConRet = 0 Then
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap Accounting")
        Else
            aSapCon = Nothing
        End If
    End Sub


    Private Sub SapAccRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        aSapGeneral = New SapGeneral
    End Sub

    Private Sub SAP_AccDoc_execute(pTest As Boolean)
        Dim aSAPAcctngDocument As New SAPAcctngDocument(aSapCon)
        Dim aSAPDocItem As New SAPDocItem
        Dim aData As New Collection

        Dim i As Integer
        Dim aRetStr As String

        Dim aCURRTYP2 As String
        Dim aWAERS2 As String
        Dim aCURRTYP3 As String
        Dim aWAERS3 As String
        Dim aCURRTYP4 As String
        Dim aWAERS4 As String

        Dim adBLDAT As Date
        Dim adBLART As String
        Dim adBUKRS As String
        Dim adBUDAT As Date
        Dim adWAERS As String
        Dim adXBLNR As String
        Dim adBKTXT As String
        Dim adFIS_PERIOD As Integer
        Dim adACC_PRINCIPLE As String

        Dim aBLDAT As Date
        Dim aBLART As String
        Dim aBUKRS As String
        Dim aBUDAT As Date
        Dim aWAERS As String
        Dim aXBLNR As String
        Dim aBKTXT As String
        Dim aACC_PRINCIPLE As String
        Dim aFIS_PERIOD As Integer

        Dim aKONTO As String
        Dim aBETRA As Double

        Dim aSGTXT As String
        Dim aMWSKZ As String

        Dim aMATNR As String
        Dim aWERKS As String
        Dim aLIFNR As String
        Dim aKOSTL As String
        Dim aAUFNR As String

        Dim aFKBERNAME As String

        Dim aDws As Excel.Worksheet
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        aWB = Globals.SapAccAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP Accounting Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            Exit Sub
        End Try
        If CStr(aPws.Cells(2, 2).Value) <> "" Then
            adBUDAT = CDate(aPws.Cells(2, 2).Value)
        End If
        If CStr(aPws.Cells(3, 2).Value) <> "" Then
            adBLDAT = CDate(aPws.Cells(3, 2).Value)
        End If
        adXBLNR = aPws.Cells(4, 2).Value
        adBKTXT = aPws.Cells(5, 2).Value
        adBUKRS = aPws.Cells(6, 2).Value
        adWAERS = aPws.Cells(7, 2).Value
        adBLART = aPws.Cells(8, 2).Value
        adFIS_PERIOD = aPws.Cells(9, 2).Value
        adACC_PRINCIPLE = aPws.Cells(10, 2).Value

        aCURRTYP2 = aPws.Cells(11, 2).Value
        aWAERS2 = aPws.Cells(12, 2).Value
        aCURRTYP3 = aPws.Cells(13, 2).Value
        aWAERS3 = aPws.Cells(14, 2).Value
        aCURRTYP4 = aPws.Cells(15, 2).Value
        aWAERS4 = aPws.Cells(16, 2).Value

        aFKBERNAME = aPws.Cells(18, 2).Value
        If CStr(aFKBERNAME) = "" Then
            aFKBERNAME = "FKBER"
        End If
        ' Check Authority
        Dim aSAPZFI_CHECK_F_BKPF_BUK As New SAPZFI_CHECK_F_BKPF_BUK(aSapCon)
        Dim aAuth As Integer
        ' Read the Data
        Try
            aDws = aWB.Worksheets("SAP-Acc-Data")
        Catch Exc As System.Exception
            MsgBox("No SAP-Acc-Data Sheet in current workbook. Check if the current workbook is a valid SAP Accounting Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            Exit Sub
        End Try
        aDws.Activate()
        i = 2
        Do
            aKONTO = CStr(aDws.Cells(i, 2).Value)
            aMATNR = CStr(aDws.Cells(i, 3).Value)
            aWERKS = CStr(aDws.Cells(i, 4).Value)
            aLIFNR = CStr(aDws.Cells(i, 5).Value)
            aKOSTL = CStr(aDws.Cells(i, 7).Value)
            aAUFNR = CStr(aDws.Cells(i, 8).Value)
            aSGTXT = CStr(aDws.Cells(i, 38).Value)
            aMWSKZ = CStr(aDws.Cells(i, 39).Value)
            aBETRA = CDbl(aDws.Cells(i, 42).Value)

            aSAPDocItem = aSAPDocItem.create(CStr(aDws.Cells(i, 1).Value), aKONTO, aBETRA, aMWSKZ, aSGTXT, aAUFNR, aMATNR, aWERKS, aKOSTL, aLIFNR,
                                            CStr(aDws.Cells(i, 14).Value), CStr(aDws.Cells(i, 15).Value), CStr(aDws.Cells(i, 16).Value), CStr(aDws.Cells(i, 17).Value),
                                            CStr(aDws.Cells(i, 18).Value), CStr(aDws.Cells(i, 19).Value), CStr(aDws.Cells(i, 21).Value),
                                            CStr(aDws.Cells(i, 30).Value), CStr(aDws.Cells(i, 34).Value), CStr(aDws.Cells(i, 41).Value),
                                            CDbl(aDws.Cells(i, 43).Value), aCURRTYP2, aWAERS2,
                                            CDbl(aDws.Cells(i, 44).Value), aCURRTYP3, aWAERS3,
                                            CDbl(aDws.Cells(i, 45).Value), aCURRTYP4, aWAERS4,
                                            CStr(aDws.Cells(i, 9).Value), CStr(aDws.Cells(i, 23).Value), CStr(aDws.Cells(i, 6).Value),
                                            CStr(aDws.Cells(i, 25).Value), CStr(aDws.Cells(i, 26).Value), CStr(aDws.Cells(i, 46).Value),
                                            CStr(aDws.Cells(i, 10).Value), CStr(aDws.Cells(i, 11).Value), CStr(aDws.Cells(i, CD + 4).Value), CStr(aDws.Cells(i, 24).Value),
                                            CStr(aDws.Cells(i, 22).Value), CStr(aDws.Cells(i, 27).Value), CStr(aDws.Cells(i, 28).Value), CStr(aDws.Cells(i, 29).Value), CStr(aDws.Cells(i, 20).Value),
                                            CStr(aDws.Cells(i, 47).Value), CStr(aDws.Cells(i, 35).Value), CStr(aDws.Cells(i, 36).Value),
                                            CStr(aDws.Cells(i, 40).Value), CStr(aDws.Cells(i, 30).Value), CStr(aDws.Cells(i, 31).Value), CStr(aDws.Cells(i, 32).Value),
                                            CStr(aDws.Cells(i, 12).Value), CStr(aDws.Cells(i, 13).Value), CStr(aDws.Cells(i, 37).Value))
            aData.Add(aSAPDocItem)
            If (aDws.Cells(i, CP).Value = "X" Or aDws.Cells(i, CP).Value = "x") Then
                If InStr(1, aDws.Cells(i, CM).Value, "BKPFF") = 0 Then
                    If CStr(aDws.Cells(i, CD).Value) <> "" Then
                        aBUDAT = CDate(aDws.Cells(i, CD).Value)
                    Else
                        aBUDAT = adBUDAT
                    End If
                    If CStr(aDws.Cells(i, CD + 1).Value) <> "" Then
                        aBLDAT = CDate(aDws.Cells(i, CD + 1).Value)
                    Else
                        aBLDAT = adBLDAT
                    End If
                    If CStr(aDws.Cells(i, CD + 2).Value) <> "" Then
                        aXBLNR = CStr(aDws.Cells(i, CD + 2).Value)
                    Else
                        aXBLNR = adXBLNR
                    End If
                    If CStr(aDws.Cells(i, CD + 3).Value) <> "" Then
                        aBKTXT = CStr(aDws.Cells(i, CD + 3).Value)
                    Else
                        aBKTXT = adBKTXT
                    End If
                    If CStr(aDws.Cells(i, CD + 4).Value) <> "" Then
                        aBUKRS = CStr(aDws.Cells(i, CD + 4).Value)
                    Else
                        aBUKRS = adBUKRS
                    End If
                    If CStr(aDws.Cells(i, CD + 5).Value) <> "" Then
                        aWAERS = CStr(aDws.Cells(i, CD + 5).Value)
                    Else
                        aWAERS = adWAERS
                    End If
                    If CStr(aDws.Cells(i, CD + 6).Value) <> "" Then
                        aBLART = CStr(aDws.Cells(i, CD + 6).Value)
                    Else
                        aBLART = adBLART
                    End If
                    If CStr(aDws.Cells(i, CD + 7).Value) <> "" Then
                        aFIS_PERIOD = CStr(aDws.Cells(i, CD + 7).Value)
                    Else
                        aFIS_PERIOD = adFIS_PERIOD
                    End If
                    If CStr(aDws.Cells(i, CD + 8).Value) <> "" Then
                        aACC_PRINCIPLE = CStr(aDws.Cells(i, CD + 8).Value)
                    Else
                        aACC_PRINCIPLE = adACC_PRINCIPLE
                    End If
                    If InStr(1, CStr(aDws.Cells(i, CM).Value), "BKPFF") = 0 Then
                        aAuth = aSAPZFI_CHECK_F_BKPF_BUK.checkAuthority(adBUKRS)
                        If aAuth <> 2 Then
                            aDws.Cells(i, CM) = "User not authorized for company code " & aBUKRS
                        Else
                            aRetStr = aSAPAcctngDocument.post(aBLDAT, aBLART, aBUKRS, aBUDAT, aWAERS, aXBLNR, aBKTXT, aFIS_PERIOD, aACC_PRINCIPLE, aData, pTest, aFKBERNAME)
                            aDws.Cells(i, CM) = CStr(aRetStr)
                            aDws.Cells(i, CM + 1) = CStr(ExtractDocNumberFromMessage(aRetStr))
                        End If
                    End If
                End If
                aDws.Cells(i, CM + 1) = CStr(ExtractDocNumberFromMessage(aDws.Cells(i, CM).Value))
                aData = New Collection
            End If
            i = i + 1
        Loop While CStr(aDws.Cells(i, 1).value) <> ""
    End Sub

    Private Function ExtractDocNumberFromMessage(Message As String) As String
        Dim aPos As Integer
        Dim aLen As Long

        aLen = Len(Message)
        aPos = InStr(1, Message, "BKPFF")
        If aPos <> 0 Then
            ExtractDocNumberFromMessage = Mid(Message, aPos + 6, 18)
        Else
            ExtractDocNumberFromMessage = ""
        End If
    End Function

End Class
