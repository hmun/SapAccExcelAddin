' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports Microsoft.Office.Tools.Ribbon
Imports System.Configuration
Imports System.Collections.Specialized

Public Class SapAccRibbon
    Private aSapCon
    Private aSapGeneral
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Const CP = 48 'column of post indicator
    Const CD = 49 'column of first header value
    Const CM = 58 'column of return message

    Const ColNew = 27 ' column of the new accounting objects for invoice reposting

    Private Sub ButtonCheckAccDoc_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCheckAccDoc.Click
        Dim aSapConRet As Integer
        Dim aSapVersionRet As Integer
        log.Debug("ButtonCheckAccDoc_Click - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            Exit Sub
        End If
        log.Debug("ButtonCheckAccDoc_Click - " & "checking Connection")
        aSapConRet = 0
        If aSapCon Is Nothing Then
            aSapCon = New SapCon()
        End If
        aSapConRet = aSapCon.checkCon()
        If aSapConRet = 0 Then
            log.Debug("ButtonCheckAccDoc_Click - " & "checking version in SAP")
            Try
                aSapVersionRet = aSapGeneral.checkVersionInSAP(aSapCon)
            Catch ex As SystemException
                log.Warn("ButtonPostAccDoc_Click-checkVersionInSAP - )" & ex.ToString)
            End Try
            log.Debug("ButtonPostAccDoc_Click - " & "aSapVersionRet=" & CStr(aSapVersionRet))
            If aSapVersionRet = True Then
                log.Debug("ButtonCheckAccDoc_Click - " & "calling SAP_AccDoc_execute")
                SAP_AccDoc_execute(pTest:=True)
            End If
        Else
            log.Debug("ButtonCheckAccDoc_Click - " & "connection check failed")
            aSapCon = Nothing
        End If
    End Sub
    Private Sub ButtonPostAccDoc_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostAccDoc.Click
        Dim aSapConRet As Integer
        Dim aSapVersionRet As Integer

        log.Debug("ButtonPostAccDoc_Click - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            Exit Sub
        End If
        log.Debug("ButtonPostAccDoc_Click - " & "checking Connection")
        aSapConRet = 0
        If aSapCon Is Nothing Then
            aSapCon = New SapCon()
        End If
        aSapConRet = aSapCon.checkCon()
        If aSapConRet = 0 Then
            log.Debug("ButtonPostAccDoc_Click - " & "checking version in SAP")
            Try
                aSapVersionRet = aSapGeneral.checkVersionInSAP(aSapCon)
            Catch ex As SystemException
                log.Warn("ButtonPostAccDoc_Click-checkVersionInSAP - )" & ex.ToString)
            End Try
            log.Debug("ButtonPostAccDoc_Click - " & "aSapVersionRet=" & CStr(aSapVersionRet))
            If aSapVersionRet = True Then
                log.Debug("ButtonPostAccDoc_Click - " & "calling SAP_AccDoc_execute")
                SAP_AccDoc_execute(pTest:=False)
            End If
        Else
            log.Debug("ButtonPostAccDoc_Click - " & "connection check failed")
            aSapCon = Nothing
        End If
    End Sub

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


    Private Sub SapAccRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Dim sAll As NameValueCollection
        Dim s As String
        Dim enableInvoiceReposting As Boolean = False
        aSapGeneral = New SapGeneral
        Try
            sAll = ConfigurationManager.AppSettings
            s = sAll("enableInvoiceReposting")
            enableInvoiceReposting = Convert.ToBoolean(s)

        Catch Exc As System.Exception
            log.Error("SapAccRibbon_Load - " & "Exception=" & Exc.ToString)
        End Try
        If Not enableInvoiceReposting Then
            Globals.Ribbons.SapAccRibbon.Invoice.Visible = False
        Else
            Globals.Ribbons.SapAccRibbon.Invoice.Visible = True
        End If
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
        Dim aTRANS_DATE As Date?

        Dim aKONTO As String
        Dim aBETRA As Double
        Dim aTaxAmountTest As String
        Dim aTaxAmount As Double
        Dim aZZBBKNG As String
        Dim aZZBBTCO As String
        Dim aGL_ACCOUNT As String

        Dim aSGTXT As String
        Dim aMWSKZ As String

        Dim aMATNR As String
        Dim aWERKS As String
        Dim aLIFNR As String
        Dim aKOSTL As String
        Dim aAUFNR As String

        Dim aFKBERNAME As String
        Dim aTransdateOffset As Integer? = Nothing

        Dim aDws As Excel.Worksheet
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        log.Debug("SAP_AccDoc_execute - " & "reading Parameter")
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
        If CStr(aPws.Cells(19, 2).Value) <> "" Then
            Try
                aTransdateOffset = CInt(aPws.Cells(19, 2).Value)
            Catch
            End Try
        End If
        ' Check Authority
        log.Debug("SAP_AccDoc_execute - " & "creating aSAPZFI_CHECK_F_BKPF_BUK")
        Dim aSAPZFI_CHECK_F_BKPF_BUK As New SAPZFI_CHECK_F_BKPF_BUK(aSapCon)
        Dim aAuth As Integer
        log.Debug("SAP_AccDoc_execute - " & "reading Data")
        ' Read the Data
        Try
            aDws = aWB.Worksheets("SAP-Acc-Data")
        Catch Exc As System.Exception
            MsgBox("No SAP-Acc-Data Sheet in current workbook. Check if the current workbook is a valid SAP Accounting Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            Exit Sub
        End Try

        aDws.Activate()
        ' process the data
        Try
            log.Debug("SAP_AccDoc_execute - " & "processing data - disabling events, screen update, cursor")
            Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapAccAddIn.Application.EnableEvents = False
            Globals.SapAccAddIn.Application.ScreenUpdating = False
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
                aTaxAmountTest = CStr(aDws.Cells(i, CM + 3).Value)
                aTaxAmount = 0
                If aTaxAmountTest <> "" Then
                    If IsNumeric(aTaxAmountTest) Then
                        aTaxAmount = CDbl(aTaxAmountTest)
                    End If
                End If
                aZZBBKNG = CStr(aDws.Cells(i, CM + 4).Value)
                aZZBBTCO = CStr(aDws.Cells(i, CM + 5).Value)
                aGL_ACCOUNT = CStr(aDws.Cells(i, CM + 6).Value)

                aSAPDocItem = aSAPDocItem.create(CStr(aDws.Cells(i, 1).Value), aKONTO, aBETRA, aMWSKZ, aSGTXT, aAUFNR, aMATNR, aWERKS, aKOSTL, aLIFNR,
                                            CStr(aDws.Cells(i, 14).Value), CStr(aDws.Cells(i, 15).Value), CStr(aDws.Cells(i, 16).Value), CStr(aDws.Cells(i, 17).Value),
                                            CStr(aDws.Cells(i, 18).Value), CStr(aDws.Cells(i, 19).Value), CStr(aDws.Cells(i, 21).Value),
                                            CStr(aDws.Cells(i, 33).Value), CStr(aDws.Cells(i, 34).Value), CStr(aDws.Cells(i, 41).Value),
                                            CDbl(aDws.Cells(i, 43).Value), aCURRTYP2, aWAERS2,
                                            CDbl(aDws.Cells(i, 44).Value), aCURRTYP3, aWAERS3,
                                            CDbl(aDws.Cells(i, 45).Value), aCURRTYP4, aWAERS4,
                                            CStr(aDws.Cells(i, 9).Value), CStr(aDws.Cells(i, 23).Value), CStr(aDws.Cells(i, 6).Value),
                                            CStr(aDws.Cells(i, 25).Value), CStr(aDws.Cells(i, 26).Value), CStr(aDws.Cells(i, 46).Value),
                                            CStr(aDws.Cells(i, 10).Value), CStr(aDws.Cells(i, 11).Value), CStr(aDws.Cells(i, CD + 4).Value), CStr(aDws.Cells(i, 24).Value),
                                            CStr(aDws.Cells(i, 22).Value), CStr(aDws.Cells(i, 27).Value), CStr(aDws.Cells(i, 28).Value), CStr(aDws.Cells(i, 29).Value), CStr(aDws.Cells(i, 20).Value),
                                            CStr(aDws.Cells(i, 47).Value), CStr(aDws.Cells(i, 35).Value), CStr(aDws.Cells(i, 36).Value),
                                            CStr(aDws.Cells(i, 40).Value), CStr(aDws.Cells(i, 30).Value), CStr(aDws.Cells(i, 31).Value), CStr(aDws.Cells(i, 32).Value),
                                            CStr(aDws.Cells(i, 12).Value), CStr(aDws.Cells(i, 13).Value), CStr(aDws.Cells(i, 37).Value),
                                            aTaxAmount, aZZBBKNG, aZZBBTCO, aGL_ACCOUNT)
                aData.Add(aSAPDocItem)
                If (aDws.Cells(i, CP).Value = "X" Or aDws.Cells(i, CP).Value = "x") Then
                    log.Debug("SAP_AccDoc_execute - " & "found posting indicator, aData.Count=" & CStr(aData.Count))
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
                        If Not aTransdateOffset Is Nothing Then
                            If CStr(aDws.Cells(i, CD + aTransdateOffset).Value) <> "" Then
                                Try
                                    aTRANS_DATE = CDate(aDws.Cells(i, CD + aTransdateOffset).Value)
                                Catch
                                    aTRANS_DATE = Nothing
                                End Try
                            Else
                                aTRANS_DATE = Nothing
                            End If
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
                            log.Debug("SAP_AccDoc_execute - " & "calling aSAPZFI_CHECK_F_BKPF_BUK.checkAuthority(aBUKRS), aBUKRS=" & aBUKRS)
                            aAuth = aSAPZFI_CHECK_F_BKPF_BUK.checkAuthority(aBUKRS)
                            log.Debug("SAP_AccDoc_execute - " & "aAuth=" & CStr(aAuth))
                            If aAuth <> 2 Then
                                log.Debug("SAP_AccDoc_execute - " & "User not authorized for company code " & aBUKRS)
                                aDws.Cells(i, CM) = "User not authorized for company code " & aBUKRS
                            Else
                                log.Debug("SAP_AccDoc_execute - " & "calling aSAPAcctngDocument.post, pTest=" & CStr(pTest))
                                aRetStr = aSAPAcctngDocument.post(aBLDAT, aBLART, aBUKRS, aBUDAT, aWAERS, aXBLNR, aBKTXT, aFIS_PERIOD, aACC_PRINCIPLE, aData, pTest, aFKBERNAME, aTRANS_DATE)
                                log.Debug("SAP_AccDoc_execute - " & "aSAPAcctngDocument.post returned, aRetStr=" & aRetStr)
                                aDws.Cells(i, CM) = CStr(aRetStr)
                                aDws.Cells(i, CM + 1) = CStr(ExtractDocNumberFromMessage(aRetStr))
                            End If
                        End If
                    End If
                    aDws.Cells(i, CM + 1) = CStr(ExtractDocNumberFromMessage(aDws.Cells(i, CM).Value))
                    log.Debug("SAP_AccDoc_execute - " & "ExtractDocNumberFromMessage=" & CStr(ExtractDocNumberFromMessage(aDws.Cells(i, CM).Value)))
                    log.Debug("SAP_AccDoc_execute - " & "starting new aData")
                    aData = New Collection
                End If
                i = i + 1
            Loop While CStr(aDws.Cells(i, 1).value) <> ""
            log.Debug("SAP_AccDoc_execute - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapAccAddIn.Application.EnableEvents = True
            Globals.SapAccAddIn.Application.ScreenUpdating = True
            Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapAccAddIn.Application.EnableEvents = True
            Globals.SapAccAddIn.Application.ScreenUpdating = True
            Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SAP_AccDoc_execute failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            log.Error("SAP_AccDoc_execute - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
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

    Private Sub ButtonReadInvoices_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadInvoices.Click
        Dim aSapConRet As Integer
        Dim aSapVersionRet As Integer

        log.Debug("ButtonReadInvoices_Click - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            Exit Sub
        End If
        log.Debug("ButtonReadInvoices_Click - " & "checking Connection")
        aSapConRet = 0
        If aSapCon Is Nothing Then
            aSapCon = New SapCon()
        End If
        aSapConRet = aSapCon.checkCon()
        If aSapConRet = 0 Then
            log.Debug("ButtonReadInvoices_Click - " & "checking version in SAP")
            Try
                aSapVersionRet = aSapGeneral.checkVersionInSAP(aSapCon)
            Catch ex As SystemException
                log.Warn("ButtonReadInvoices_Click-checkVersionInSAP - )" & ex.ToString)
            End Try
            log.Debug("ButtonReadInvoices_Click - " & "aSapVersionRet=" & CStr(aSapVersionRet))
            If aSapVersionRet = True Then
                log.Debug("ButtonReadInvoices_Click - " & "calling ReadInvoices_exec")
                ReadInvoices_exec()
            End If
        Else
            log.Debug("ButtonReadInvoices_Click - " & "connection check failed")
            aSapCon = Nothing
        End If
    End Sub

    Sub ReadInvoices_exec()
        Dim aSAPIncomingInvoice As SAPIncomingInvoice
        Try
            aSAPIncomingInvoice = New SAPIncomingInvoice(aSapCon)
        Catch Exc As System.Exception
            log.Warn("ButtonReadInvoices_Click - " & Exc.ToString)
            Exit Sub
        End Try
        Dim aSAPFormat As New SAPFormat
        Dim aILWS As Excel.Worksheet
        Dim aIDWS As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer
        log.Debug("ButtonReadInvoices_Click - " & "InvoiceList Sheet")
        aWB = Globals.SapAccAddIn.Application.ActiveWorkbook
        Try
            aILWS = aWB.Worksheets("InvoiceList")
            aILWS.Activate()
        Catch Exc As System.Exception
            log.Warn("ButtonReadInvoices_Click - " & "No InvoiceList Sheet in current workbook.")
            MsgBox("No InvoiceList Sheet in current workbook. Check if the current workbook is a valid SAP Invoice Reposting Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            Exit Sub
        End Try
        log.Debug("ButtonReadInvoices_Click - " & "InvoiceData Sheet")
        Try
            aIDWS = aWB.Worksheets("InvoiceData")
        Catch Exc As System.Exception
            log.Warn("ButtonReadInvoices_Click - " & "No InvoiceData Sheet in current workbook.")
            MsgBox("No InvoiceData Sheet in current workbook. Check if the current workbook is a valid SAP Invoice Reposting Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            Exit Sub
        End Try
        Try
            Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapAccAddIn.Application.EnableEvents = False
            Globals.SapAccAddIn.Application.ScreenUpdating = False
            log.Debug("ButtonReadInvoices_Click - " & "processing data - disabling events, screen update, cursor")
            ' clear the InvoiceData
            Dim aRange As Excel.Range
            If CStr(aIDWS.Cells(2, 1).Value) <> "" Then
                aRange = aIDWS.Range("A2")
                i = 2
                Do
                    i += 1
                Loop While CStr(aIDWS.Cells(i, 1).Value) <> ""
                aRange = aIDWS.Range(aRange, aIDWS.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            ' read the invoices
            Dim IDCol As New Collection
            Dim IDColTmp As New Dictionary(Of String, BInvRec)
            Dim iL As Integer = 2
            Do
                IDColTmp = New Dictionary(Of String, BInvRec)
                IDColTmp = aSAPIncomingInvoice.getDetail(aSAPFormat.unpack(aILWS.Cells(iL, 2).value, 10), CStr(aILWS.Cells(iL, 1).value))
                If Not IDColTmp Is Nothing Then
                    For Each aBInvRec In IDColTmp.Values
                        IDCol.Add(aBInvRec)
                    Next
                End If
                iL += 1
            Loop While CStr(aILWS.Cells(iL, 1).Value) <> ""
            ' write the invoice data

            Dim iA As Integer = 2
            Dim aCells As Excel.Range
            For Each aBInvRec In IDCol
                If aBInvRec.aSERIAL_NO.Value <> "" Then
                    Dim aRetArray As Object = aBInvRec.toStringValue()
                    aCells = aIDWS.Range(aIDWS.Cells(iA, 1), aIDWS.Cells(iA, aRetArray.Length - 1))
                    aCells.Value = aRetArray
                    '    aIDWS.Cells(iA, 10).Value = CDbl(aBInvRec.aEXCH_RATE.Value)
                    aIDWS.Cells(iA, 5).Value = CDate(aBInvRec.aDOC_DATE.Value)
                    aIDWS.Cells(iA, 6).Value = CDate(aBInvRec.aPSTNG_DATE.Value)
                    If CStr(aBInvRec.aITEM_AMOUNT.Value) = "" Then
                        aIDWS.Cells(iA, 18).Value = ""
                    Else
                        aIDWS.Cells(iA, 18).Value = CDbl(aBInvRec.aITEM_AMOUNT.Value)
                    End If
                    If CStr(aBInvRec.aQUANTITY.Value) = "" Then
                        aIDWS.Cells(iA, 19).Value = ""
                    Else
                        aIDWS.Cells(iA, 19).Value = CDbl(aBInvRec.aQUANTITY.Value)
                    End If
                    iA += 1
                End If
            Next
            Globals.SapAccAddIn.Application.EnableEvents = True
            Globals.SapAccAddIn.Application.ScreenUpdating = True
            Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            aIDWS.Activate()
        Catch ex As System.Exception
            Globals.SapAccAddIn.Application.EnableEvents = True
            Globals.SapAccAddIn.Application.ScreenUpdating = True
            Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("ButtonReadInvoices_Click failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            log.Error("ButtonReadInvoices_Click - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

    End Sub

    Private Sub ButtonGenGLData_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonGenGLData.Click
        Dim aMigHelper As MigHelper
        Dim aBasis As New Collection
        Dim aBasisLine As New Dictionary(Of String, SAPCommon.TField)
        Dim aPostingLine As New Dictionary(Of String, SAPCommon.TField)
        Dim aContraLine As New Dictionary(Of String, SAPCommon.TField)
        Dim aPostingLines As Collection
        Dim aPWs As Excel.Worksheet
        Dim aIWs As Excel.Worksheet
        Dim aDWs As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aUselocal As Boolean = False
        Dim aUseBasis As Boolean = False
        Dim i As Integer

        log.Debug("ButtonGenGLData_Click - " & "Invoice Sheet")
        aWB = Globals.SapAccAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
            If CStr(aPWs.Cells(17, 2).Value) = "X" Then
                aUselocal = True
            End If
        Catch Exc As System.Exception
            log.Debug("ButtonGenGLData_Click - " & "No Parameter Sheet in current workbook. -> aDoGeneral = True")
        End Try
        Try
            aIWs = aWB.Worksheets("InvoiceData")
        Catch Ex As System.Exception
            Try
                aIWs = aWB.Worksheets("Basis")
                aUseBasis = True
            Catch Exc As System.Exception
                log.Warn("ButtonGenGLData_Click - " & "No InvoiceData or Basis Sheet in current workbook.")
                MsgBox("No InvoiceData Sheet or Basis Sheet in current workbook. Check if the current workbook is a valid Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
                Exit Sub
            End Try
        End Try
        aIWs.Activate()
        If aUselocal Then
            aMigHelper = New MigHelper(uselocal:=True)
        Else
            aMigHelper = New MigHelper(uselocal:=False)
        End If
        ' process the data
        Try
            log.Debug("ButtonGenGLData_Click - " & "processing data - disabling events, screen update, cursor")
            Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapAccAddIn.Application.EnableEvents = False
            Globals.SapAccAddIn.Application.ScreenUpdating = False
            ' read the invoice reposting lines
            i = 2
            Do
                If aUseBasis _
                    Or CStr(aIWs.Cells(i, ColNew).value) <> "" Or CStr(aIWs.Cells(i, ColNew + 1).value) <> "" Or CStr(aIWs.Cells(i, ColNew + 2).value) <> "" _
                    Or CStr(aIWs.Cells(i, ColNew + 3).value) <> "" Or CStr(aIWs.Cells(i, ColNew + 4).value) <> "" Or CStr(aIWs.Cells(i, ColNew + 5).value) <> "" Then
                    aBasisLine = aMigHelper.makeDictForRules(aIWs, i, 1, 1, 100)
                    aBasis.Add(aBasisLine)
                End If
                i = i + 1
            Loop While CStr(aIWs.Cells(i, 1).value) <> ""
            ' create the posting lines
            aPostingLines = New Collection
            For Each aBasisLine In aBasis
                aPostingLine = aMigHelper.mig.ApplyRules(aBasisLine, "P")
                aPostingLines.Add(aPostingLine)
                aContraLine = aMigHelper.mig.ApplyRules(aBasisLine, "C")
                ' some hard-coding here
                If aContraLine("Amount").Value <> "" Then
                    aContraLine("Amount").Value = CStr(CDbl(aContraLine("Amount").Value) * -1)
                Else
                    aContraLine("Amount").Value = ""
                End If
                aPostingLines.Add(aContraLine)
            Next
            'output the posting lines
            Try
                aDWs = aWB.Worksheets("SAP-Acc-Data")
            Catch Exc As System.Exception
                Globals.SapAccAddIn.Application.EnableEvents = True
                Globals.SapAccAddIn.Application.ScreenUpdating = True
                Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
                MsgBox("No SAP-Acc-Data Sheet in current workbook. Check if the current workbook is a valid SAP Accounting Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
                Exit Sub
            End Try
            Dim aRange As Excel.Range
            If CStr(aDWs.Cells(2, 1).Value) <> "" Then
                aRange = aDWs.Range("A2")
                i = 2
                Do
                    i += 1
                Loop While CStr(aDWs.Cells(i, 1).Value) <> ""
                aRange = aDWs.Range(aRange, aDWs.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            Dim aKey As String
            Dim aValue As String
            i = 2
            For Each aPostingLine In aPostingLines
                For j = 1 To CM + 10
                    If CStr(aDWs.Cells(1, j).Value) <> "" Then
                        aKey = CStr(aDWs.Cells(1, j).Value)
                        If aPostingLine.ContainsKey(aKey) Then
                            aValue = aPostingLine(aKey).Value
                            If aKey = "Amount" And aValue <> "" Then
                                aDWs.Cells(i, j).Value = CDbl(aValue)
                            ElseIf aPostingLine(aKey).FType = "F" Then
                                aDWs.Cells(i, j).FormulaR1C1 = "=" & CStr(aValue)
                            Else
                                aDWs.Cells(i, j).Value = aValue
                            End If
                        End If
                    End If
                Next j
                i += 1
            Next
            Globals.SapAccAddIn.Application.EnableEvents = True
            Globals.SapAccAddIn.Application.ScreenUpdating = True
            Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            aDWs.Activate()
        Catch ex As System.Exception
            Globals.SapAccAddIn.Application.EnableEvents = True
            Globals.SapAccAddIn.Application.ScreenUpdating = True
            Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("ButtonGenGLData_Click failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            log.Error("ButtonGenGLData_Click - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

        '        aMigHelper.saveToConfig()
    End Sub

End Class
