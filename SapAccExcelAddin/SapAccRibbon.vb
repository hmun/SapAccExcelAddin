' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports Microsoft.Office.Tools.Ribbon
Imports System.Configuration
Imports System.Collections.Specialized

Public Class SapAccRibbon
    Private aSapCon
    Private aSapGeneral
    Private aAccPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
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
            aSapCon = New SapConHelper()
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
                SAP_AccDoc_execute_dyn(pTest:=True)
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
            aSapCon = New SapConHelper()
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
                SAP_AccDoc_execute_dyn(pTest:=False)
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
            aSapCon = New SapConHelper()
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
        Dim enableCosReposting As Boolean = False
        aSapGeneral = New SapGeneral
        Try
            sAll = ConfigurationManager.AppSettings
            s = sAll("enableInvoiceReposting")
            enableInvoiceReposting = Convert.ToBoolean(s)
            s = sAll("enableCosReposting")
            enableCosReposting = Convert.ToBoolean(s)

        Catch Exc As System.Exception
            log.Error("SapAccRibbon_Load - " & "Exception=" & Exc.ToString)
        End Try
        If Not enableInvoiceReposting Then
            Globals.Ribbons.SapAccRibbon.Invoice.Visible = False
        Else
            Globals.Ribbons.SapAccRibbon.Invoice.Visible = True
        End If
        If Not enableCosReposting Then
            Globals.Ribbons.SapAccRibbon.Cos_Split.Visible = False
        Else
            Globals.Ribbons.SapAccRibbon.Cos_Split.Visible = True
        End If
    End Sub

    Private Function getAccParameters() As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim akey As String
        Dim aName As String
        Dim i As Integer

        log.Debug("getAccParameters - " & "reading Parameter")
        aWB = Globals.SapAccAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP Accounting Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            getAccParameters = False
            Exit Function
        End Try
        aName = "SAPAccDoc"
        akey = CStr(aPws.Cells(1, 1).Value)
        If akey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid SAP Accounting Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP PS")
            getAccParameters = False
            Exit Function
        End If
        i = 2
        aAccPar = New SAPCommon.TStr
        Do
            aAccPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 4).value), pFORMAT:=CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters for AccDoc - otherwise check here
        getAccParameters = True
    End Function

    Private Function getIntParameters() As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = Globals.SapAccAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SAP Accounting Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            getIntParameters = False
            Exit Function
        End Try
        i = 2
        aIntPar = New SAPCommon.TStr
        Do
            aIntPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters check - we should know what we are doing
        getIntParameters = True
    End Function

    Private Sub SAP_AccDoc_execute_dyn(pTest As Boolean)
        Dim aSAPAcctngDocument As New SAPAcctngDocument(aSapCon)
        ' get posting parameters
        If Not getAccParameters() Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters() Then
            Exit Sub
        End If

        ' Check Authority
        log.Debug("SAP_AccDoc_execute - " & "creating aSAPZFI_CHECK_F_BKPF_BUK")
        Dim aSAPZFI_CHECK_F_BKPF_BUK As New SAPZFI_CHECK_F_BKPF_BUK(aSapCon)
        Dim aAuth As Integer

        log.Debug("SAP_AccDoc_execute - " & "reading Data")
        ' Read the Data
        Dim aLOff As Integer = If(aIntPar.value("LOFF", "DATA") <> "", CInt(aIntPar.value("LOFF", "DATA")), 4)
        Dim aDwsName As String = If(aIntPar.value("WS", "DATA") <> "", aIntPar.value("WS", "DATA"), "Data")
        Dim aDws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        aWB = Globals.SapAccAddIn.Application.ActiveWorkbook
        log.Debug("SAP_AccDoc_execute - " & "reading Data")
        ' Read the Data
        Try
            aDws = aWB.Worksheets(aDwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Accounting Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            Exit Sub
        End Try

        aDws.Activate()
        Dim aAccItems As New TData(aIntPar)
        Dim aAccItem As New TDataRec
        Dim aKey As String
        Dim j As Integer
        Dim jMax As UInt64 = 0
        Dim aDocNr As UInt64 = 0
        Dim aDumpDocNr As UInt64 = If(aIntPar.value("DBG", "DUMPDOCNR") <> "", CInt(aIntPar.value("DBG", "DUMPDOCNR")), 0)
        Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aDocClmn As String = If(aIntPar.value("COL", "DATADOC") <> "", aIntPar.value("COL", "DATADOC"), "INT-DOC")
        Dim aDocClmnNr As Integer = 0
        Dim aCompCode As String
        Dim aRetStr As String
        Do
            jMax += 1
            If CStr(aDws.Cells(1, jMax).value) = aMsgClmn Then
                aMsgClmnNr = jMax
            ElseIf CStr(aDws.Cells(1, jMax).value) = aDocClmn Then
                aDocClmnNr = jMax
            End If
        Loop While CStr(aDws.Cells(aLOff - 3, jMax + 1).value) <> ""
        ' process the data
        Try
            log.Debug("SAP_AccDoc_execute - " & "processing data - disabling events, screen update, cursor")
            Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapAccAddIn.Application.EnableEvents = False
            Globals.SapAccAddIn.Application.ScreenUpdating = False
            Dim i As UInt64 = aLOff + 1
            Do
                aKey = CStr(i)
                For j = 1 To jMax
                    If CStr(aDws.Cells(1, j).value) <> "N/A" And CStr(aDws.Cells(1, j).value) <> "" And
                        CStr(aDws.Cells(1, j).value) <> aMsgClmn And CStr(aDws.Cells(1, j).value) <> aMsgClmn Then
                        aAccItems.addValue(aKey, CStr(aDws.Cells(aLOff - 3, j).value), CStr(aDws.Cells(i, j).value),
                                           CStr(aDws.Cells(aLOff - 2, j).value), CStr(aDws.Cells(aLOff - 1, j).value),
                                           pEmptyChar:="")
                    End If
                Next
                aAccItem = aAccItems.aTDataDic(aKey)
                ' if the posting indicator is set -> call the sap AC_DOCUMENT BAPI
                If aAccItem.getPost(aIntPar) <> "" Then
                    log.Debug("SAP_AccDoc_execute - " & "found posting indicator, aData.Count=" & CStr(aAccItems.aTDataDic.Count))
                    ' only convert and process the data of documents that have not been processed
                    aDocNr += 1
                    If InStr(1, aDws.Cells(i, aMsgClmnNr).Value, "BKPFF") = 0 Then
                        Dim aTSAP_DocData As New TSAP_DocData(aAccPar, aIntPar, aSAPAcctngDocument, pTest)
                        If aTSAP_DocData.fillHeader(aAccItems) And aTSAP_DocData.fillData(aAccItems) Then
                            ' check if we should dump this document
                            If aDocNr = aDumpDocNr Then
                                log.Debug("SAP_AccDoc_execute - " & "dumping Document Nr " & CStr(aDocNr))
                                aTSAP_DocData.dumpHeader()
                                aTSAP_DocData.dumpData()
                            End If
                            ' post the document here
                            log.Debug("SAP_AccDoc_execute - " & "filled Header and Data in aTSAP_DocData, aTSAP_DocData.aData.Count=" & CStr(aTSAP_DocData.aData.aTDataDic.Count))

                            log.Debug("SAP_AccDoc_execute - " & "checking mandatory header fields")
                            If Not aTSAP_DocData.checkHeader() Then
                                log.Debug("SAP_AccDoc_execute - " & "checking mandatory header fields - failed")
                                aDws.Cells(i, aMsgClmnNr) = "Error: Fill all mandatory header fields for document"
                            Else
                                aCompCode = aTSAP_DocData.getCompanyCode()
                                log.Debug("SAP_AccDoc_execute - " & "calling aSAPZFI_CHECK_F_BKPF_BUK.checkAuthority(aCompCode), aCompCode=" & aCompCode)
                                aAuth = aSAPZFI_CHECK_F_BKPF_BUK.checkAuthority(aCompCode)
                                log.Debug("SAP_AccDoc_execute - " & "aAuth=" & CStr(aAuth))
                                If aAuth <> 2 Then
                                    log.Warn("SAP_AccDoc_execute - " & "User not authorized for company code " & aCompCode)
                                    aDws.Cells(i, aMsgClmnNr) = "User not authorized for company code " & aCompCode
                                Else
                                    log.Debug("SAP_AccDoc_execute - " & "calling aSAPAcctngDocument.post, pTest=" & CStr(pTest))
                                    aRetStr = aSAPAcctngDocument.post(aTSAP_DocData, pTest:=pTest)
                                    log.Debug("SAP_AccDoc_execute - " & "aSAPAcctngDocument.post returned, aRetStr=" & aRetStr)
                                    aDws.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                                    aDws.Cells(i, aDocClmnNr) = CStr(ExtractDocNumberFromMessage(aRetStr))
                                End If
                            End If
                        Else
                            log.Warn("SAP_AccDoc_execute - " & "filling Header or Data in aTSAP_DocData failed!")
                            aDws.Cells(i, aMsgClmnNr) = "Filling Header or Data in aTSAP_DocData failed!"
                        End If
                    End If
                    aAccItems = New TData(aIntPar)
                End If
                i += 1
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
            aSapCon = New SapConHelper()
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
        ' get internal parameters
        If Not getIntParameters() Then
            log.Error("ButtonGenerate_Click getIntParameters - " & "failed - exit")
            Exit Sub
        End If
        ' get the ruleset limits
        Dim aGenNrFrom As Integer = If(aIntPar.value("GEN", "RULESET_FROM") <> "", CInt(aIntPar.value("GEN", "RULESET_FROM")), 0)
        Dim aGenNrTo As Integer = If(aIntPar.value("GEN", "RULESET_TO") <> "", CInt(aIntPar.value("GEN", "RULESET_TO")), 0)
        Dim aGenNr As String = ""
        For i As Integer = aGenNrFrom To aGenNrTo
            Dim aNr As String = If(i = 0, "", CStr(i))
            ButtonGenGLData_exec(pSapCon:=aSapCon, pNr:=aNr)
        Next
    End Sub


    Private Sub ButtonGenGLData_exec(ByRef pSapCon As SapConHelper, Optional pNr As String = "")
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

        ' get internal parameters
        If Not getIntParameters() Then
            Exit Sub
        End If
        Dim aDwsName As String = If(aIntPar.value("GEN" & pNr, "WS_DATA") <> "", aIntPar.value("GEN" & pNr, "WS_DATA"), "SAP-Acc-Data")
        Dim aBwsName As String = If(aIntPar.value("GEN" & pNr, "WS_BASE") <> "", aIntPar.value("GEN" & pNr, "WS_BASE"), "InvoiceData")
        Dim aEmptyChar As String = If(aIntPar.value("GEN" & pNr, "CHAR_EMPTY") <> "", aIntPar.value("GEN" & pNr, "CHAR_EMPTY"), "#")
        Dim aIgnoreEmpty As String = If(aIntPar.value("GEN" & pNr, "IGNORE_EMPTY") <> "", aIntPar.value("GEN" & pNr, "IGNORE_EMPTY"), "X")
        Dim aGenEmpty As Boolean = If(aIgnoreEmpty = "X", False, True)
        Dim aDeleteData As String = If(aIntPar.value("GEN" & pNr, "DELETE_DATA") <> "", aIntPar.value("GEN" & pNr, "DELETE_DATA"), "X")
        Dim aGenDeleteData As Boolean = If(aDeleteData = "X", True, False)
        Dim aLOff As Integer = If(aIntPar.value("GEN" & pNr, "LOFF_DATA") <> "", CInt(aIntPar.value("GEN" & pNr, "LOFF_DATA")), 4)
        Dim aLOffBData As Integer = If(aIntPar.value("GEN" & pNr, "LOFF_BDATA") <> "", CInt(aIntPar.value("GEN" & pNr, "LOFF_BDATA")), 1)
        Dim aLOffBNames As Integer = If(aIntPar.value("GEN" & pNr, "LOFF_BNAMES") <> "", CInt(aIntPar.value("GEN" & pNr, "LOFF_BNAMES")), 0)
        Dim aLOffTNames As Integer = If(aIntPar.value("GEN" & pNr, "LOFF_TNAMES") <> "", CInt(aIntPar.value("GEN" & pNr, "LOFF_TNAMES")), aLOff - 1)
        Dim aLineOut As Integer = If(aIntPar.value("GEN" & pNr, "LINE_OUT") <> "", CInt(aIntPar.value("GEN" & pNr, "LINE_OUT")), 0)
        Dim aBaseColFrom As Integer = If(aIntPar.value("GEN" & pNr, "BASE_COLFROM") <> "", CInt(aIntPar.value("GEN" & pNr, "BASE_COLFROM")), 1)
        Dim aBaseColTo As Integer = If(aIntPar.value("GEN" & pNr, "BASE_COLTO") <> "", CInt(aIntPar.value("GEN" & pNr, "BASE_COLTO")), 100)
        Dim aBaseFilter As String = If(aIntPar.value("GEN" & pNr, "BASE_FILTER") <> "", CStr(aIntPar.value("GEN" & pNr, "BASE_FILTER")), "")
        Dim aTargetFilter As String = If(aIntPar.value("GEN" & pNr, "TARGET_FILTER") <> "", CStr(aIntPar.value("GEN" & pNr, "TARGET_FILTER")), "")
        ' should we compress posting lines?
        Dim aGenCompData As String = If(aIntPar.value("GEN", "COMP_DATA") <> "", CStr(aIntPar.value("GEN", "COMP_DATA")), "")
        Dim aCompress As Boolean = If(aGenCompData = "X", True, False)
        ' should we suppress line with zero values
        Dim aGenSupprZero As String = If(aIntPar.value("GEN", "SUPPR_ZERO") <> "", CStr(aIntPar.value("GEN", "SUPPR_ZERO")), "")
        Dim aSupprZero As Boolean = If(aGenSupprZero = "X", True, False)

        log.Debug("ButtonGenGLData_Click - " & "Basis Sheet")
        aWB = Globals.SapAccAddIn.Application.ActiveWorkbook
        Dim aGenLocalRules As String = If(aIntPar.value("GEN", "LOCAL_RULES") <> "", CStr(aIntPar.value("GEN", "LOCAL_RULES")), "")
        If aGenLocalRules = "X" Then
            aUselocal = True
            log.Debug("ButtonGenGLData_Click - " & "aUselocal = True")
        Else
            ' Fallback for compatibilty to old templates
            Try
                aPWs = aWB.Worksheets("Parameter")
                If CStr(aPWs.Cells(17, 2).Value) = "X" Then
                    aUselocal = True
                    log.Debug("ButtonGenGLData_Click - " & "aUselocal = True")
                End If
            Catch Exc As System.Exception
                log.Debug("ButtonGenGLData_Click - " & "No Parameter Sheet in current workbook. -> aUselocal = False")
            End Try
        End If
        Try
            aIWs = aWB.Worksheets("InvoiceData")
        Catch Ex As System.Exception
            Try
                aIWs = aWB.Worksheets(aBwsName)
                aUseBasis = True
            Catch Exc As System.Exception
                log.Warn("ButtonGenGLData_Click - " & "No InvoiceData or " & aBwsName & " in current workbook.")
                MsgBox("No InvoiceData Sheet or " & aBwsName & " in current workbook. Check if the current workbook is a valid Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
                Exit Sub
            End Try
        End Try
        aIWs.Activate()
        aMigHelper = New MigHelper(pPar:=aIntPar, pNr:=pNr, pFilterStr:=aBaseFilter, pUselocal:=aUselocal)
        ' process the data
        Try
            log.Debug("ButtonGenGLData_Click - " & "processing data - disabling events, screen update, cursor")
            Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapAccAddIn.Application.EnableEvents = False
            Globals.SapAccAddIn.Application.ScreenUpdating = False
            ' read the invoice reposting lines
            i = aLOffBData + 1
            Do
                If aUseBasis _
                    Or CStr(aIWs.Cells(i, ColNew).value) <> "" Or CStr(aIWs.Cells(i, ColNew + 1).value) <> "" Or CStr(aIWs.Cells(i, ColNew + 2).value) <> "" _
                    Or CStr(aIWs.Cells(i, ColNew + 3).value) <> "" Or CStr(aIWs.Cells(i, ColNew + 4).value) <> "" Or CStr(aIWs.Cells(i, ColNew + 5).value) <> "" Then
                    aBasisLine = aMigHelper.makeDictForRules(aIWs, i, aLOffBNames + 1, aBaseColFrom, aBaseColTo)
                    If Not aMigHelper.isFiltered(aBasisLine) Then
                        aBasis.Add(aBasisLine)
                    End If
                End If
                i = i + 1
            Loop While CStr(aIWs.Cells(i, 1).value) <> ""

            Dim aTPostingData As New TPostingData(aIntPar)
            Dim aTPostingDataRec As TPostingDataRec
            Dim aTPostingDataRecKey As String
            Dim aTPostingDataRecNum As UInt64 = 1
            ' create the posting lines
            aPostingLines = New Collection
            For Each aBasisLine In aBasis
                aPostingLine = aMigHelper.mig.ApplyRules(aBasisLine, "P")
                If aPostingLine.Count > 0 Then
                    aTPostingDataRec = aTPostingData.newTPostingDataRec(pDic:=aPostingLine, pEmpty:=aGenEmpty, pEmptyChar:=aEmptyChar)
                    If aCompress Then
                        aTPostingDataRecKey = aTPostingDataRec.getKey()
                    Else
                        aTPostingDataRecKey = CStr(aTPostingDataRecNum)
                    End If
                    aTPostingData.addTPostingDataRec(aTPostingDataRecKey, aTPostingDataRec)
                    aTPostingDataRecNum += 1
                End If
                aContraLine = aMigHelper.mig.ApplyRules(aBasisLine, "C")
                If aContraLine.Count > 0 Then
                    aTPostingDataRec = aTPostingData.newTPostingDataRec(pDic:=aContraLine, pEmpty:=aGenEmpty, pEmptyChar:=aEmptyChar)
                    If aCompress Then
                        aTPostingDataRecKey = aTPostingDataRec.getKey()
                    Else
                        aTPostingDataRecKey = CStr(aTPostingDataRecNum)
                    End If
                    aTPostingData.addTPostingDataRec(aTPostingDataRecKey, aTPostingDataRec)
                    aTPostingDataRecNum += 1
                End If
            Next
            'output the posting lines
            Dim aColDelData As Integer = If(aIntPar.value("GEN" & pNr, "DATA_COLDEL") <> "", CInt(aIntPar.value("GEN" & pNr, "DATA_COLDEL")), 1)
            Try
                aDWs = aWB.Worksheets(aDwsName)
            Catch Exc As System.Exception
                Globals.SapAccAddIn.Application.EnableEvents = True
                Globals.SapAccAddIn.Application.ScreenUpdating = True
                Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
                MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Accounting Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
                Exit Sub
            End Try
            Dim aRange As Excel.Range
            i = aLOff + 1
            Do
                i += 1
            Loop While CStr(aDWs.Cells(i, aColDelData).Value) <> ""
            If aGenDeleteData And i >= aLOff + 1 Then
                aRange = aDWs.Range(aDWs.Cells(aLOff + 1, 1), aDWs.Cells(i, 1))
                aRange.EntireRow.Delete()
                i = aLOff + 1
            End If
            Dim jMax As Integer = 0
            Do
                jMax += 1
            Loop While CStr(aDWs.Cells(aLOff, jMax + 1).value) <> ""
            Dim aKey As String
            Dim aValue As String
            Dim aKvb As KeyValuePair(Of String, TPostingDataRec)
            i = If(aLineOut <> 0, aLineOut, i)
            Dim aSuppressLine As Boolean
            For Each aKvb In aTPostingData.aTPostingDataDic
                aSuppressLine = False
                aTPostingDataRec = aKvb.Value
                If aSupprZero And aTPostingDataRec.isZero Then
                    aSuppressLine = True
                End If
                If Not String.IsNullOrEmpty(aTargetFilter) Then
                    If isTargetFiltered(aTargetFilter, aTPostingDataRec) Then
                        aSuppressLine = True
                    End If
                End If
                If Not aSuppressLine Then
                    For j = 1 To jMax
                        If CStr(aDWs.Cells(aLOff, j).Value) <> "" Then
                            aKey = CStr(aDWs.Cells(aLOff, j).Value)
                            If Not aKey.Contains("-") Then
                                aKey = "-" & aKey
                            End If
                            If aTPostingDataRec.aTPostingDataRecCol.Contains(aKey) Then
                                aValue = aTPostingDataRec.aTPostingDataRecCol(aKey).Value
                                If aTPostingDataRec.isValue(aKey) Then
                                    aDWs.Cells(i, j).Value = CDbl(aValue)
                                ElseIf aTPostingDataRec.aTPostingDataRecCol(aKey).Format = "F" Then
                                    aDWs.Cells(i, j).FormulaR1C1 = "=" & CStr(aValue)
                                Else
                                    aDWs.Cells(i, j).Value = aValue
                                End If
                            End If
                        End If
                    Next j
                    i += 1
                End If
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

    Private Sub ButtonGeneratePostings_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonGeneratePostings.Click
        Dim aIWs As Excel.Worksheet
        Dim aDWs As Excel.Worksheet
        Dim aAccWS As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aAccDic As New Dictionary(Of String, String)
        Dim aScaleDic As New Dictionary(Of String, Double)
        Dim aBasisLine As New Dictionary(Of String, SAPCommon.TField)
        Dim aPostingLine As New Dictionary(Of String, SAPCommon.TField)
        Dim aContraLine As New Dictionary(Of String, SAPCommon.TField)
        Dim aPostingLines As Collection
        Dim aAcc As String
        Dim aScale As Double
        Dim i As Integer
        Dim j As Integer

        ' get internal parameters
        If Not getIntParameters() Then
            Exit Sub
        End If
        Dim aLOff As Integer = If(aIntPar.value("LOFF", "DATA") <> "", CInt(aIntPar.value("LOFF", "DATA")), 4)
        Dim CTot As Integer = If(aIntPar.value("COL", "COSTOT") <> "", CInt(aIntPar.value("COL", "COSTOT")), 12)

        log.Debug("ButtonGenGLData_Click - " & "Invoice Sheet")
        aWB = Globals.SapAccAddIn.Application.ActiveWorkbook
        Try
            aIWs = aWB.Worksheets("Basis")
        Catch Exc As System.Exception
            log.Warn("ButtonGenGLData_Click - " & "No Basis Sheet in current workbook.")
            MsgBox("No Basis Sheet in current workbook. Check if the current workbook is a valid Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            Exit Sub
        End Try
        aIWs.Activate()

        Dim aAccKey As String
        ' get the account mapping
        Try
            aAccWS = aWB.Worksheets("Mapping")
        Catch Exc As System.Exception
            log.Warn("ButtonGenGLData_Click - " & "No Mapping Sheet in current workbook.")
            MsgBox("No Mapping Sheet in current workbook. Check if the current workbook is a valid Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            Exit Sub
        End Try
        i = 2
        Do
            aAccKey = CStr(aAccWS.Cells(i, 1).Value)
            aAccDic.Add(aAccKey, CStr(aAccWS.Cells(i, 2).Value))
            aScaleDic.Add(aAccKey, CDbl(aAccWS.Cells(i, 3).Value))
            i += 1
        Loop While Not String.IsNullOrEmpty(CStr(aAccWS.Cells(i, 1).Value))

        ' process the data
        Dim aBasItems As New TData(aIntPar)
        Dim aBasItem As New TDataRec
        Dim lineKey As String
        Dim aRestRange As Excel.Range
        Dim aSumRest As Double
        Dim aPost As String

        Dim CLast As Integer = 0
        Do
            CLast += 1
        Loop While CStr(aIWs.Cells(aLOff, CLast + 1).value) <> ""

        Try
            log.Debug("ButtonGenGLData_Click - " & "processing data - disabling events, screen update, cursor")
            Globals.SapAccAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapAccAddIn.Application.EnableEvents = False
            Globals.SapAccAddIn.Application.ScreenUpdating = False
            ' read the invoice reposting lines
            i = 2
            Do
                For j = CTot To CLast
                    aAccKey = CStr(aIWs.Cells(1, j).Value)
                    aAcc = aAccDic.Item(aAccKey)
                    aScale = aScaleDic.Item(aAccKey)
                    aRestRange = aIWs.Range(aIWs.Cells(i, j + 1), aIWs.Cells(i, CLast))
                    aSumRest = totalAbs(aRestRange)
                    If aSumRest = 0 Or j = CLast Then
                        aPost = "X"
                    Else
                        aPost = ""
                    End If
                    If CDbl(aIWs.Cells(i, j).Value) <> 0 Then
                        lineKey = CStr(i) + "_" + CStr(j)
                        aBasItems.addValue(lineKey, "INT-ACCTYPE", "S", "", "")
                        aBasItems.addValue(lineKey, "INT-ACCOUNT", aAcc, "", "")
                        aBasItems.addValue(lineKey, "GL-MATERIAL", CStr(aIWs.Cells(i, 3).Value), "", "")
                        aBasItems.addValue(lineKey, "GL-PLANT", CStr(aIWs.Cells(i, 2).Value), "", "")
                        aBasItems.addValue(lineKey, "GL+CU+VE-PROFIT_CTR", CStr(aIWs.Cells(i, 4).Value), "", "")
                        aBasItems.addValue(lineKey, "GL+CU+VE-BUS_AREA", CStr(aIWs.Cells(i, 6).Value), "", "")
                        aBasItems.addValue(lineKey, "GL-SEGMENT", CStr(aIWs.Cells(i, 5).Value), "", "")
                        aBasItems.addValue(lineKey, "GL-FUNC_AREA", CStr(aIWs.Cells(i, 8).Value), "", "")
                        aBasItems.addValue(lineKey, "GL+CU+VE-ITEM_TEXT", CStr(aIWs.Cells(i, 3).Value), "", "")
                        aBasItems.addValue(lineKey, "A00-AMT_DOCCUR", CDbl(aIWs.Cells(i, j).Value) * aScale, "X", "")
                        aBasItems.addValue(lineKey, "INT-POST", aPost, "", "")
                    End If
                Next j
                i = i + 1
            Loop While Not String.IsNullOrEmpty(CStr(aIWs.Cells(i, 1).value))

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
            If CStr(aDWs.Cells(aLOff + 1, 1).Value) <> "" Then
                ' aRange = aDWs.Range(aDWs.Cells(aLOff + 1, 1))
                i = aLOff + 1
                Do
                    i += 1
                Loop While CStr(aDWs.Cells(i, 1).Value) <> ""
                aRange = aDWs.Range(aDWs.Cells(aLOff + 1, 1), aDWs.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            Dim jMax As Integer = 0
            Do
                jMax += 1
            Loop While CStr(aDWs.Cells(aLOff, jMax + 1).value) <> ""
            Dim aKey As String
            Dim aValue As String
            Dim aKvB As KeyValuePair(Of String, TDataRec)

            i = aLOff + 1
            For Each aKvB In aBasItems.aTDataDic
                aBasItem = aKvB.Value
                For j = 1 To jMax
                    If CStr(aDWs.Cells(1, j).Value) <> "" Then
                        aKey = CStr(aDWs.Cells(1, j).Value)
                        If aBasItem.aTDataRecCol.Contains(aKey) Then
                            aValue = aBasItem.aTDataRecCol(aKey).Value
                            If aBasItem.aTDataRecCol(aKey).Currency = "X" And aValue <> "" Then
                                aDWs.Cells(i, j).Value = CDbl(aValue)
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

    End Sub

    Private Function totalAbs(pRange As Excel.Range) As Double
        Dim aTot As Double
        For Each cell In pRange
            aTot = aTot + Math.Abs(CDbl(cell.Value))
        Next
        totalAbs = aTot
    End Function

    Private Function isTargetFiltered(pTargetFilterStr As String, pTPostingDataRec As TPostingDataRec) As Boolean
        Dim aFilterField As String = ""
        Dim aFilterOperation As String = ""
        Dim aFilterCompare As String = ""
        If Not String.IsNullOrEmpty(pTargetFilterStr) Then
            Dim aFilterStr() As String = {}
            aFilterStr = pTargetFilterStr.Split(";")
            If aFilterStr.Length = 3 Then
                aFilterField = aFilterStr(0)
                aFilterOperation = aFilterStr(1)
                aFilterCompare = aFilterStr(2)
                If aFilterCompare.ToUpper() = "NULL" Then
                    aFilterCompare = ""
                End If
            End If
        End If
        isTargetFiltered = False
        Dim aTStrRec As SAPCommon.TStrRec
        If pTPostingDataRec.aTPostingDataRecCol.Contains("-" & aFilterField) Then
            aTStrRec = pTPostingDataRec.aTPostingDataRecCol("-" & aFilterField)
            If aFilterOperation = "EQ" And aTStrRec.Value = aFilterCompare Then
                isTargetFiltered = True
            ElseIf aFilterOperation = "NE" And aTStrRec.Value <> aFilterCompare Then
                isTargetFiltered = True
            End If
        Else
            If aFilterOperation = "NE" And (String.IsNullOrEmpty(aFilterCompare) Or aFilterCompare = "#") Then
                isTargetFiltered = False
            ElseIf aFilterOperation = "EQ" And (String.IsNullOrEmpty(aFilterCompare) Or aFilterCompare = "#") Then
                isTargetFiltered = True
            End If
        End If
    End Function

End Class
