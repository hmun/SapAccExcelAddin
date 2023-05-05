' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPAcctngDocument

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapConHelper

    Sub New(aSapCon As SapConHelper)
        Try
            log.Debug("New - " & "checking connection")
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            log.Error("New - Exception=" & ex.ToString)
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAcctngDocument")
        End Try
    End Sub

    Public Sub getMeta(ByRef pHdr_Fields() As String, ByRef pGla_Fields() As String, ByRef pCus_Fields() As String, ByRef pVen_Fields() As String, ByRef pTax_Fields() As String, ByRef pAmt_Fields() As String, ByRef pCpd_Fields() As String, pTest As Boolean)
        Try
            If pTest Then
                log.Debug("post - " & "creating Function BAPI_ACC_DOCUMENT_CHECK")
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_DOCUMENT_CHECK")
            Else
                log.Debug("post - " & "creating Function BAPI_ACC_DOCUMENT_POST")
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_DOCUMENT_POST")
            End If
            Dim oDocumentHeader As IRfcStructure = oRfcFunction.GetStructure("DOCUMENTHEADER")
            Dim oCustomerCPD As IRfcStructure = oRfcFunction.GetStructure("CUSTOMERCPD")
            Dim oAccountGl As IRfcTable = oRfcFunction.GetTable("ACCOUNTGL")
            Dim oAccountTax As IRfcTable = oRfcFunction.GetTable("ACCOUNTTAX")
            Dim oAccountPayable As IRfcTable = oRfcFunction.GetTable("ACCOUNTPAYABLE")
            Dim oAccountReceivable As IRfcTable = oRfcFunction.GetTable("ACCOUNTRECEIVABLE")
            Dim oCurrencyAmount As IRfcTable = oRfcFunction.GetTable("CURRENCYAMOUNT")
            ' Hdr_Fields
            pHdr_Fields = {}
            For i As Integer = 0 To oDocumentHeader.Metadata.FieldCount - 1
                Array.Resize(pHdr_Fields, pHdr_Fields.Length + 1)
                pHdr_Fields(pHdr_Fields.Length - 1) = oDocumentHeader.Metadata(i).Name
            Next
            ' Cpd_Fields
            pCpd_Fields = {}
            For i As Integer = 0 To oCustomerCPD.Metadata.FieldCount - 1
                Array.Resize(pCpd_Fields, pCpd_Fields.Length + 1)
                pCpd_Fields(pCpd_Fields.Length - 1) = oCustomerCPD.Metadata(i).Name
            Next
            ' Gla_Fields - skip first field which is the ITEMNO_ACC
            pGla_Fields = {}
            For i As Integer = 1 To oAccountGl.Metadata.LineType.FieldCount - 1
                Array.Resize(pGla_Fields, pGla_Fields.Length + 1)
                pGla_Fields(pGla_Fields.Length - 1) = oAccountGl.Metadata.LineType(i).Name
            Next
            ' Cus_Fields - skip first field which is the ITEMNO_ACC
            pCus_Fields = {}
            For i As Integer = 1 To oAccountReceivable.Metadata.LineType.FieldCount - 1
                Array.Resize(pCus_Fields, pCus_Fields.Length + 1)
                pCus_Fields(pCus_Fields.Length - 1) = oAccountReceivable.Metadata.LineType(i).Name
            Next
            ' Ven_Fields - skip first field which is the ITEMNO_ACC
            pVen_Fields = {}
            For i As Integer = 1 To oAccountPayable.Metadata.LineType.FieldCount - 1
                Array.Resize(pVen_Fields, pVen_Fields.Length + 1)
                pVen_Fields(pVen_Fields.Length - 1) = oAccountPayable.Metadata.LineType(i).Name
            Next
            ' Tax_Fields - skip first field which is the ITEMNO_ACC
            pTax_Fields = {}
            For i As Integer = 1 To oAccountTax.Metadata.LineType.FieldCount - 1
                Array.Resize(pTax_Fields, pTax_Fields.Length + 1)
                pTax_Fields(pTax_Fields.Length - 1) = oAccountTax.Metadata.LineType(i).Name
            Next
            ' Amt_Fields - skip first field which is the ITEMNO_ACC
            pAmt_Fields = {}
            For i As Integer = 1 To oCurrencyAmount.Metadata.LineType.FieldCount - 1
                Array.Resize(pAmt_Fields, pAmt_Fields.Length + 1)
                pAmt_Fields(pAmt_Fields.Length - 1) = oCurrencyAmount.Metadata.LineType(i).Name
            Next
        Catch Ex As System.Exception
            log.Error("getMeta - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAcctngDocument")
        Finally
            log.Debug("getMeta - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function post(ByRef pTSAP_DocData As TSAP_DocData, pTest As Boolean) As String
        Dim aSAPCalcTaxesFromGross As New SAPCalcTaxesFromGross(sapcon)
        Dim aSapFormat As New SAPCommon.SAPFormat()
        post = ""
        Try
            If pTest Then
                log.Debug("post - " & "creating Function BAPI_ACC_DOCUMENT_CHECK")
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_DOCUMENT_CHECK")
            Else
                log.Debug("post - " & "creating Function BAPI_ACC_DOCUMENT_POST")
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_DOCUMENT_POST")
            End If

            log.Debug("post - " & "oRfcFunction.Metadata.Name=" & oRfcFunction.Metadata.Name)
            log.Debug("post - " & "BeginContext")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPWbsElement As New SAPWbsElement(sapcon)
            log.Debug("post - " & "Getting Function parameters")
            Dim oDocumentHeader As IRfcStructure = oRfcFunction.GetStructure("DOCUMENTHEADER")
            Dim oCustomerCPD As IRfcStructure = oRfcFunction.GetStructure("CUSTOMERCPD")
            Dim oAccountGl As IRfcTable = oRfcFunction.GetTable("ACCOUNTGL")
            Dim oAccountTax As IRfcTable = oRfcFunction.GetTable("ACCOUNTTAX")
            Dim oAccountPayable As IRfcTable = oRfcFunction.GetTable("ACCOUNTPAYABLE")
            Dim oAccountReceivable As IRfcTable = oRfcFunction.GetTable("ACCOUNTRECEIVABLE")
            Dim oCurrencyAmount As IRfcTable = oRfcFunction.GetTable("CURRENCYAMOUNT")
            Dim oCriteria As IRfcTable = oRfcFunction.GetTable("CRITERIA")
            Dim oExtension2 As IRfcTable = oRfcFunction.GetTable("EXTENSION2")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oAccountGl.Clear()
            oAccountTax.Clear()
            oAccountPayable.Clear()
            oAccountReceivable.Clear()
            oCurrencyAmount.Clear()
            oCriteria.Clear()
            oExtension2.Clear()
            oRETURN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            ' fill the header
            log.Debug("post - " & "setting header values")
            oDocumentHeader.SetValue("BUS_ACT", "RFBU")
            If destination.User Is Nothing Then
                oDocumentHeader.SetValue("USERNAME", destination.SystemAttributes.User)
            Else
                oDocumentHeader.SetValue("USERNAME", destination.User)
            End If
            For Each aTStrRec In pTSAP_DocData.aHdrRec.aTDataRecCol
                oDocumentHeader.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
            Next
            ' fill the data
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            Dim lCnt As Integer
            Dim lCntSav As Integer
            lCnt = 0
            For Each aKvP In pTSAP_DocData.aData.aTDataDic
                Dim oAccountGlAppended As Boolean = False
                Dim oAccountReceivableAppended As Boolean = False
                Dim oAccountPayableAppended As Boolean = False
                Dim oExtension2Appended As Boolean = False
                Dim lastStrucname As String
                Dim lTaxCode As String = ""
                Dim lTaxJurCode As String = ""
                aTDataRec = aKvP.Value
                lCnt += 1
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "CUSTOMERCPD"
                            oCustomerCPD.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                        Case "ACCOUNTGL"
                            If Not oAccountGlAppended Then
                                oAccountGl.Append()
                                oAccountGl.SetValue("ITEMNO_ACC", lCnt)
                                oAccountGlAppended = True
                            End If
                            oAccountGl.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                            If aTStrRec.Fieldname = "TAX_CODE" Then
                                lTaxCode = aTStrRec.Value
                            ElseIf aTStrRec.Fieldname = "TAXJURCODE" Then
                                lTaxJurCode = aTStrRec.Value
                            End If
                        Case "ACCOUNTRECEIVABLE"
                            If Not oAccountReceivableAppended Then
                                oAccountReceivable.Append()
                                oAccountReceivable.SetValue("ITEMNO_ACC", lCnt)
                                oAccountReceivableAppended = True
                            End If
                            oAccountReceivable.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                        Case "ACCOUNTPAYABLE"
                            If Not oAccountPayableAppended Then
                                oAccountPayable.Append()
                                oAccountPayable.SetValue("ITEMNO_ACC", lCnt)
                                oAccountPayableAppended = True
                            End If
                            oAccountPayable.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                        Case "CRITERIA"
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", aTStrRec.Fieldname)
                            If aTStrRec.Format = "WBS" Then
                                oCriteria.SetValue("CHARACTER", lSAPWbsElement.GetPspnr(aTStrRec.Value))
                            Else
                                oCriteria.SetValue("CHARACTER", aTStrRec.formated())
                            End If
                        Case "ACCOUNTTAX"
                        Case "CURRENCYAMOUNT" ' should never happen because amounts are in the aAmounts Dictionary
                        Case Else
                            If Not oExtension2Appended Or lastStrucname <> aTStrRec.Strucname Then
                                oExtension2.Append()
                                oExtension2.SetValue("STRUCTURE", aTStrRec.Strucname)
                                oExtension2.SetValue("VALUEPART1", aSapFormat.unpack(CStr(lCnt), 10))
                                oExtension2Appended = True
                                lastStrucname = aTStrRec.Strucname
                            End If
                            oExtension2.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
                Dim aDummyPar As New SAPCommon.TStr
                Dim aData_Am As New TData(aDummyPar)
                Dim aDataRec_Am As New TDataRec
                Dim aTStrRec_Am As New SAPCommon.TStrRec
                ' read the amounts for the data record
                aData_Am = pTSAP_DocData.aAmounts(aKvP.Key)
                lCntSav = lCnt
                Dim oAccountTaxAppended = False
                For Each aKvB_Am In aData_Am.aTDataDic
                    Dim oCurrencyAmountAppended = False
                    Dim aCurrType As String = ""
                    Dim aCurrency As String = ""
                    lCnt = lCntSav
                    aDataRec_Am = aKvB_Am.Value
                    For Each aTStrRec_Am In aDataRec_Am.aTDataRecCol
                        If Not oCurrencyAmountAppended Then
                            oCurrencyAmount.Append()
                            oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                            oCurrencyAmountAppended = True
                        End If
                        oCurrencyAmount.SetValue(aTStrRec_Am.Fieldname, aTStrRec_Am.formated())
                        If aTStrRec_Am.Fieldname = "CURR_TYPE" Then
                            aCurrType = aTStrRec_Am.Value
                            aCurrency = pTSAP_DocData.getCurrency(aTStrRec_Am.Value)
                            oCurrencyAmount.SetValue("CURRENCY", aCurrency)
                        End If
                    Next
                    ' check if this is a direct postin to a tax account
                    Dim lTaxAmtOri As Double = 0
                    Dim lTaxBaseOri As Double = 0
                    If oCurrencyAmount.CurrentRow("TAX_AMT").GetDouble() <> 0 And
                        oCurrencyAmount.CurrentRow("TAX_AMT").GetDouble() = oCurrencyAmount.CurrentRow("AMT_DOCCUR").GetDouble() Then
                        ' set the original position to 0 - or should we delete the row?
                        lTaxAmtOri = oCurrencyAmount.CurrentRow("TAX_AMT").GetDouble()
                        lTaxBaseOri = oCurrencyAmount.CurrentRow("AMT_BASE").GetDouble()
                        oCurrencyAmount.SetValue("TAX_AMT", Format$(CDbl(0), "0.00"))
                        oCurrencyAmount.SetValue("AMT_BASE", Format$(CDbl(0), "0.00"))
                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(CDbl(0), "0.00"))
                        lCnt += 1
                        If Not oAccountTaxAppended Then
                            oAccountTax.Append()
                            oAccountTax.SetValue("ITEMNO_ACC", lCnt)
                            oAccountTax.SetValue("ITEMNO_ACC", lCnt)
                            Dim lCondKey As String = aSAPCalcTaxesFromGross.getTaxCondKey(oDocumentHeader("COMP_CODE").GetString(), lTaxCode, aCurrency,
                                                                            CDate(oDocumentHeader("PSTNG_DATE").GetString()),
                                                                            oCurrencyAmount.CurrentRow("AMT_DOCCUR").GetDouble(), lTaxJurCode)
                            oAccountTax.SetValue("COND_KEY", lCondKey)
                            oAccountTax.SetValue("TAX_CODE", lTaxCode)
                            oAccountTax.SetValue("TAXJURCODE", lTaxJurCode)
                            oAccountTax.SetValue("DIRECT_TAX", "X")
                            oAccountTaxAppended = True
                        End If
                        oCurrencyAmount.Append()
                        oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                        oCurrencyAmount.SetValue("CURR_TYPE", aCurrType)
                        oCurrencyAmount.SetValue("CURRENCY", aCurrency)
                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lTaxAmtOri, "0.00"))
                        If lTaxBaseOri <> 0 Then
                            oCurrencyAmount.SetValue("AMT_BASE", Format$(lTaxBaseOri, "0.00"))
                        Else
                            oCurrencyAmount.SetValue("AMT_BASE", Format$(lTaxAmtOri, "0.00"))
                        End If
                        ' calculate and add VAT
                    ElseIf lTaxCode <> "" Then
                        Dim lTaxSum As Double = 0
                        Dim lTaxBase As Double = 0
                        Dim lTaxLines As Integer = 0
                        Dim oTAX_ITEM_OUT As IRfcTable
                        log.Debug("post - " & "calling aSAPCalcTaxesFromGross.getTaxAmount")
                        oTAX_ITEM_OUT = aSAPCalcTaxesFromGross.getTaxAmount(oDocumentHeader("COMP_CODE").GetString(), lTaxCode, aCurrency,
                                                                            CDate(oDocumentHeader("PSTNG_DATE").GetString()), oCurrencyAmount.CurrentRow("AMT_DOCCUR").GetDouble(), lTaxJurCode)
                        lTaxLines = oTAX_ITEM_OUT.Count
                        ' calculate the taxsum
                        lTaxSum = 0
                        For i As Integer = 0 To lTaxLines - 1
                            lTaxSum = lTaxSum + oTAX_ITEM_OUT(i).GetDouble("FWSTE")
                        Next i
                        lTaxBase = oCurrencyAmount.CurrentRow("AMT_DOCCUR").GetDouble() - lTaxSum
                        log.Debug("post - " & "lTaxLines=" & CStr(lTaxLines) & " lTaxSum=" & CStr(lTaxSum) & " lTaxBase=" & CStr(lTaxBase))
                        ' change the ammount of the row to the net value
                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lTaxBase, "0.00"))
                        ' add the tax positions
                        If lTaxSum <> 0 Or lTaxLines > 1 Then
                            For i As Integer = 0 To lTaxLines - 1
                                lCnt += 1
                                oAccountTax.Append()
                                oAccountTax.SetValue("ITEMNO_ACC", lCnt)
                                oAccountTax.SetValue("COND_KEY", oTAX_ITEM_OUT(i).GetValue("KSCHL"))
                                oAccountTax.SetValue("TAX_CODE", oTAX_ITEM_OUT(i).GetValue("MWSKZ"))
                                oAccountTax.SetValue("TAXJURCODE", oTAX_ITEM_OUT(i).GetValue("TXJCD"))
                                oCurrencyAmount.Append()
                                oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                                oCurrencyAmount.SetValue("CURR_TYPE", aCurrType)
                                oCurrencyAmount.SetValue("CURRENCY", aCurrency)
                                oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(oTAX_ITEM_OUT(i).GetDouble("FWSTE"), "0.00"))
                                oCurrencyAmount.SetValue("AMT_BASE", Format$(oTAX_ITEM_OUT(i).GetDouble("FWBAS"), "0.00"))
                            Next i
                        End If
                    End If
                Next
            Next
            ' call the BAPI
            log.Debug("post - " & "invoking " & oRfcFunction.Metadata.Name)
            oRfcFunction.Invoke(destination)
            log.Debug("post - " & "oRETURN.Count=" & CStr(oRETURN.Count))
            If oRETURN.Count > 0 Then
                If oRETURN(0).GetValue("TYPE") = "S" Then
                    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                    If Not pTest Then
                        log.Debug("post - " & "calling aSAPBapiTranctionCommit.commit()")
                        aSAPBapiTranctionCommit.commit()
                    End If
                    post = oRETURN(0).GetValue("MESSAGE")
                Else
                    For i As Integer = 0 To oRETURN.Count - 1
                        post = post & ";" & oRETURN(i).GetValue("MESSAGE")
                    Next i
                End If
            Else
                log.Debug("post - " & "Error: No Return message from SAP")
                post = "Error: No Return message from SAP"
            End If
        Catch Ex As System.Exception
            log.Error("commit - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAcctngDocument")
            post = "Error: Exception in posting"
        Finally
            log.Debug("post - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function post(pBLDAT As Date, pBLART As String, pBUKRS As String,
        pBUDAT As Date, pWAERS As String, pXBLNR As String,
        pBKTXT As String, pFIS_PERIOD As Integer, pACC_PRINCIPLE As String, pData As Collection, pTest As Boolean,
        pFKBERNAME As String, Optional pTRANS_DATE As Date? = Nothing) As String

        post = ""
        Try
            If pTest Then
                log.Debug("post - " & "creating Function BAPI_ACC_DOCUMENT_CHECK")
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_DOCUMENT_CHECK")
            Else
                log.Debug("post - " & "creating Function BAPI_ACC_DOCUMENT_POST")
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_DOCUMENT_POST")
            End If
            log.Debug("post - " & "oRfcFunction.Metadata.Name=" & oRfcFunction.Metadata.Name)
            log.Debug("post - " & "BeginContext")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim lSAPWbsElement As New SAPWbsElement(sapcon)
            log.Debug("post - " & "Getting Function parameters")
            Dim oDocumentHeader As IRfcStructure = oRfcFunction.GetStructure("DOCUMENTHEADER")
            Dim oAccountGl As IRfcTable = oRfcFunction.GetTable("ACCOUNTGL")
            Dim oAccountTax As IRfcTable = oRfcFunction.GetTable("ACCOUNTTAX")
            Dim oAccountPayable As IRfcTable = oRfcFunction.GetTable("ACCOUNTPAYABLE")
            Dim oAccountReceivable As IRfcTable = oRfcFunction.GetTable("ACCOUNTRECEIVABLE")
            Dim oCurrencyAmount As IRfcTable = oRfcFunction.GetTable("CURRENCYAMOUNT")
            Dim oCriteria As IRfcTable = oRfcFunction.GetTable("CRITERIA")
            Dim oExtension2 As IRfcTable = oRfcFunction.GetTable("EXTENSION2")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oAccountGl.Clear()
            oAccountTax.Clear()
            oAccountPayable.Clear()
            oAccountReceivable.Clear()
            oCurrencyAmount.Clear()
            oCriteria.Clear()
            oExtension2.Clear()
            oRETURN.Clear()

            log.Debug("post - " & "setting header values")
            oDocumentHeader.SetValue("BUS_ACT", "RFBU")
            oDocumentHeader.SetValue("ACC_PRINCIPLE", pACC_PRINCIPLE)
            oDocumentHeader.SetValue("COMP_CODE", pBUKRS)
            oDocumentHeader.SetValue("PSTNG_DATE", pBUDAT)
            oDocumentHeader.SetValue("FIS_PERIOD", pFIS_PERIOD) '23.01.2012 Buchungsperiode
            oDocumentHeader.SetValue("DOC_DATE", pBLDAT)
            If Not pTRANS_DATE Is Nothing Then
                oDocumentHeader.SetValue("TRANS_DATE", pTRANS_DATE)
            End If
            If destination.User Is Nothing Then
                oDocumentHeader.SetValue("USERNAME", destination.SystemAttributes.User)
            Else
                oDocumentHeader.SetValue("USERNAME", destination.User)
            End If
            oDocumentHeader.SetValue("DOC_TYPE", pBLART)
            oDocumentHeader.SetValue("REF_DOC_NO", pXBLNR)
            oDocumentHeader.SetValue("HEADER_TXT", pBKTXT)
            Dim lRow As Object
            Dim lCnt As Integer
            Dim lCntSav As Integer
            lCnt = 0
            log.Debug("post - " & "processing pData")
            For Each lRow In pData
                lCnt = lCnt + 1
                If lRow.ACCTYPE = "S" Or lRow.ACCTYPE = "G" Then
                    log.Debug("post - " & "adding AccountGl ITEMNO_ACC=" & CStr(lCnt) & " GL_ACCOUNT=" & lSAPFormat.unpack(lRow.NEWKO, 10))
                    oAccountGl.Append()
                    oAccountGl.SetValue("ITEMNO_ACC", lCnt)
                    oAccountGl.SetValue("GL_ACCOUNT", lSAPFormat.unpack(lRow.NEWKO, 10))
                    oAccountGl.SetValue("ITEM_TEXT", lRow.SGTXT)
                    If lRow.TXJCD <> "" Then
                        oAccountGl.SetValue("TAXJURCODE", lRow.TXJCD)
                    End If
                    oAccountGl.SetValue("TAX_CODE", lRow.MWSKZ)
                    oAccountGl.SetValue("ALLOC_NMBR", lRow.ALLOC_NMBR)
                    oAccountGl.SetValue("REF_KEY_3", lRow.REF_KEY_3)
                    If lRow.COMP_CODE <> "" Then
                        oAccountGl.SetValue("COMP_CODE", lRow.COMP_CODE)
                    Else
                        oAccountGl.SetValue("COMP_CODE", pBUKRS)
                    End If
                    If lRow.PRCTR <> "" Then
                        oAccountGl.SetValue("PROFIT_CTR", lRow.PRCTR)
                    End If
                    If lRow.PART_PRCTR <> "" Then
                        oAccountGl.SetValue("PART_PRCTR", lRow.PART_PRCTR)
                    End If
                    If lRow.SEGMENT <> "" Then
                        oAccountGl.SetValue("SEGMENT", lSAPFormat.unpack(lRow.SEGMENT, 10))
                    End If
                    If lRow.PARTNER_SEGMENT <> "" Then
                        oAccountGl.SetValue("PARTNER_SEGMENT", lSAPFormat.unpack(lRow.PARTNER_SEGMENT, 10))
                    End If
                    If lRow.BEWAR <> "" Then
                        oAccountGl.SetValue("CS_TRANS_T", Right(lSAPFormat.unpack(lRow.BEWAR, 10), 3))
                    End If
                    If lRow.FUNC_AREA <> "" Then
                        oAccountGl.SetValue("FUNC_AREA", lSAPFormat.unpack(lRow.FUNC_AREA, 4))
                    End If
                    If lRow.TRADE_ID <> "" Then
                        oAccountGl.SetValue("TRADE_ID", lSAPFormat.unpack(lRow.TRADE_ID, 6))
                    End If
                    If lRow.BUS_AREA <> "" Then
                        oAccountGl.SetValue("BUS_AREA", lSAPFormat.unpack(lRow.BUS_AREA, 4))
                    End If
                    ' Extensions Fields
                    If lRow.ZZETXT <> "" Then
                        oExtension2.Append()
                        oExtension2.SetValue("STRUCTURE", "ZFI_EXT2_ZZETXT")
                        oExtension2.SetValue("VALUEPART1", lSAPFormat.unpack(CStr(lCnt), 10))
                        oExtension2.SetValue("VALUEPART2", lRow.ZZETXT)
                    End If
                    ' HFM Sales Country (in ISO 3 Char)
                    If lRow.ZZHFMC1 <> "" Then
                        oExtension2.Append()
                        oExtension2.SetValue("STRUCTURE", "ZFI_EXT2_ZZHFMC1")
                        oExtension2.SetValue("VALUEPART1", lSAPFormat.unpack(CStr(lCnt), 10))
                        oExtension2.SetValue("VALUEPART2", lRow.ZZHFMC1)
                    End If
                    ' HFM Customer Group
                    If lRow.ZZHFMC3 <> "" Then
                        oExtension2.Append()
                        oExtension2.SetValue("STRUCTURE", "ZFI_EXT2_ZZHFMC3")
                        oExtension2.SetValue("VALUEPART1", lSAPFormat.unpack(CStr(lCnt), 10))
                        oExtension2.SetValue("VALUEPART2", lSAPFormat.unpack(lRow.ZZHFMC3, 3))
                    End If
                    ' Bilanzkennung
                    If lRow.ZZBBKNG <> "" Then
                        oExtension2.Append()
                        oExtension2.SetValue("STRUCTURE", "ZFI_EXT2_ZZBBKNG")
                        oExtension2.SetValue("VALUEPART1", lSAPFormat.unpack(CStr(lCnt), 10))
                        oExtension2.SetValue("VALUEPART2", lSAPFormat.unpack(lRow.ZZBBKNG, 8))
                    End If
                    ' HFM Customer Group
                    If lRow.ZZBBTCO <> "" Then
                        oExtension2.Append()
                        oExtension2.SetValue("STRUCTURE", "ZFI_EXT2_ZZBBTCO")
                        oExtension2.SetValue("VALUEPART1", lSAPFormat.unpack(CStr(lCnt), 10))
                        oExtension2.SetValue("VALUEPART2", lSAPFormat.unpack(lRow.ZZBBTCO, 4))
                    End If
                    ' Region, OEM (for Magna Template System)
                    If lRow.ZZDIM06 <> "" Or lRow.ZZDIM07 <> "" Then
                        oExtension2.Append()
                        oExtension2.SetValue("STRUCTURE", "ZFI_BAPIEXT_ST")
                        oExtension2.SetValue("VALUEPART1", lSAPFormat.unpack(CStr(lCnt), 10))
                        oExtension2.SetValue("VALUEPART2", lSAPFormat.fixLen(lRow.ZZDIM06, 20))
                        oExtension2.SetValue("VALUEPART3", lSAPFormat.fixLen(lRow.ZZDIM07, 20))
                    End If
                    ' Business Place
                    If lRow.BUPLA <> "" Then
                        oExtension2.Append()
                        oExtension2.SetValue("STRUCTURE", "ZFI_BAPIEXT_BUPLA")
                        oExtension2.SetValue("VALUEPART1", lSAPFormat.unpack(CStr(lCnt), 10))
                        oExtension2.SetValue("VALUEPART2", lSAPFormat.unpack(lRow.BUPLA, 4))
                    End If
                    ' CO-PA charactereistics
                    If lRow.PA = "X" Or lRow.PA = "x" Then
                        '  BUKRS
                        oCriteria.Append()
                        oCriteria.SetValue("ITEMNO_ACC", lCnt)
                        oCriteria.SetValue("FIELDNAME", "BUKRS")
                        If lRow.COMP_CODE <> "" Then
                            oCriteria.SetValue("CHARACTER", lRow.COMP_CODE)
                        Else
                            oCriteria.SetValue("CHARACTER", pBUKRS)
                        End If
                        '  VKORG
                        If lRow.VKORG <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "VKORG")
                            oCriteria.SetValue("CHARACTER", lRow.VKORG)
                        End If
                        '  VTWEG
                        If lRow.VTWEG <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "VTWEG")
                            oCriteria.SetValue("CHARACTER", lSAPFormat.unpack(lRow.VTWEG, 2))
                        End If
                        '  SPART
                        If lRow.SPART <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "SPART")
                            oCriteria.SetValue("CHARACTER", lSAPFormat.unpack(lRow.SPART, 2))
                        End If
                        '  KNDNR
                        If lRow.KNDNR <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "KNDNR")
                            oCriteria.SetValue("CHARACTER", lSAPFormat.unpack(lRow.KNDNR, 10))
                        End If
                        '  WERKS
                        If lRow.WERKS <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "WERKS")
                            oCriteria.SetValue("CHARACTER", lRow.WERKS)
                        End If
                        '  ARTNR
                        If lRow.MATNR <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "ARTNR")
                            oCriteria.SetValue("CHARACTER", lRow.MATNR)
                        End If
                        '  KTGRM
                        If lRow.KTGRM <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "KTGRM")
                            oCriteria.SetValue("CHARACTER", lSAPFormat.unpack(lRow.KTGRM, 2))
                        End If
                        '  GSBER
                        If lRow.BUS_AREA <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "GSBER")
                            oCriteria.SetValue("CHARACTER", lSAPFormat.unpack(lRow.BUS_AREA, 4))
                        End If
                        '  SEGMENT
                        If lRow.SEGMENT <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "SEGMENT")
                            oCriteria.SetValue("CHARACTER", lSAPFormat.unpack(lRow.SEGMENT, 10))
                        End If
                        'PARTNER_SEGMENT
                        If lRow.PARTNER_SEGMENT <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "WWPSE")
                            oCriteria.SetValue("CHARACTER", lSAPFormat.unpack(lRow.PARTNER_SEGMENT, 10))
                        End If
                        ' PRCTR
                        If lRow.PRCTR <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "PRCTR")
                            oCriteria.SetValue("CHARACTER", lRow.PRCTR)
                        End If
                        ' PART_PRCTR
                        If lRow.PART_PRCTR <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "PPRCTR")
                            oCriteria.SetValue("CHARACTER", lRow.PART_PRCTR)
                        End If
                        ' FUNC_AREA
                        If lRow.FUNC_AREA <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", pFKBERNAME)
                            oCriteria.SetValue("CHARACTER", lSAPFormat.unpack(lRow.FUNC_AREA, 4))
                        End If
                        ' ZZHFMC3
                        '   If lRow.ZZHFMC3 <> "" Then
                        '     oCriteria.Append()
                        '     oCriteria.SetValue("ITEMNO_ACC", lCnt)
                        '     oCriteria.SetValue("FIELDNAME", "WWHC3")
                        '     oCriteria.SetValue("CHARACTER", lSAPFormat.unpack(lRow.ZZHFMC3, 3))
                        '   End If
                        If lRow.WBS <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "PSPNR")
                            oCriteria.SetValue("CHARACTER", lSAPWbsElement.GetPspnr(lRow.WBS))
                        End If
                        If lRow.MTART <> "" Then
                            oCriteria.Append()
                            oCriteria.SetValue("ITEMNO_ACC", lCnt)
                            oCriteria.SetValue("FIELDNAME", "MTART")
                            oCriteria.SetValue("CHARACTER", lRow.MTART)
                        End If
                    Else
                        oAccountGl.SetValue("COSTCENTER", lSAPFormat.unpack(lRow.KOSTL, 10))
                        ' oAccountGl.SetValue("MATERIAL", lSAPFormat.unpack(lRow.MATNR, 18))
                        oAccountGl.SetValue("MATERIAL", lRow.MATNR)
                        oAccountGl.SetValue("PLANT", lRow.WERKS)
                        oAccountGl.SetValue("VENDOR_NO", lSAPFormat.unpack(lRow.LIFNR, 10))
                        oAccountGl.SetValue("ORDERID", lSAPFormat.unpack(lRow.AUFNR, 12))
                        oAccountGl.SetValue("WBS_ELEMENT", lRow.WBS)
                        oAccountGl.SetValue("NETWORK", lSAPFormat.unpack(lRow.NETWORK, 12))
                        oAccountGl.SetValue("ACTIVITY", lSAPFormat.unpack(lRow.ACTIVITY, 4))
                        oAccountGl.SetValue("SALES_ORD", lSAPFormat.unpack(lRow.SALES_ORD, 10))
                        oAccountGl.SetValue("S_ORD_ITEM", lSAPFormat.unpack(lRow.S_ORD_ITEM, 6))
                    End If
                    ' check if this is a direct postin to a tax account
                    If lRow.Betrag = lRow.TaxAmount And lRow.TaxAmount <> 0 Then
                        oCurrencyAmount.Append()
                        oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                        oCurrencyAmount.SetValue("CURRENCY", pWAERS)
                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(CDbl(0), "0.00"))
                        lCnt = lCnt + 1
                        oAccountTax.Append()
                        oAccountTax.SetValue("ITEMNO_ACC", lCnt)
                        oAccountTax.SetValue("COND_KEY", "MWVS")
                        oAccountTax.SetValue("TAX_CODE", lRow.MWSKZ)
                        oCurrencyAmount.Append()
                        oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                        oCurrencyAmount.SetValue("CURRENCY", pWAERS)
                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lRow.TaxAmount, "0.00"))
                        oCurrencyAmount.SetValue("AMT_BASE", Format$(lRow.TaxAmount, "0.00"))
                    Else
                        oCurrencyAmount.Append()
                        oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                        oCurrencyAmount.SetValue("CURRENCY", pWAERS)
                        oCurrencyAmount.SetValue("CURR_TYPE", "00")
                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lRow.Betrag, "0.00"))
                        lCntSav = lCnt
                        Dim aSAPCalcTaxesFromGross As New SAPCalcTaxesFromGross(sapcon)
                        Dim lTaxSum As Double
                        Dim lTaxBase As Double
                        Dim lTaxLines As Integer
                        Dim oTAX_ITEM_OUT As IRfcTable
                        If lRow.MWSKZ <> "" Then
                            log.Debug("post - " & "calling aSAPCalcTaxesFromGross.getTaxAmount")
                            oTAX_ITEM_OUT = aSAPCalcTaxesFromGross.getTaxAmount(pBUKRS, lRow.MWSKZ, pWAERS, pBUDAT, lRow.Betrag, lRow.TXJCD)
                            lTaxLines = oTAX_ITEM_OUT.Count
                            ' calculate the taxsum
                            lTaxSum = 0
                            For i As Integer = 0 To lTaxLines - 1
                                lTaxSum = lTaxSum + oTAX_ITEM_OUT(i).GetDouble("FWSTE")
                            Next i
                            lTaxBase = lRow.Betrag - lTaxSum
                            log.Debug("post - " & "lTaxLines=" & CStr(lTaxLines) & " lTaxSum=" & CStr(lTaxSum) & " lTaxBase=" & CStr(lTaxBase))
                            ' change the ammount of the row to the net value
                            oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lTaxBase, "0.00"))
                            ' add the tax positions
                            If lTaxSum <> 0 Or lTaxLines > 1 Then
                                For i As Integer = 0 To lTaxLines - 1
                                    lCnt = lCnt + 1
                                    oAccountTax.Append()
                                    oAccountTax.SetValue("ITEMNO_ACC", lCnt)
                                    oAccountTax.SetValue("COND_KEY", oTAX_ITEM_OUT(i).GetValue("KSCHL"))
                                    oAccountTax.SetValue("TAX_CODE", oTAX_ITEM_OUT(i).GetValue("MWSKZ"))
                                    oAccountTax.SetValue("TAXJURCODE", oTAX_ITEM_OUT(i).GetValue("TXJCD"))
                                    oCurrencyAmount.Append()
                                    oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                                    oCurrencyAmount.SetValue("CURRENCY", pWAERS)
                                    oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(oTAX_ITEM_OUT(i).GetDouble("FWSTE"), "0.00"))
                                    ' oCurrencyAmount.SetValue("AMT_BASE", Format$(lTaxBase, "0.00"))
                                    oCurrencyAmount.SetValue("AMT_BASE", Format$(oTAX_ITEM_OUT(i).GetDouble("FWBAS"), "0.00"))
                                Next i
                            End If
                        End If
                        If lRow.BETR2 <> 0 Then
                            lCnt = lCntSav
                            oCurrencyAmount.Append()
                            oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                            oCurrencyAmount.SetValue("CURR_TYPE", lRow.CURRTYP2)
                            oCurrencyAmount.SetValue("CURRENCY", lRow.WAERS2)
                            oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lRow.BETR2, "0.00"))
                            If lRow.MWSKZ <> "" Then
                                log.Debug("post - " & "calling aSAPCalcTaxesFromGross.getTaxAmount for BETR2")
                                oTAX_ITEM_OUT = aSAPCalcTaxesFromGross.getTaxAmount(pBUKRS, lRow.MWSKZ, lRow.WAERS2, pBUDAT, lRow.BETR2, lRow.TXJCD)
                                lTaxLines = oTAX_ITEM_OUT.Count
                                ' calculate the taxsum
                                lTaxSum = 0
                                For i As Integer = 0 To lTaxLines - 1
                                    lTaxSum = lTaxSum + oTAX_ITEM_OUT(i).GetDouble("FWSTE")
                                Next i
                                lTaxBase = lRow.BETR2 - lTaxSum
                                log.Debug("post - " & "lTaxLines=" & CStr(lTaxLines) & " lTaxSum=" & CStr(lTaxSum) & " lTaxBase=" & CStr(lTaxBase))
                                ' change the ammount of the row to the net value
                                oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lTaxBase, "0.00"))
                                ' add the tax positions
                                If lTaxSum <> 0 Or lTaxLines > 1 Then
                                    For i As Integer = 0 To lTaxLines - 1
                                        lCnt = lCnt + 1
                                        oCurrencyAmount.Append()
                                        oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                                        oCurrencyAmount.SetValue("CURR_TYPE", lRow.CURRTYP2)
                                        oCurrencyAmount.SetValue("CURRENCY", lRow.WAERS2)
                                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(oTAX_ITEM_OUT(i).GetDouble("FWSTE"), "0.00"))
                                        ' oCurrencyAmount.SetValue("AMT_BASE", Format$(lTaxBase, "0.00"))
                                        oCurrencyAmount.SetValue("AMT_BASE", Format$(oTAX_ITEM_OUT(i).GetDouble("FWBAS"), "0.00"))
                                    Next i
                                End If
                            End If
                        End If
                        If lRow.BETR3 <> 0 Then
                            lCnt = lCntSav
                            oCurrencyAmount.Append()
                            oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                            oCurrencyAmount.SetValue("CURR_TYPE", lRow.CURRTYP3)
                            oCurrencyAmount.SetValue("CURRENCY", lRow.WAERS3)
                            oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lRow.BETR3, "0.00"))
                            If lRow.MWSKZ <> "" Then
                                log.Debug("post - " & "calling aSAPCalcTaxesFromGross.getTaxAmount for BETR3")
                                oTAX_ITEM_OUT = aSAPCalcTaxesFromGross.getTaxAmount(pBUKRS, lRow.MWSKZ, lRow.WAERS3, pBUDAT, lRow.BETR3, lRow.TXJCD)
                                lTaxLines = oTAX_ITEM_OUT.Count
                                ' calculate the taxsum
                                lTaxSum = 0
                                For i As Integer = 0 To lTaxLines - 1
                                    lTaxSum = lTaxSum + oTAX_ITEM_OUT(i).GetDouble("FWSTE")
                                Next i
                                lTaxBase = lRow.BETR2 - lTaxSum
                                log.Debug("post - " & "lTaxLines=" & CStr(lTaxLines) & " lTaxSum=" & CStr(lTaxSum) & " lTaxBase=" & CStr(lTaxBase))
                                ' change the ammount of the row to the net value
                                oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lTaxBase, "0.00"))
                                ' add the tax positions
                                If lTaxSum <> 0 Or lTaxLines > 1 Then
                                    For i As Integer = 0 To lTaxLines - 1
                                        lCnt = lCnt + 1
                                        oCurrencyAmount.Append()
                                        oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                                        oCurrencyAmount.SetValue("CURR_TYPE", lRow.CURRTYP3)
                                        oCurrencyAmount.SetValue("CURRENCY", lRow.WAERS3)
                                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(oTAX_ITEM_OUT(i).GetDouble("FWSTE"), "0.00"))
                                        ' oCurrencyAmount.SetValue("AMT_BASE", Format$(lTaxBase, "0.00"))
                                        oCurrencyAmount.SetValue("AMT_BASE", Format$(oTAX_ITEM_OUT(i).GetDouble("FWBAS"), "0.00"))
                                    Next i
                                End If
                            End If
                        End If
                        If lRow.BETR4 <> 0 Then
                            lCnt = lCntSav
                            oCurrencyAmount.Append()
                            oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                            oCurrencyAmount.SetValue("CURR_TYPE", lRow.CURRTYP4)
                            oCurrencyAmount.SetValue("CURRENCY", lRow.WAERS4)
                            oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lRow.BETR4, "0.00"))
                            If lRow.MWSKZ <> "" Then
                                log.Debug("post - " & "calling aSAPCalcTaxesFromGross.getTaxAmount for BETR4")
                                oTAX_ITEM_OUT = aSAPCalcTaxesFromGross.getTaxAmount(pBUKRS, lRow.MWSKZ, lRow.WAERS4, pBUDAT, lRow.BETR4, lRow.TXJCD)
                                lTaxLines = oTAX_ITEM_OUT.Count
                                ' calculate the taxsum
                                lTaxSum = 0
                                For i As Integer = 0 To lTaxLines - 1
                                    lTaxSum = lTaxSum + oTAX_ITEM_OUT(i).GetDouble("FWSTE")
                                Next i
                                lTaxBase = lRow.BETR2 - lTaxSum
                                log.Debug("post - " & "lTaxLines=" & CStr(lTaxLines) & " lTaxSum=" & CStr(lTaxSum) & " lTaxBase=" & CStr(lTaxBase))
                                ' change the ammount of the row to the net value
                                oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lTaxBase, "0.00"))
                                ' add the tax positions
                                If lTaxSum <> 0 Or lTaxLines > 1 Then
                                    For i As Integer = 0 To lTaxLines - 1
                                        lCnt = lCnt + 1
                                        oCurrencyAmount.Append()
                                        oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                                        oCurrencyAmount.SetValue("CURR_TYPE", lRow.CURRTYP4)
                                        oCurrencyAmount.SetValue("CURRENCY", lRow.WAERS4)
                                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(oTAX_ITEM_OUT(i).GetDouble("FWSTE"), "0.00"))
                                        ' oCurrencyAmount.SetValue("AMT_BASE", Format$(lTaxBase, "0.00"))
                                        oCurrencyAmount.SetValue("AMT_BASE", Format$(oTAX_ITEM_OUT(i).GetDouble("FWBAS"), "0.00"))
                                    Next i
                                End If
                            End If
                        End If
                    End If
                End If
                If lRow.ACCTYPE = "D" Or lRow.ACCTYPE = "C" Then
                    log.Debug("post - " & "adding oAccountReceivable ITEMNO_ACC=" & CStr(lCnt) & " CUSTOMER=" & lSAPFormat.unpack(lRow.NEWKO, 10))
                    oAccountReceivable.Append()
                    oAccountReceivable.SetValue("ITEMNO_ACC", lCnt)
                    oAccountReceivable.SetValue("CUSTOMER", lSAPFormat.unpack(lRow.NEWKO, 10))
                    If lRow.GL_ACCOUNT <> "" Then
                        oAccountReceivable.SetValue("GL_ACCOUNT", lSAPFormat.unpack(lRow.GL_ACCOUNT, 10))
                    End If
                    oAccountReceivable.SetValue("ITEM_TEXT", lRow.SGTXT)
                    oAccountReceivable.SetValue("TAX_CODE", lRow.MWSKZ)
                    oAccountReceivable.SetValue("PMNTTRMS", lRow.PMNTTRMS)
                    oAccountReceivable.SetValue("PMNT_BLOCK", lRow.PMNT_BLOCK)
                    oAccountReceivable.SetValue("ALLOC_NMBR", lRow.ALLOC_NMBR)
                    oAccountReceivable.SetValue("REF_KEY_3", lRow.REF_KEY_3)
                    oAccountReceivable.SetValue("PARTNER_BK", lRow.PARTNER_BK)
                    If lRow.BLINE_DATE <> "" Then
                        oAccountReceivable.SetValue("BLINE_DATE", CDate(lRow.BLINE_DATE))
                    End If
                    If lRow.BUS_AREA <> "" Then
                        oAccountReceivable.SetValue("BUS_AREA", lSAPFormat.unpack(lRow.BUS_AREA, 4))
                    End If
                    If lRow.PRCTR <> "" Then
                        oAccountReceivable.SetValue("PROFIT_CTR", lRow.PRCTR)
                    End If
                    If lRow.SP_GL_IND <> "" Then
                        oAccountReceivable.SetValue("SP_GL_IND", lRow.SP_GL_IND)
                    End If
                    ' Business Place
                    If lRow.BUPLA <> "" Then
                        oExtension2.Append()
                        oExtension2.SetValue("STRUCTURE", "ZFI_BAPIEXT_BUPLA")
                        oExtension2.SetValue("VALUEPART1", lSAPFormat.unpack(CStr(lCnt), 10))
                        oExtension2.SetValue("VALUEPART2", lSAPFormat.unpack(lRow.BUPLA, 4))
                    End If
                    oCurrencyAmount.Append()
                    oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                    oCurrencyAmount.SetValue("CURRENCY", pWAERS)
                    oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lRow.Betrag, "0.00"))
                    If lRow.BETR2 <> 0 Then
                        oCurrencyAmount.Append()
                        oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                        oCurrencyAmount.SetValue("CURR_TYPE", lRow.CURRTYP2)
                        oCurrencyAmount.SetValue("CURRENCY", lRow.WAERS2)
                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lRow.BETR2, "0.00"))
                    End If
                    If lRow.BETR3 <> 0 Then
                        oCurrencyAmount.Append()
                        oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                        oCurrencyAmount.SetValue("CURR_TYPE", lRow.CURRTYP3)
                        oCurrencyAmount.SetValue("CURRENCY", lRow.WAERS3)
                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lRow.BETR3, "0.00"))
                    End If
                    If lRow.BETR4 <> 0 Then
                        oCurrencyAmount.Append()
                        oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                        oCurrencyAmount.SetValue("CURR_TYPE", lRow.CURRTYP4)
                        oCurrencyAmount.SetValue("CURRENCY", lRow.WAERS4)
                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lRow.BETR4, "0.00"))
                    End If
                End If
                If lRow.ACCTYPE = "K" Or lRow.ACCTYPE = "V" Then
                    log.Debug("post - " & "adding oAccountPayable ITEMNO_ACC=" & CStr(lCnt) & " VENDOR_NO=" & lSAPFormat.unpack(lRow.NEWKO, 10))
                    oAccountPayable.Append()
                    oAccountPayable.SetValue("ITEMNO_ACC", lCnt)
                    oAccountPayable.SetValue("VENDOR_NO", lSAPFormat.unpack(lRow.NEWKO, 10))
                    If lRow.GL_ACCOUNT <> "" Then
                        oAccountPayable.SetValue("GL_ACCOUNT", lSAPFormat.unpack(lRow.GL_ACCOUNT, 10))
                    End If
                    oAccountPayable.SetValue("ITEM_TEXT", lRow.SGTXT)
                    oAccountPayable.SetValue("TAX_CODE", lRow.MWSKZ)
                    oAccountPayable.SetValue("PMNTTRMS", lRow.PMNTTRMS)
                    oAccountPayable.SetValue("PMNT_BLOCK", lRow.PMNT_BLOCK)
                    oAccountPayable.SetValue("ALLOC_NMBR", lRow.ALLOC_NMBR)
                    oAccountPayable.SetValue("REF_KEY_3", lRow.REF_KEY_3)
                    oAccountPayable.SetValue("PARTNER_BK", lRow.PARTNER_BK)
                    If lRow.BLINE_DATE <> "" Then
                        oAccountPayable.SetValue("BLINE_DATE", CDate(lRow.BLINE_DATE))
                    End If
                    If lRow.BUS_AREA <> "" Then
                        oAccountPayable.SetValue("BUS_AREA", lSAPFormat.unpack(lRow.BUS_AREA, 4))
                    End If
                    If lRow.PRCTR <> "" Then
                        oAccountPayable.SetValue("PROFIT_CTR", lRow.PRCTR)
                    End If
                    If lRow.SP_GL_IND <> "" Then
                        oAccountPayable.SetValue("SP_GL_IND", lRow.SP_GL_IND)
                    End If
                    ' Business Place
                    If lRow.BUPLA <> "" Then
                        oExtension2.Append()
                        oExtension2.SetValue("STRUCTURE", "ZFI_BAPIEXT_BUPLA")
                        oExtension2.SetValue("VALUEPART1", lSAPFormat.unpack(CStr(lCnt), 10))
                        oExtension2.SetValue("VALUEPART2", lSAPFormat.unpack(lRow.BUPLA, 4))
                    End If
                    oCurrencyAmount.Append()
                    oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                    oCurrencyAmount.SetValue("CURRENCY", pWAERS)
                    oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lRow.Betrag, "0.00"))
                    If lRow.BETR2 <> 0 Then
                        oCurrencyAmount.Append()
                        oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                        oCurrencyAmount.SetValue("CURR_TYPE", lRow.CURRTYP2)
                        oCurrencyAmount.SetValue("CURRENCY", lRow.WAERS2)
                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lRow.BETR2, "0.00"))
                    End If
                    If lRow.BETR3 <> 0 Then
                        oCurrencyAmount.Append()
                        oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                        oCurrencyAmount.SetValue("CURR_TYPE", lRow.CURRTYP3)
                        oCurrencyAmount.SetValue("CURRENCY", lRow.WAERS3)
                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lRow.BETR3, "0.00"))
                    End If
                    If lRow.BETR4 <> 0 Then
                        oCurrencyAmount.Append()
                        oCurrencyAmount.SetValue("ITEMNO_ACC", lCnt)
                        oCurrencyAmount.SetValue("CURR_TYPE", lRow.CURRTYP4)
                        oCurrencyAmount.SetValue("CURRENCY", lRow.WAERS4)
                        oCurrencyAmount.SetValue("AMT_DOCCUR", Format$(lRow.BETR4, "0.00"))
                    End If
                End If
            Next lRow
            ' call the BAPI
            log.Debug("post - " & "invoking " & oRfcFunction.Metadata.Name)
            oRfcFunction.Invoke(destination)
            log.Debug("post - " & "oRETURN.Count=" & CStr(oRETURN.Count))
            If oRETURN.Count > 0 Then
                If oRETURN(0).GetValue("TYPE") = "S" Then
                    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                    If Not pTest Then
                        log.Debug("post - " & "calling aSAPBapiTranctionCommit.commit()")
                        aSAPBapiTranctionCommit.commit()
                    End If
                    post = oRETURN(0).GetValue("MESSAGE")
                Else
                    For i As Integer = 0 To oRETURN.Count - 1
                        post = post & ";" & oRETURN(i).GetValue("MESSAGE")
                    Next i
                End If
            Else
                log.Debug("post - " & "Error: No Return message from SAP")
                post = "Error: No Return message from SAP"
            End If
        Catch Ex As System.Exception
            log.Error("commit - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAcctngDocument")
            post = "Error: Exception in posting"
        Finally
            log.Debug("post - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
