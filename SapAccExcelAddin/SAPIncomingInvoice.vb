' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPIncomingInvoice

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private SapCon As SapConHelper

    Sub New(ByRef aSapCon As SapConHelper)
        SapCon = aSapCon
        aSapCon.getDestination(destination)
        log.Debug("New - " & "creating Function BAPI_INCOMINGINVOICE_GETDETAIL")
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_INCOMINGINVOICE_GETDETAIL")
            log.Debug("New - " & "oRfcFunction.Metadata.Name=" & oRfcFunction.Metadata.Name)
        Catch ex As Exception
            oRfcFunction = Nothing
            log.Warn("New - Exception=" & ex.ToString)
        End Try
    End Sub

    Public Function getDetail(pInvoiceDocNumber As String, pFiscalYear As String) As Dictionary(Of String, BInvRec)
        Dim aFXrate As Double
        Dim aRetMessage As String = ""
        Dim aBInv As New BInv

        SapCon.checkCon()
        Try
            log.Debug("post - " & "Getting Function parameters")
            Dim oHEADERDATA As IRfcStructure = oRfcFunction.GetStructure("HEADERDATA")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            Dim oITEMDATA As IRfcTable = oRfcFunction.GetTable("ITEMDATA")
            Dim oACCOUNTINGDATA As IRfcTable = oRfcFunction.GetTable("ACCOUNTINGDATA")
            Dim oGLACCOUNTDATA As IRfcTable = oRfcFunction.GetTable("GLACCOUNTDATA")

            oRETURN.Clear()
            oITEMDATA.Clear()
            oACCOUNTINGDATA.Clear()
            oGLACCOUNTDATA.Clear()
            log.Debug("post - " & "setting header values")

            oRfcFunction.SetValue("INVOICEDOCNUMBER", pInvoiceDocNumber)
            oRfcFunction.SetValue("FISCALYEAR", pFiscalYear)

            log.Debug("getDetail - " & "invoking " & oRfcFunction.Metadata.Name)
            oRfcFunction.Invoke(destination)
            log.Debug("getDetail - " & "oRETURN.Count=" & CStr(oRETURN.Count))

            If oRETURN.Count = 0 Then
                If CStr(oHEADERDATA.GetValue("EXCH_RATE")) = "0,00000" And CStr(oHEADERDATA.GetValue("EXCH_RATE_V")) <> "0,00000" Then
                    aFXrate = 1 / CDbl(oHEADERDATA.GetValue("EXCH_RATE_V"))
                ElseIf CStr(oHEADERDATA.GetValue("EXCH_RATE")) <> "0,00000" Then
                    aFXrate = CDbl(oHEADERDATA.GetValue("EXCH_RATE"))
                Else
                    aFXrate = 1
                End If
                Dim iItem As Integer
                For iI As Integer = 0 To oITEMDATA.Count - 1
                    ' add the Invoice items
                    aBInv.addBInvItem("OK", CStr(oHEADERDATA.GetValue("INV_DOC_NO")), CStr(oHEADERDATA.GetValue("FISC_YEAR")), CStr(oHEADERDATA.GetValue("DOC_TYPE")),
                            CStr(oHEADERDATA.GetValue("DOC_DATE")), CStr(oHEADERDATA.GetValue("PSTNG_DATE")), CStr(oHEADERDATA.GetValue("REF_DOC_NO")),
                            CStr(oHEADERDATA.GetValue("COMP_CODE")), CStr(oHEADERDATA.GetValue("CURRENCY")), CStr(aFXrate),
                            CStr(oHEADERDATA.GetValue("HEADER_TXT")), CStr(oHEADERDATA.GetValue("DIFF_INV")),
                            CStr(oITEMDATA(iI).GetValue("INVOICE_DOC_ITEM")), CStr(oITEMDATA(iI).GetValue("PO_NUMBER")), CStr(oITEMDATA(iI).GetValue("PO_ITEM")), CStr(oITEMDATA(iI).GetValue("ITEM_TEXT")))
                    iItem = iI
                Next iI
                Dim aPosCnt As Integer = 1
                For i As Integer = 0 To oACCOUNTINGDATA.Count - 1
                    ' add the Invoice items
                    aBInv.addBInvAcc("OK", CStr(oHEADERDATA.GetValue("INV_DOC_NO")), CStr(oHEADERDATA.GetValue("FISC_YEAR")), CStr(oHEADERDATA.GetValue("DOC_TYPE")),
                            CStr(oHEADERDATA.GetValue("DOC_DATE")), CStr(oHEADERDATA.GetValue("PSTNG_DATE")), CStr(oHEADERDATA.GetValue("REF_DOC_NO")),
                            CStr(oHEADERDATA.GetValue("COMP_CODE")), CStr(oHEADERDATA.GetValue("CURRENCY")), CStr(aFXrate),
                            CStr(oHEADERDATA.GetValue("HEADER_TXT")), CStr(oHEADERDATA.GetValue("DIFF_INV")),
                            CStr(oACCOUNTINGDATA(i).GetValue("INVOICE_DOC_ITEM")), CStr(aPosCnt),
                            aBInv.getITEM_TEXT(CStr(oHEADERDATA.GetValue("INV_DOC_NO")), CStr(oHEADERDATA.GetValue("FISC_YEAR")), CStr(oACCOUNTINGDATA(i).GetValue("INVOICE_DOC_ITEM"))),
                            aBInv.getPO_NUMBER(CStr(oHEADERDATA.GetValue("INV_DOC_NO")), CStr(oHEADERDATA.GetValue("FISC_YEAR")), CStr(oACCOUNTINGDATA(i).GetValue("INVOICE_DOC_ITEM"))),
                            aBInv.getPO_ITEM(CStr(oHEADERDATA.GetValue("INV_DOC_NO")), CStr(oHEADERDATA.GetValue("FISC_YEAR")), CStr(oACCOUNTINGDATA(i).GetValue("INVOICE_DOC_ITEM"))),
                            CStr(oACCOUNTINGDATA(i).GetValue("ITEM_AMOUNT")), CStr(oACCOUNTINGDATA(i).GetValue("QUANTITY")),
                            CStr(oACCOUNTINGDATA(i).GetValue("PO_UNIT")), CStr(oACCOUNTINGDATA(i).GetValue("GL_ACCOUNT")), CStr(oACCOUNTINGDATA(i).GetValue("COSTCENTER")),
                            CStr(oACCOUNTINGDATA(i).GetValue("NETWORK")), CStr(oACCOUNTINGDATA(i).GetValue("ACTIVITY")), CStr(oACCOUNTINGDATA(i).GetValue("WBS_ELEM")),
                            CStr(oACCOUNTINGDATA(i).GetValue("ASSET_NO")), CStr(oACCOUNTINGDATA(i).GetValue("SUB_NUMBER")), CStr(oACCOUNTINGDATA(i).GetValue("ORDERID")))
                    aPosCnt += 1
                Next i
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    aRetMessage = aRetMessage & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
                ' add the message
                aBInv.addBInvMessage(aRetMessage, pInvoiceDocNumber, pFiscalYear, "")
            End If
            getDetail = aBInv.aBInv
        Catch ex As Exception
            MsgBox("Exception in getDetail! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPIncomingInvoice")
            getDetail = Nothing
            log.Error("getDetail - " & "ex= " & ex.ToString)
        End Try
    End Function
End Class
