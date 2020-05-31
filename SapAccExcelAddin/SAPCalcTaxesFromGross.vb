' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPCalcTaxesFromGross

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        sapcon = aSapCon
        aSapCon.getDestination(destination)
        sapcon.checkCon()
        log.Debug("New - " & "creating Function CALCULATE_TAXES_GROSS")
        Try
            oRfcFunction = destination.Repository.CreateFunction("CALCULATE_TAXES_GROSS")
            log.Debug("New - " & "oRfcFunction.Metadata.Name=" & oRfcFunction.Metadata.Name)
        Catch Exc As System.Exception
            log.Error("New - Exception=" & Exc.ToString)
        End Try
    End Sub

    Public Function getTaxAmount(pBUKRS As String, pMWSKZ As String, pWAERS As String, pBUDAT As Date,
                                 pWRBTR As Double, Optional ByVal pTXJCD As String = "") As IRfcTable
        Try
            log.Debug("getTaxAmount - " & "Getting Function parameters")
            Dim oTAX_ITEM_IN As IRfcTable = oRfcFunction.GetTable("TAX_ITEM_IN")
            Dim oTAX_ITEM_OUT As IRfcTable = oRfcFunction.GetTable("TAX_ITEM_OUT")
            oTAX_ITEM_IN.Clear()
            oTAX_ITEM_OUT.Clear()

            oTAX_ITEM_IN.Append()
            oTAX_ITEM_IN.SetValue("BUKRS", pBUKRS)
            oTAX_ITEM_IN.SetValue("MWSKZ", pMWSKZ)
            oTAX_ITEM_IN.SetValue("TXJCD", pTXJCD)
            oTAX_ITEM_IN.SetValue("WAERS", pWAERS)
            oTAX_ITEM_IN.SetValue("BUDAT", pBUDAT)
            oTAX_ITEM_IN.SetValue("WRBTR", Format$(pWRBTR, "0.00"))
            log.Debug("getTaxAmount - " & "invoking " & oRfcFunction.Metadata.Name)
            oRfcFunction.Invoke(destination)
            getTaxAmount = oTAX_ITEM_OUT
        Catch ex As Exception
            MsgBox("Exception in getTaxAmount! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCalcTaxesFromGross")
            log.Error("getTaxAmount - Exception=" & ex.ToString)
            getTaxAmount = Nothing
        End Try
    End Function

End Class
