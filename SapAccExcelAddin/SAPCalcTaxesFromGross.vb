' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPCalcTaxesFromGross

    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        sapcon = aSapCon
        destination = aSapCon.getDestination()
        sapcon.checkCon()
        oRfcFunction = destination.Repository.CreateFunction("CALCULATE_TAXES_GROSS")
    End Sub

    Public Function getTaxAmount(pBUKRS As String, pMWSKZ As String, pWAERS As String, pBUDAT As Date,
                                 pWRBTR As Double, Optional ByVal pTXJCD As String = "") As IRfcTable
        Try
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
            oRfcFunction.Invoke(destination)
            getTaxAmount = oTAX_ITEM_OUT
        Catch ex As Exception
            MsgBox("Exception in getTaxAmount! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCalcTaxesFromGross")
            getTaxAmount = Nothing
        End Try
    End Function

End Class
