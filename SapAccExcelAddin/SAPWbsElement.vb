Imports SAP.Middleware.Connector

Public Class SAPWbsElement

    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        sapcon = aSapCon
        destination = aSapCon.getDestination()
        oRfcFunction = destination.Repository.CreateFunction("Z_CO_PS_PSP_INTERNAL")
    End Sub

    Public Function GetPspnr(pPOSID As String) As String
        sapcon.checkCon()
        Try
            oRfcFunction.SetValue("I_POSID", pPOSID)
            oRfcFunction.Invoke(destination)
            GetPspnr = oRfcFunction.GetValue("E_PSPNR")
            Exit Function
        Catch ex As Exception
            MsgBox("Exception in GetPspnr! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWbsElement")
            GetPspnr = "Fehler"
        End Try
    End Function

End Class
