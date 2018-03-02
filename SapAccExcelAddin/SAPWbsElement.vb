Imports SAP.Middleware.Connector

Public Class SAPWbsElement

    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        sapcon = aSapCon
        destination = aSapCon.getDestination()
        Try
            oRfcFunction = destination.Repository.CreateFunction("Z_CO_PS_PSP_INTERNAL")
        Catch ex As Exception
            oRfcFunction = Nothing
        End Try
    End Sub

    Public Function GetPspnr(pPOSID As String) As String
        If Not oRfcFunction Is Nothing Then
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
        Else
            GetPspnr = pPOSID
        End If
    End Function

End Class
