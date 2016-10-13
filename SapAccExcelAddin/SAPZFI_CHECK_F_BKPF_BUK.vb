Imports SAP.Middleware.Connector

Public Class SAPZFI_CHECK_F_BKPF_BUK

    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        sapcon = aSapCon
        destination = aSapCon.getDestination()
        oRfcFunction = destination.Repository.CreateFunction("ZFAGL_UPD_YMPNUM")
    End Sub

    Public Function checkAuthority(pBUKRS As String) As Integer
        sapcon.checkCon()
        Try
            oRfcFunction.SetValue("I_BUKRS", pBUKRS)
            oRfcFunction.Invoke(destination)
            checkAuthority = oRfcFunction.GetValue("E_RETURN")
        Catch ex As Exception
            MsgBox("Exception in checkAuthority! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPZFI_CHECK_F_BKPF_BUK")
            checkAuthority = 8
        End Try
        Exit Function
    End Function

End Class
