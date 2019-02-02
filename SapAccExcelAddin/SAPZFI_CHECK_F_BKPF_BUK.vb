﻿' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPZFI_CHECK_F_BKPF_BUK

    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        sapcon = aSapCon
        destination = aSapCon.getDestination()
        Try
            oRfcFunction = destination.Repository.CreateFunction("ZFI_CHECK_F_BKPF_BUK")
        Catch ex As Exception
            oRfcFunction = Nothing
        End Try
    End Sub

    Public Function checkAuthority(pBUKRS As String) As Integer
        sapcon.checkCon()
        If oRfcFunction Is Nothing Then
            ' for systems that do not contain ZFI_CHECK_F_BKPF_BUK we can not check the authorization
            checkAuthority = 2
        Else
            Try
                oRfcFunction.SetValue("I_BUKRS", pBUKRS)
                oRfcFunction.Invoke(destination)
                checkAuthority = oRfcFunction.GetValue("E_SUBRC")
            Catch ex As Exception
                MsgBox("Exception in checkAuthority! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPZFI_CHECK_F_BKPF_BUK")
                checkAuthority = 8
            End Try
        End If
        Exit Function
    End Function

End Class
