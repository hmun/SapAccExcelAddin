' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPZFI_CHECK_F_BKPF_BUK

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapConHelper

    Sub New(ByRef aSapCon As SapConHelper)
        sapcon = aSapCon
        aSapCon.getDestination(destination)
        log.Debug("New - " & "creating Function ZFI_CHECK_F_BKPF_BUK")
        Try
            oRfcFunction = destination.Repository.CreateFunction("ZFI_CHECK_F_BKPF_BUK")
            log.Debug("New - " & "oRfcFunction.Metadata.Name=" & oRfcFunction.Metadata.Name)
        Catch ex As Exception
            oRfcFunction = Nothing
            log.Warn("New - Exception=" & ex.ToString)
        End Try
    End Sub

    Public Function checkAuthority(pBUKRS As String) As Integer
        sapcon.checkCon()
        If oRfcFunction Is Nothing Then
            ' for systems that do not contain ZFI_CHECK_F_BKPF_BUK we can not check the authorization
            checkAuthority = 2
            log.Debug("checkAuthority - " & "oRfcFunction is Nothing, skiping check. checkAuthority=" & checkAuthority)
        Else
            Try
                log.Debug("checkAuthority - " & "Setting Function parameters")
                oRfcFunction.SetValue("I_BUKRS", pBUKRS)
                oRfcFunction.Invoke(destination)
                log.Debug("checkAuthority - " & "invoking " & oRfcFunction.Metadata.Name)
                checkAuthority = oRfcFunction.GetValue("E_SUBRC")
                log.Debug("checkAuthority - " & "checkAuthority=" & CStr(checkAuthority))
            Catch ex As Exception
                MsgBox("Exception in checkAuthority! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPZFI_CHECK_F_BKPF_BUK")
                checkAuthority = 8
                log.Error("checkAuthority - " & "ex= " & ex.ToString & ", checkAuthority=" & CStr(checkAuthority))
            End Try
        End If
        Exit Function
    End Function

End Class
