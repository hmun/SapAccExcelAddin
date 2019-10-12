' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports System.Reflection
Imports System.Diagnostics

Public Class SapGeneral
    Const cVersion As String = "1.0.3.5"
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Function checkVersion() As Integer
        Dim aCws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aFromVersion As String
        Dim aToVersion As String
        Dim assembly As Assembly
        Dim fileVersionInfo As FileVersionInfo
        Dim aVersion As String
        Dim theVersion As Version

        aWB = Globals.SapAccAddIn.Application.ActiveWorkbook
        Try
            aCws = aWB.Worksheets("SAP-Con")
        Catch Exc As System.Exception
            MsgBox("No SAP-Con Sheet in current workbook. Check if the current workbook is a valid SAP Accounting Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapGeneral")
            checkVersion = False
            log.Error("checkVersion - Exception=" & Exc.ToString)
            Exit Function
        End Try
        log.Debug("checkVersion - " & "reading Versions")
        aFromVersion = aCws.Cells(15, 2).Value
        log.Debug("checkVersion - " & "aFromVersion=" & CStr(aFromVersion))
        aToVersion = aCws.Cells(16, 2).Value
        log.Debug("checkVersion - " & "aToVersion=" & CStr(aToVersion))

        Try
            log.Debug("checkVersion - " & "reading assembly versions")
            assembly = System.Reflection.Assembly.GetExecutingAssembly()
            fileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location)
            aVersion = fileVersionInfo.ProductVersion
            log.Debug("checkVersion - " & "aVersion=" & CStr(aVersion))
        Catch Exc As System.Exception
            log.Debug("checkVersion - " & "failed to read assembly versions using default")
            aVersion = cVersion
            log.Debug("checkVersion - " & "aVersion=" & CStr(aVersion))
        End Try
        log.Debug("checkVersion - " & "using default version")
        aVersion = cVersion
        log.Debug("checkVersion - " & "aVersion=" & CStr(aVersion))
        If aVersion > aToVersion Or aVersion < aFromVersion Then
            ' try Publish Version
            log.Debug("checkVersion - " & "version invalid")
            MsgBox("The Version of the Excel-Template is not valid for this Add-In. Please use a Template that is valid for version " & aVersion,
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapGeneral")
            checkVersion = False
            Exit Function
        End If

        log.Debug("checkVersion - " & "version OK")
        checkVersion = True
    End Function

    Public Function checkVersionInSAP(pSapCon As SapCon) As Integer
        Dim aSAPZ_BC_EXCEL_ADDIN_VERS_CHK As New SAPZ_BC_EXCEL_ADDIN_VERS_CHK(pSapCon)
        Dim assembly As Assembly
        Dim assemblyNames As String()
        Dim aAddIn As String
        Dim aRet As Integer

        checkVersionInSAP = True
        aAddIn = ""
        Try
            log.Debug("checkVersionInSAP - " & "reading assemblyName")
            assembly = System.Reflection.Assembly.GetExecutingAssembly()
            assemblyNames = assembly.GetName().ToString.Split(New Char() {","c})
            aAddIn = assemblyNames(0)
            log.Debug("checkVersionInSAP - " & "aAddIn=" & CStr(aAddIn))
        Catch Exc As System.Exception
            MsgBox("Exception: " & Exc.Message,
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapGeneral")
            checkVersionInSAP = False
            log.Error("checkVersionInSAP - " & "failed to read assemblyName")
            Exit Function
        End Try
        log.Debug("checkVersionInSAP - " & "calling aSAPZ_BC_EXCEL_ADDIN_VERS_CHK.checkVersion")
        aRet = aSAPZ_BC_EXCEL_ADDIN_VERS_CHK.checkVersion(aAddIn, cVersion)
        If aRet <> 0 Then
            MsgBox("The Version " & cVersion & " of the Add-In " & aAddIn & " is not allowed in this SAP-System!",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapGeneral")
            log.Debug("checkVersionInSAP - " & "The Version " & cVersion & " of the Add-In " & aAddIn & " is not allowed in this SAP-System!")
            checkVersionInSAP = False
        End If
    End Function

End Class
