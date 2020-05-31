' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports System.Configuration
Imports System.Environment
Imports System.Uri
Imports System.IO
Imports SAPCommon

Public Class MigHelper
    Public mig As SAPCommon.Migration
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub New(uselocal As Boolean)
        Dim aWs As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim configFile As String = ""
        ' Check for local rules first
        If Not uselocal Then
            Dim assemblyName As System.Reflection.AssemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName()
            Dim assembly As String = assemblyName.Name
            Dim appData As String = GetFolderPath(Environment.SpecialFolder.ApplicationData)
            configFile = Uri.UnescapeDataString(appData & "\SapExcel\" & assembly & "\inv_mig_rules.config")
            log.Debug("New - " & "looking for config file=" & configFile)
            If Not System.IO.File.Exists(configFile) Then
                appData = GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                configFile = Uri.UnescapeDataString(appData & "\SapExcel\" & assembly & "\inv_mig_rules.config")
                log.Debug("New - " & "looking for config file=" & configFile)
                If Not System.IO.File.Exists(configFile) Then
                    appData = New Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).AbsolutePath
                    appData = Path.GetDirectoryName(appData)
                    configFile = Uri.UnescapeDataString(appData & "\inv_mig_rules.config")
                    log.Debug("New - " & "looking for config file=" & configFile)
                    If Not System.IO.File.Exists(configFile) Then
                        configFile = ""
                    End If
                End If
            End If
        End If
        ' setup the migration engine
        If Not configFile = "" Then
            log.Debug("New - " & "found config file=" & configFile)
            mig = New SAPCommon.Migration(configFile)
        Else
            log.Debug("New - " & "No config file found looking for config worksheets")
            mig = New SAPCommon.Migration()
            ' try to read the rules from the excel workbook
            Dim i As Integer
            aWB = Globals.SapAccAddIn.Application.ActiveWorkbook
            Try
                aWs = aWB.Worksheets("Rules")
                i = 2
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    mig.AddRule(CStr(aWs.Cells(i, 1).Value), CStr(aWs.Cells(i, 2).Value), CStr(aWs.Cells(i, 3).Value), CStr(aWs.Cells(i, 4).Value))
                    i += 1
                Loop
            Catch Exc As System.Exception
                MsgBox("No Rules Sheet in current workbook.",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting, MigHelper")
            End Try
            Try
                aWs = aWB.Worksheets("Constant")
                i = 2
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    mig.AddConstant(CStr(aWs.Cells(i, 1).Value), CStr(aWs.Cells(i, 2).Value), CStr(aWs.Cells(i, 3).Value))
                    i += 1
                Loop
            Catch Exc As System.Exception
                MsgBox("No Constant Sheet in current workbook.",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting, MigHelper")
            End Try
            Try
                aWs = aWB.Worksheets("Formula")
                i = 2
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    mig.AddFormula(CStr(aWs.Cells(i, 1).Value), CStr(aWs.Cells(i, 2).Value), CStr(aWs.Cells(i, 3).Value))
                    i += 1
                Loop
            Catch Exc As System.Exception
                MsgBox("No Formula Sheet in current workbook.",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting, MigHelper")
            End Try
            Try
                aWs = aWB.Worksheets("Mapping")
                i = 2
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    mig.AddMapping(CStr(aWs.Cells(i, 2).Value), CStr(aWs.Cells(i, 1).Value), CStr(aWs.Cells(i, 3).Value), CStr(aWs.Cells(i, 4).Value))
                    i += 1
                Loop
            Catch Exc As System.Exception
                MsgBox("No Formula Sheet in current workbook.",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting, MigHelper")
            End Try
        End If
    End Sub

    Sub saveToConfig()
        Dim assemblyName As System.Reflection.AssemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName()
        Dim assembly As String = assemblyName.Name
        Dim appData As String = GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        Dim configFile As String = appData & "\SapExcel\" & assembly & "\inv_mig_rules.config"
        Dim config As Configuration
        Dim configMap As New ExeConfigurationFileMap
        configMap.ExeConfigFilename = configFile
        config = TryCast(ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None), Configuration)
        config.Sections.Add("MigRules", mig.MRS)
        mig.MRS.SectionInformation.ForceSave = True
        config.Save(ConfigurationSaveMode.Full)
    End Sub

    Public Function makeDictForRules(ByRef pWs As Excel.Worksheet, pRow As Integer, pHeaderRow As Integer, pFromCol As Integer, pToCol As Integer) As Dictionary(Of String, SAPCommon.TField)
        Dim retDict As New Dictionary(Of String, SAPCommon.TField)
        Dim tfield As New SAPCommon.TField
        For j = pFromCol To pToCol
            If Not CStr(pWs.Cells(pHeaderRow, j).Value) = "" Then
                If mig.ContainsSource("P", CStr(pWs.Cells(pHeaderRow, j).Value)) Or
                   mig.ContainsSource("C", CStr(pWs.Cells(pHeaderRow, j).Value)) Then
                    tfield = New SAPCommon.TField(CStr(pWs.Cells(pHeaderRow, j).Value), CStr(pWs.Cells(pRow, j).Value))
                    retDict.Add(tfield.Name, tfield)
                End If
            End If
        Next
        makeDictForRules = retDict
    End Function
    Function makeDict(ByRef pWs As Excel.Worksheet, pRow As Integer, pHeaderRow As Integer, pFromCol As Integer, pToCol As Integer) As Dictionary(Of String, SAPCommon.TField)
        Dim retDict As New Dictionary(Of String, SAPCommon.TField)
        Dim tfield As New SAPCommon.TField
        For j = pFromCol To pToCol
            If Not CStr(pWs.Cells(pHeaderRow, j).Value) = "" Then
                tfield = New SAPCommon.TField(CStr(pWs.Cells(pHeaderRow, j).Value), CStr(pWs.Cells(pRow, j).Value))
                retDict.Add(tfield.Name, tfield)
            End If
        Next j
        makeDict = retDict
    End Function

End Class
