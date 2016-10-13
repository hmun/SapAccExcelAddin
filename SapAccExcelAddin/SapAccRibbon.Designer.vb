Partial Class SapAccRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.SapAcc = Me.Factory.CreateRibbonTab
        Me.Accounting = Me.Factory.CreateRibbonGroup
        Me.Logon = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.ButtonCheckAccDoc = Me.Factory.CreateRibbonButton
        Me.ButtonPostAccDoc = Me.Factory.CreateRibbonButton
        Me.SapAcc.SuspendLayout()
        Me.Accounting.SuspendLayout()
        Me.Logon.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapAcc
        '
        Me.SapAcc.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.SapAcc.Groups.Add(Me.Accounting)
        Me.SapAcc.Groups.Add(Me.Logon)
        Me.SapAcc.Label = "SAP Document"
        Me.SapAcc.Name = "SapAcc"
        '
        'Accounting
        '
        Me.Accounting.Items.Add(Me.ButtonCheckAccDoc)
        Me.Accounting.Items.Add(Me.ButtonPostAccDoc)
        Me.Accounting.Label = "Accounting"
        Me.Accounting.Name = "Accounting"
        '
        'Logon
        '
        Me.Logon.Items.Add(Me.ButtonLogon)
        Me.Logon.Items.Add(Me.ButtonLogoff)
        Me.Logon.Label = "Logon"
        Me.Logon.Name = "Logon"
        '
        'ButtonLogon
        '
        Me.ButtonLogon.Label = "SAP Logon"
        Me.ButtonLogon.Name = "ButtonLogon"
        '
        'ButtonLogoff
        '
        Me.ButtonLogoff.Label = "SAP Logoff"
        Me.ButtonLogoff.Name = "ButtonLogoff"
        '
        'ButtonCheckAccDoc
        '
        Me.ButtonCheckAccDoc.Label = "Check Acc Document"
        Me.ButtonCheckAccDoc.Name = "ButtonCheckAccDoc"
        '
        'ButtonPostAccDoc
        '
        Me.ButtonPostAccDoc.Label = "Post Acc Document"
        Me.ButtonPostAccDoc.Name = "ButtonPostAccDoc"
        '
        'SapAccRibbon
        '
        Me.Name = "SapAccRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapAcc)
        Me.SapAcc.ResumeLayout(False)
        Me.SapAcc.PerformLayout()
        Me.Accounting.ResumeLayout(False)
        Me.Accounting.PerformLayout()
        Me.Logon.ResumeLayout(False)
        Me.Logon.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapAcc As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Accounting As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Logon As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonCheckAccDoc As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostAccDoc As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property SapAccRibbon() As SapAccRibbon
        Get
            Return Me.GetRibbon(Of SapAccRibbon)()
        End Get
    End Property
End Class
