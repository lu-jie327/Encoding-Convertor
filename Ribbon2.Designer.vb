Partial Class Ribbon2
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
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.UTF8Btn = Me.Factory.CreateRibbonButton
        Me.UTF16LEBtn = Me.Factory.CreateRibbonButton
        Me.UTF16BEBtn = Me.Factory.CreateRibbonButton
        Me.SettingBtn = Me.Factory.CreateRibbonButton
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "Encoding"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.UTF8Btn)
        Me.Group1.Items.Add(Me.UTF16LEBtn)
        Me.Group1.Items.Add(Me.UTF16BEBtn)
        Me.Group1.Items.Add(Me.SettingBtn)
        Me.Group1.Label = "Convertor"
        Me.Group1.Name = "Group1"
        '
        'UTF8Btn
        '
        Me.UTF8Btn.Label = "To UTF-8"
        Me.UTF8Btn.Name = "UTF8Btn"
        '
        'UTF16LEBtn
        '
        Me.UTF16LEBtn.Label = "To UTF-16 LE"
        Me.UTF16LEBtn.Name = "UTF16LEBtn"
        '
        'UTF16BEBtn
        '
        Me.UTF16BEBtn.Label = "To UTF-16 BE"
        Me.UTF16BEBtn.Name = "UTF16BEBtn"
        '
        'SettingBtn
        '
        Me.SettingBtn.Label = "Setting"
        Me.SettingBtn.Name = "SettingBtn"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Ribbon2
        '
        Me.Name = "Ribbon2"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents UTF16LEBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents UTF16BEBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents OpenFileDialog1 As Windows.Forms.OpenFileDialog
    Friend WithEvents UTF8Btn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SettingBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon2() As Ribbon2
        Get
            Return Me.GetRibbon(Of Ribbon2)()
        End Get
    End Property
End Class
