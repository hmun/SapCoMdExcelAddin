Partial Class SapCoMdRibbon
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SapCoMdRibbon))
        Me.SapCoMd = Me.Factory.CreateRibbonTab
        Me.SAPCostElement = Me.Factory.CreateRibbonGroup
        Me.SAPActivityType = Me.Factory.CreateRibbonGroup
        Me.SAPCostCenter = Me.Factory.CreateRibbonGroup
        Me.ButtonCCCreate = Me.Factory.CreateRibbonButton
        Me.ButtonCCChange = Me.Factory.CreateRibbonButton
        Me.SAPInternalOrder = Me.Factory.CreateRibbonGroup
        Me.SAPProfitCenter = Me.Factory.CreateRibbonGroup
        Me.SAPGLAccount = Me.Factory.CreateRibbonGroup
        Me.ButtonGLAccCreate = Me.Factory.CreateRibbonButton
        Me.SapCoMdLogon = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.ButtonCECreate = Me.Factory.CreateRibbonButton
        Me.ButtonCEChange = Me.Factory.CreateRibbonButton
        Me.SapCoMd.SuspendLayout()
        Me.SAPCostElement.SuspendLayout()
        Me.SAPCostCenter.SuspendLayout()
        Me.SAPGLAccount.SuspendLayout()
        Me.SapCoMdLogon.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapCoMd
        '
        Me.SapCoMd.Groups.Add(Me.SAPCostElement)
        Me.SapCoMd.Groups.Add(Me.SAPActivityType)
        Me.SapCoMd.Groups.Add(Me.SAPCostCenter)
        Me.SapCoMd.Groups.Add(Me.SAPInternalOrder)
        Me.SapCoMd.Groups.Add(Me.SAPProfitCenter)
        Me.SapCoMd.Groups.Add(Me.SAPGLAccount)
        Me.SapCoMd.Groups.Add(Me.SapCoMdLogon)
        Me.SapCoMd.Label = "SAP CO Md"
        Me.SapCoMd.Name = "SapCoMd"
        '
        'SAPCostElement
        '
        Me.SAPCostElement.Items.Add(Me.ButtonCECreate)
        Me.SAPCostElement.Items.Add(Me.ButtonCEChange)
        Me.SAPCostElement.Label = "CO Cost Element"
        Me.SAPCostElement.Name = "SAPCostElement"
        '
        'SAPActivityType
        '
        Me.SAPActivityType.Label = "CO Activity Type"
        Me.SAPActivityType.Name = "SAPActivityType"
        '
        'SAPCostCenter
        '
        Me.SAPCostCenter.Items.Add(Me.ButtonCCCreate)
        Me.SAPCostCenter.Items.Add(Me.ButtonCCChange)
        Me.SAPCostCenter.Label = "CO Cost Center"
        Me.SAPCostCenter.Name = "SAPCostCenter"
        '
        'ButtonCCCreate
        '
        Me.ButtonCCCreate.Image = CType(resources.GetObject("ButtonCCCreate.Image"), System.Drawing.Image)
        Me.ButtonCCCreate.Label = "Create CC"
        Me.ButtonCCCreate.Name = "ButtonCCCreate"
        Me.ButtonCCCreate.ScreenTip = "Create Cost Centers"
        Me.ButtonCCCreate.ShowImage = True
        '
        'ButtonCCChange
        '
        Me.ButtonCCChange.Image = CType(resources.GetObject("ButtonCCChange.Image"), System.Drawing.Image)
        Me.ButtonCCChange.Label = "Change CC"
        Me.ButtonCCChange.Name = "ButtonCCChange"
        Me.ButtonCCChange.ScreenTip = "Change Cost Centers"
        Me.ButtonCCChange.ShowImage = True
        '
        'SAPInternalOrder
        '
        Me.SAPInternalOrder.Label = "CO Internal Order"
        Me.SAPInternalOrder.Name = "SAPInternalOrder"
        '
        'SAPProfitCenter
        '
        Me.SAPProfitCenter.Label = "CO Profit Center"
        Me.SAPProfitCenter.Name = "SAPProfitCenter"
        '
        'SAPGLAccount
        '
        Me.SAPGLAccount.Items.Add(Me.ButtonGLAccCreate)
        Me.SAPGLAccount.Label = "GL Account"
        Me.SAPGLAccount.Name = "SAPGLAccount"
        '
        'ButtonGLAccCreate
        '
        Me.ButtonGLAccCreate.Image = CType(resources.GetObject("ButtonGLAccCreate.Image"), System.Drawing.Image)
        Me.ButtonGLAccCreate.Label = "Create Account"
        Me.ButtonGLAccCreate.Name = "ButtonGLAccCreate"
        Me.ButtonGLAccCreate.ScreenTip = "Create GL Account"
        Me.ButtonGLAccCreate.ShowImage = True
        '
        'SapCoMdLogon
        '
        Me.SapCoMdLogon.Items.Add(Me.ButtonLogon)
        Me.SapCoMdLogon.Items.Add(Me.ButtonLogoff)
        Me.SapCoMdLogon.Label = "Logon"
        Me.SapCoMdLogon.Name = "SapCoMdLogon"
        '
        'ButtonLogon
        '
        Me.ButtonLogon.Image = CType(resources.GetObject("ButtonLogon.Image"), System.Drawing.Image)
        Me.ButtonLogon.Label = "SAP Logon"
        Me.ButtonLogon.Name = "ButtonLogon"
        Me.ButtonLogon.ShowImage = True
        '
        'ButtonLogoff
        '
        Me.ButtonLogoff.Image = CType(resources.GetObject("ButtonLogoff.Image"), System.Drawing.Image)
        Me.ButtonLogoff.Label = "SAP Logoff"
        Me.ButtonLogoff.Name = "ButtonLogoff"
        Me.ButtonLogoff.ShowImage = True
        '
        'ButtonCECreate
        '
        Me.ButtonCECreate.Image = CType(resources.GetObject("ButtonCECreate.Image"), System.Drawing.Image)
        Me.ButtonCECreate.Label = "Create CE"
        Me.ButtonCECreate.Name = "ButtonCECreate"
        Me.ButtonCECreate.ScreenTip = "Create Cost Centers"
        Me.ButtonCECreate.ShowImage = True
        '
        'ButtonCEChange
        '
        Me.ButtonCEChange.Image = CType(resources.GetObject("ButtonCEChange.Image"), System.Drawing.Image)
        Me.ButtonCEChange.Label = "Change CE"
        Me.ButtonCEChange.Name = "ButtonCEChange"
        Me.ButtonCEChange.ScreenTip = "Change Cost Element"
        Me.ButtonCEChange.ShowImage = True
        '
        'SapCoMdRibbon
        '
        Me.Name = "SapCoMdRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapCoMd)
        Me.SapCoMd.ResumeLayout(False)
        Me.SapCoMd.PerformLayout()
        Me.SAPCostElement.ResumeLayout(False)
        Me.SAPCostElement.PerformLayout()
        Me.SAPCostCenter.ResumeLayout(False)
        Me.SAPCostCenter.PerformLayout()
        Me.SAPGLAccount.ResumeLayout(False)
        Me.SAPGLAccount.PerformLayout()
        Me.SapCoMdLogon.ResumeLayout(False)
        Me.SapCoMdLogon.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapCoMd As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SapCoMdLogon As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPCostCenter As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonCCCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonCCChange As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPActivityType As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents SAPInternalOrder As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents SAPCostElement As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents SAPProfitCenter As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents SAPGLAccount As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonGLAccCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonCECreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonCEChange As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property SapCoMdRibbon() As SapCoMdRibbon
        Get
            Return Me.GetRibbon(Of SapCoMdRibbon)()
        End Get
    End Property
End Class
