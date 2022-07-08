Partial Class RibbonPhotoExcel
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms クラス作成デザイナーのサポートに必要です
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'この呼び出しは、コンポーネント デザイナーで必要です。
        InitializeComponent()

    End Sub

    'Component は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
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

    'コンポーネント デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャはコンポーネント デザイナーで必要です。
    'コンポーネント デザイナーを使って変更できます。
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.BtnAddNew = Me.Factory.CreateRibbonButton
        Me.BtnAddPage = Me.Factory.CreateRibbonButton
        Me.BtnModPage = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.BtnAddPhoto = Me.Factory.CreateRibbonButton
        Me.BtnAddAllPhoto = Me.Factory.CreateRibbonButton
        Me.BtnAddBlank = Me.Factory.CreateRibbonButton
        Me.BtnResize = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.BtnPhotoUp = Me.Factory.CreateRibbonButton
        Me.BtnPhotoDown = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.ChkReSize = Me.Factory.CreateRibbonCheckBox
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.LbVersion = Me.Factory.CreateRibbonLabel
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Label = "フォトエクセル"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.BtnAddNew)
        Me.Group1.Items.Add(Me.BtnAddPage)
        Me.Group1.Items.Add(Me.BtnModPage)
        Me.Group1.Label = "ページ"
        Me.Group1.Name = "Group1"
        '
        'BtnAddNew
        '
        Me.BtnAddNew.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnAddNew.Label = "新規作成"
        Me.BtnAddNew.Name = "BtnAddNew"
        Me.BtnAddNew.OfficeImageId = "CreateReportBlankReport"
        Me.BtnAddNew.ShowImage = True
        '
        'BtnAddPage
        '
        Me.BtnAddPage.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnAddPage.Label = "ページ追加"
        Me.BtnAddPage.Name = "BtnAddPage"
        Me.BtnAddPage.OfficeImageId = "SourceControlAddObjects"
        Me.BtnAddPage.ShowImage = True
        '
        'BtnModPage
        '
        Me.BtnModPage.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnModPage.Label = "ページ修正"
        Me.BtnModPage.Name = "BtnModPage"
        Me.BtnModPage.OfficeImageId = "ClickToRunUpdateOptions"
        Me.BtnModPage.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btnAddPhoto)
        Me.Group2.Items.Add(Me.btnAddAllPhoto)
        Me.Group2.Items.Add(Me.btnAddBlank)
        Me.Group2.Items.Add(Me.btnResize)
        Me.Group2.Label = "写真"
        Me.Group2.Name = "Group2"
        '
        'BtnAddPhoto
        '
        Me.BtnAddPhoto.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnAddPhoto.Label = "写真追加"
        Me.BtnAddPhoto.Name = "BtnAddPhoto"
        Me.BtnAddPhoto.OfficeImageId = "PictureReflectionGalleryItem"
        Me.BtnAddPhoto.ShowImage = True
        '
        'BtnAddAllPhoto
        '
        Me.BtnAddAllPhoto.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnAddAllPhoto.Label = "フォルダ読込"
        Me.BtnAddAllPhoto.Name = "BtnAddAllPhoto"
        Me.BtnAddAllPhoto.OfficeImageId = "FileOpen"
        Me.BtnAddAllPhoto.ShowImage = True
        '
        'BtnAddBlank
        '
        Me.BtnAddBlank.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnAddBlank.Label = "空白追加"
        Me.BtnAddBlank.Name = "BtnAddBlank"
        Me.BtnAddBlank.OfficeImageId = "BevelShapeGallery"
        Me.BtnAddBlank.ShowImage = True
        '
        'BtnResize
        '
        Me.BtnResize.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnResize.Label = "リサイズ"
        Me.BtnResize.Name = "BtnResize"
        Me.BtnResize.OfficeImageId = "ControlLogo"
        Me.BtnResize.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.btnPhotoUp)
        Me.Group3.Items.Add(Me.btnPhotoDown)
        Me.Group3.Label = "移動"
        Me.Group3.Name = "Group3"
        '
        'BtnPhotoUp
        '
        Me.BtnPhotoUp.Label = "一段上げる"
        Me.BtnPhotoUp.Name = "BtnPhotoUp"
        Me.BtnPhotoUp.OfficeImageId = "OutlineMoveUp"
        Me.BtnPhotoUp.ShowImage = True
        '
        'BtnPhotoDown
        '
        Me.BtnPhotoDown.Label = "一段下げる"
        Me.BtnPhotoDown.Name = "BtnPhotoDown"
        Me.BtnPhotoDown.OfficeImageId = "OutlineMoveDown"
        Me.BtnPhotoDown.ShowImage = True
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.ChkReSize)
        Me.Group4.Label = "縮小"
        Me.Group4.Name = "Group4"
        '
        'ChkReSize
        '
        Me.ChkReSize.Label = "縮小して取込"
        Me.ChkReSize.Name = "ChkReSize"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.LbVersion)
        Me.Group5.Label = "バージョン"
        Me.Group5.Name = "Group5"
        '
        'LbVersion
        '
        Me.LbVersion.Label = "Label1"
        Me.LbVersion.Name = "LbVersion"
        '
        'RibbonPhotoExcel
        '
        Me.Name = "RibbonPhotoExcel"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BtnAddNew As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnAddPage As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnModPage As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BtnAddPhoto As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnAddAllPhoto As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnAddBlank As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnResize As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BtnPhotoUp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnPhotoDown As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ChkReSize As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents LbVersion As Microsoft.Office.Tools.Ribbon.RibbonLabel
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As RibbonPhotoExcel
        Get
            Return Me.GetRibbon(Of RibbonPhotoExcel)()
        End Get
    End Property
End Class
