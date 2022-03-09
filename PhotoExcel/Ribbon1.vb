Imports Microsoft.Office.Tools.Ribbon

Public Class RibbonPhotoExcel
    Private Sub BtnAddNew_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAddNew.Click
        PageNew()
    End Sub

    Private Sub btnAddPage_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAddPage.Click
        PageAdd()
    End Sub

    Private Sub btnAddPhoto_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAddPhoto.Click
        PutPhotos(GetFiles())
    End Sub

    Private Sub ChkReSize_Click(sender As Object, e As RibbonControlEventArgs) Handles ChkReSize.Click
        ReSize = ChkReSize.Checked
    End Sub

    Private Sub btnAddAllPhoto_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAddAllPhoto.Click
        PutPhotos(GetFolder())
    End Sub

    Private Sub btnAddBlank_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAddBlank.Click
        BlankAdd()
    End Sub

    Private Sub btnResize_Click(sender As Object, e As RibbonControlEventArgs) Handles btnResize.Click
        PhotoResize()
    End Sub

    Private Sub btnModPage_Click(sender As Object, e As RibbonControlEventArgs) Handles btnModPage.Click
        PageModify()
    End Sub

    Private Sub btnPhotoUp_Click(sender As Object, e As RibbonControlEventArgs) Handles btnPhotoUp.Click
        PhotoUp()
    End Sub

    Private Sub btnPhotoDown_Click(sender As Object, e As RibbonControlEventArgs) Handles btnPhotoDown.Click
        PhotoDown()
    End Sub

    Private Sub RibbonPhotoExcel_Load(sender As Object, e As RibbonUIEventArgs) Handles Me.Load
        With My.Application.Info.Version
            LbVersion.Label = .Major & "." & .Minor & "." & .Build & "." & .Revision
        End With
    End Sub
End Class
