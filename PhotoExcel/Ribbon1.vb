Imports Microsoft.Office.Tools.Ribbon

Public Class RibbonPhotoExcel
    Private Sub BtnAddNew_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnAddNew.Click
        PageNew()
    End Sub

    Private Sub BtnAddPage_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnAddPage.Click
        PageAdd()
    End Sub

    Private Sub BtnAddPhoto_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnAddPhoto.Click
        PutPhotos(GetFiles())
    End Sub

    Private Sub ChkReSize_Click(sender As Object, e As RibbonControlEventArgs) Handles ChkReSize.Click
        ReSize = ChkReSize.Checked
    End Sub

    Private Sub BtnAddAllPhoto_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnAddAllPhoto.Click
        PutPhotos(GetFolder())
    End Sub

    Private Sub BtnAddBlank_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnAddBlank.Click
        BlankAdd()
    End Sub

    Private Sub BtnResize_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnResize.Click
        PhotoResize()
    End Sub

    Private Sub BtnModPage_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnModPage.Click
        PageModify()
    End Sub

    Private Sub BtnPhotoUp_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnPhotoUp.Click
        PhotoUp()
    End Sub

    Private Sub BtnPhotoDown_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnPhotoDown.Click
        PhotoDown()
    End Sub

    Private Sub RibbonPhotoExcel_Load(sender As Object, e As RibbonUIEventArgs) Handles Me.Load
        With My.Application.Info.Version
            LbVersion.Label = .Major & "." & .Minor & "." & .Build & "." & .Revision
        End With
    End Sub
End Class
