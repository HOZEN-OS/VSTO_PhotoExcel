Public Class JpegImage
    Private ReadOnly SetSize As Size
    Private ReadOnly ImageQuality As Long

    Public Sub New(size As Size, quality As Long)
        SetSize = size
        ImageQuality = quality
    End Sub

    ''' <summary>
    ''' 画像の縮小
    ''' </summary>
    ''' <param name="ImageFile"></param>
    Public Sub ReSize(ImageFile As String)
        Dim ReadBitmap As New Bitmap(ImageFile)
        Dim OrgSize As New Size(ReadBitmap.Width, ReadBitmap.Height)
        Dim ImgSize As New Rectangle

        If OrgSize.Width > OrgSize.Height Then
            ImgSize.Size = GetHorizontalSize(OrgSize.Width, OrgSize.Height)
        Else
            ImgSize.Size = GetVerticalSize(OrgSize.Width, OrgSize.Height)
        End If

        If OrgSize.Width <= ImgSize.Width AndAlso OrgSize.Height <= ImgSize.Height Then
            ReadBitmap.Dispose()
        Else
            Dim SaveBitmap As New Bitmap(ImgSize.Width, ImgSize.Height)
            Dim NewGraphics As Graphics = Graphics.FromImage(SaveBitmap)
            NewGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic
            NewGraphics.CompositingQuality = CompositingQuality.HighQuality
            NewGraphics.DrawImage(ReadBitmap, ImgSize)
            ReadBitmap.Dispose()

            Dim eps As New EncoderParameters(1)
            eps.Param(0) = New EncoderParameter(Imaging.Encoder.Quality, ImageQuality)
            Dim ici As ImageCodecInfo = GetEncoderInfo("image/jpeg")
            SaveBitmap.Save(ImageFile, ici, eps)
            SaveBitmap.Dispose()
        End If
    End Sub

    Private Function GetHorizontalSize(Width As Integer, Height As Integer) As Size
        If Width * 0.75 > Height Then
            Return New Size(SetSize.Width, CInt(Height * (SetSize.Width / Width)))
        ElseIf Width * 0.75 < Height Then
            Return New Size(CInt(Width * (SetSize.Height / Height)), SetSize.Height)
        Else
            Return SetSize
        End If
    End Function

    Private Function GetVerticalSize(Width As Integer, Height As Integer) As Size
        If Height * 0.75 > Width Then
            Return New Size(CInt(Width * (SetSize.Width / Height)), SetSize.Width)
        ElseIf Height * 0.75 < Width Then
            Return New Size(SetSize.Height, CInt(Height * (SetSize.Height / Width)))
        Else
            Return New Size(SetSize.Height, SetSize.Width)
        End If
    End Function

    ''' <summary>
    ''' 画像の縦横表示を固定
    ''' </summary>
    ''' <param name="ImageFile"></param>
    Public Sub ChangeRotate(ImageFile As String)
        Dim ReadBitmap As New Bitmap(ImageFile)
        Dim SaveBitmap As Bitmap = DirectCast(ReadBitmap.Clone(), Bitmap)
        Dim Rotation As RotateFlipType = RotateFlipType.RotateNoneFlipNone
        Dim PropItem As PropertyItem = Nothing

        For Each PropItem In ReadBitmap.PropertyItems
            If PropItem.Id = &H112 Then
                Select Case PropItem.Value(0)
                    Case 3
                        Rotation = RotateFlipType.Rotate180FlipNone
                    Case 6
                        Rotation = RotateFlipType.Rotate90FlipNone
                    Case 8
                        Rotation = RotateFlipType.Rotate270FlipNone
                End Select
                Exit For
            End If
        Next
        ReadBitmap.Dispose()

        If Rotation <> RotateFlipType.RotateNoneFlipNone Then
            PropItem.Value(0) = &H1
            PropItem.Len = PropItem.Value.Length
            SaveBitmap.RotateFlip(Rotation)
            SaveBitmap.SetPropertyItem(PropItem)

            Dim eps As New EncoderParameters(1)
            eps.Param(0) = New EncoderParameter(Imaging.Encoder.Quality, 90)

            SaveBitmap.Save(ImageFile, GetEncoderInfo("image/jpeg"), eps)
        End If
        SaveBitmap.Dispose()
    End Sub

    Private Function GetEncoderInfo(mineType As String) As ImageCodecInfo
        For Each enc As ImageCodecInfo In ImageCodecInfo.GetImageEncoders()
            If enc.MimeType = mineType Then
                Return enc
            End If
        Next
        Return Nothing
    End Function
End Class
