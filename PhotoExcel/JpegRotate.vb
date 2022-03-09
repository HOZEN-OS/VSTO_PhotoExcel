Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO

Public Class JpegRotate
    Public Shared Function ChangeRotate(FileName As String) As String
        Dim mFileInfo As New FileInfo(FileName)
        mFileInfo.Refresh()

        Using origin As New Bitmap(FileName)
            Dim pi As PropertyItem
            Dim rotation As RotateFlipType = RotateFlipType.RotateNoneFlipNone
            For Each item As PropertyItem In origin.PropertyItems
                If item.Id = &H112 Then
                    Select Case item.Value(0)
                        Case 3
                            rotation = RotateFlipType.Rotate180FlipNone
                        Case 6
                            rotation = RotateFlipType.Rotate90FlipNone
                        Case 8
                            rotation = RotateFlipType.Rotate270FlipNone
                    End Select
                    pi = item
                    Exit For
                End If
            Next

            If rotation = RotateFlipType.RotateNoneFlipNone Then
                Return FileName
            End If

            Using rotated As Bitmap = DirectCast(origin.Clone(), Bitmap)
                Dim mBackupName As String = mFileInfo.DirectoryName & "\rotate_" & mFileInfo.Name

                rotated.RotateFlip(rotation)

                pi.Value(0) = &H1
                pi.Len = pi.Value.Length
                rotated.SetPropertyItem(pi)

                Dim eps As New EncoderParameters(1)
                Dim ep As New EncoderParameter(Encoder.Quality, 90)
                eps.Param(0) = ep
                Dim ici As ImageCodecInfo = GetEncoderInfo("image/jpeg")

                rotated.Save(mBackupName, ici, eps)

                Return mBackupName
            End Using
        End Using
    End Function

    Private Shared Function GetEncoderInfo(mineType As String) As ImageCodecInfo
        For Each enc As ImageCodecInfo In ImageCodecInfo.GetImageEncoders()
            If enc.MimeType = mineType Then
                Return enc
            End If
        Next
        Return Nothing
    End Function
End Class
