Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.IO

Public Enum BackupStyle As Integer
    None = 0        '無
    Overwrite = 1   '上書きorigi_*.jpg
    NewFile = 2     '新規resize_*.jpg
End Enum

Public Class ClassJpegResize
    Private mSetSize As Rectangle
    Private mBackup As BackupStyle
    Private mImageQuality As Long

    Public mErrorText As String = ""
    Private mResizeName As String = ""
    Private mBackupName As String = ""
    Private mImgSize As Rectangle
    Private mFileInfo As FileInfo
    Private mOriginalImage As Bitmap

    Public Sub New(size As Rectangle, quality As Long, backup As BackupStyle)
        mSetSize = size
        mBackup = backup
        mImageQuality = quality
    End Sub

    Private Sub Reset()
        mErrorText = ""
        mResizeName = ""
        mBackupName = ""
        mFileInfo = Nothing
        mOriginalImage = Nothing
    End Sub

    Public Sub Resize(FileName As String)
        Reset()

        If Not chkSource(FileName) Then
            Return
        End If

        If Not chkSize() Then
            Return
        End If

        makeImage()
    End Sub

    Private Function chkSource(FileName As String) As Boolean
        mFileInfo = New FileInfo(FileName)
        mFileInfo.Refresh()
        If Not mFileInfo.Exists Then
            mErrorText = "ファイルが存在しません"
            Return False
        End If

        Select Case mBackup
            Case BackupStyle.None
                mResizeName = mFileInfo.FullName
                mBackupName = mFileInfo.FullName
            Case BackupStyle.Overwrite
                mResizeName = mFileInfo.FullName
                mBackupName = mFileInfo.DirectoryName & "\origi_" & mFileInfo.Name
                mFileInfo.MoveTo(mBackupName)
            Case BackupStyle.NewFile
                mResizeName = mFileInfo.DirectoryName & "\resize_" & mFileInfo.Name
                mBackupName = mFileInfo.FullName
        End Select

        Return True
    End Function

    Private Function chkSize() As Boolean
        Dim Ret As Boolean
        mOriginalImage = New Bitmap(mBackupName)

        If mOriginalImage.Width > mOriginalImage.Height Then
            Ret = setHorizontal()
        Else
            Ret = setVertical()
        End If

        If Not Ret Then
            mOriginalImage.Dispose()
            If mBackup = BackupStyle.Overwrite Then
                mFileInfo.MoveTo(mResizeName)
            End If
            mResizeName = ""
        End If

        Return Ret
    End Function

    Private Function setHorizontal() As Boolean
        '横画像
        If mOriginalImage.Width <= mSetSize.Width And mOriginalImage.Height <= mSetSize.Height Then
            mErrorText = "縮小サイズより小さいです"
            Return False
        End If

        If mOriginalImage.Width * 0.75 > mOriginalImage.Height Then
            mImgSize.Width = mSetSize.Width
            mImgSize.Height = CInt(mOriginalImage.Height * (mSetSize.Width / mOriginalImage.Width))
        ElseIf mOriginalImage.Width * 0.75 < mOriginalImage.Height Then
            mImgSize.Width = CInt(mOriginalImage.Width * (mSetSize.Height / mOriginalImage.Height))
            mImgSize.Height = mSetSize.Height
        Else
            mImgSize.Width = mSetSize.Width
            mImgSize.Height = mSetSize.Height
        End If

        Return True
    End Function

    Private Function setVertical() As Boolean
        '縦画像
        If mOriginalImage.Height <= mSetSize.Width And mOriginalImage.Width <= mSetSize.Height Then
            mErrorText = "縮小サイズより小さいです"
            Return False
        End If

        If mOriginalImage.Height * 0.75 > mOriginalImage.Width Then
            mImgSize.Width = CInt(mOriginalImage.Width * (mSetSize.Width / mOriginalImage.Height))
            mImgSize.Height = mSetSize.Width
        ElseIf mOriginalImage.Height * 0.75 < mOriginalImage.Height Then
            mImgSize.Width = mSetSize.Height
            mImgSize.Height = CInt(mOriginalImage.Height * (mSetSize.Height / mOriginalImage.Width))
        Else
            mImgSize.Width = mSetSize.Height
            mImgSize.Height = mSetSize.Width
        End If

        Return True
    End Function

    Private Sub makeImage()
        Dim NewImage As Bitmap = New Bitmap(mImgSize.Width, mImgSize.Height)
        Dim NewGraphics As Graphics = Graphics.FromImage(NewImage)

        NewGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic
        NewGraphics.CompositingQuality = CompositingQuality.HighQuality
        NewGraphics.DrawImage(mOriginalImage, mImgSize)
        mOriginalImage.Dispose()

        Dim eps As New EncoderParameters(1)
        Dim ep As New EncoderParameter(Encoder.Quality, mImageQuality)
        eps.Param(0) = ep
        Dim ici As ImageCodecInfo = GetEncoderInfo("image/jpeg")

        NewImage.Save(mResizeName, ici, eps)
    End Sub

    Private Shared Function GetEncoderInfo(mineType As String) As ImageCodecInfo
        For Each enc As ImageCodecInfo In ImageCodecInfo.GetImageEncoders()
            If enc.MimeType = mineType Then
                Return enc
            End If
        Next
        Return Nothing
    End Function

    Public ReadOnly Property ReductionSize As Long
        Get
            Dim fi As FileInfo = New FileInfo(mResizeName)
            fi.Refresh()
            Return fi.Length
        End Get
    End Property

    Public ReadOnly Property GetResizeFile As String
        Get
            If String.IsNullOrEmpty(mResizeName) Then
                Return mBackupName
            End If

            Return mResizeName
        End Get
    End Property

    Public ReadOnly Property GetBackupFile As String
        Get
            Return mBackupName
        End Get
    End Property
End Class
