﻿Module Photo
    Public Const PhotoWidth As Double = 338
    Public Const PhotoHeight As Double = 253.5

    Public ImageResize As New Size With {.Width = 1024, .Height = 768}
    Public Const ImageQuality As Long = 90

    Public ReSize As Boolean = False
    Public AddDate As Boolean = False
    Public Application As Excel.Application = Globals.ThisAddIn.Application
    Public ActiveSheet As Excel.Worksheet = Application.ActiveSheet
    Private OpenDirectory As String = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)

    Public Function GetFiles() As ArrayList
        Dim ofd As New OpenFileDialog With {
            .InitialDirectory = OpenDirectory,
            .Filter = "JPEG画像(*.jpg; *.jpeg)|*.jpg; *.jpeg",
            .Multiselect = True
        }

        Dim al As New ArrayList
        If ofd.ShowDialog() <> DialogResult.Cancel Then
        	OpenDirectory = Path.GetDirectoryName(ofd.FileNames(0))
            al.AddRange(ofd.FileNames)
        End If

        Return al
    End Function

    Public Function GetFolder() As ArrayList
        Dim fbd As New FolderBrowserDialog With {
            .Description = "フォルダを指定してください。",
            .RootFolder = Environment.SpecialFolder.Desktop
        }

        Dim al As New ArrayList
        If fbd.ShowDialog() = DialogResult.OK Then
            Dim mFolderName As String = fbd.SelectedPath
            If Not String.IsNullOrEmpty(mFolderName) Then
                Dim files As String() = IO.Directory.GetFiles(mFolderName, "*.jpg", IO.SearchOption.AllDirectories)
                If files.Count > 0 Then
                    al.AddRange(files)
                End If

                Dim files2 As String() = IO.Directory.GetFiles(mFolderName, "*.jpeg", IO.SearchOption.AllDirectories)
                If files2.Count > 0 Then
                    al.AddRange(files2)
                End If
            End If
        End If

        Return al
    End Function

    Public Sub PutPhotos(mFileList As ArrayList)
        If ActiveSheet.Columns("A").ColumnWidth <> COL_WIDTH_A Then
            PageNew()
        End If

        If ActiveSheet.Columns("A").ColumnWidth = COL_WIDTH_A Then
            If mFileList.Count > 0 Then
                Application.ScreenUpdating = False
                Dim Row As Integer
                Dim C As Integer

                GetSelectCell(Row, C)
                AllPageNum = Int((mFileList.Count + C) / 3 + 0.9)

                For P = 1 To AllPageNum - 1
                    SetPageStyle(P)
                Next P

                ActiveSheet.Cells(Row, 1).Select
                PutPhoto(mFileList)
                Application.ScreenUpdating = True
            End If
        End If
    End Sub

    Private Sub GetSelectCell(ByRef Row As Integer, ByRef Cnt As Integer)
        Dim ActiveCell As Excel.Range = Application.ActiveCell
        If ActiveCell.Column = 1 Then
            Dim P As Integer = Int(ActiveCell.Row / PageRows)

            Cnt = Math.Min(Math.Ceiling((ActiveCell.Row - (P * PageRows)) / 14), 3) - 1
            Row = P * PageRows + Cnt * 14 + 2
        Else
            Row = 2
            Cnt = 0
        End If
    End Sub

    Private Function GetSelectRow() As Integer
        Dim R1 As Integer = 2

        For Page As Integer = 0 To PageNum() - 1
            Dim sRow As Integer = Page * PageRows + 2
            For R As Integer = 0 To 2
                R1 = R * PhotoRows + sRow + (R * 1)
                If Application.Selection.ShapeRange.Top = ActiveSheet.Cells(R1, 1).Top Then
                    Return R1
                End If
            Next R
        Next Page

        Return R1
    End Function

    Private Sub PutPhoto(fl As ArrayList)
        Dim Row As Integer
        Dim C As Integer
        Dim FileName As String
        Dim CopyFile As String
        Dim OriginalFile As FileInfo
        GetSelectCell(Row, C)
        Dim Img As New JpegImage(ImageResize, ImageQuality)
        For Each FileName In fl
            C += 1

            OriginalFile = New FileInfo(FileName)
            CopyFile = OriginalFile.DirectoryName & "\copy_" & OriginalFile.Name
            OriginalFile.CopyTo(CopyFile, True)

            Img.ChangeRotate(CopyFile)
            If ReSize Then
                Img.ReSize(CopyFile)
            End If

            With ActiveSheet.Shapes.AddPicture(CopyFile, False, True, 0, 0, 0, 0)
                .LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                .ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue)
                .ScaleWidth(1, Microsoft.Office.Core.MsoTriState.msoTrue)

                .Width = PhotoWidth
                If .Height > PhotoHeight Then
                    .LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                    .Height = PhotoHeight
                End If
                .Top = Globals.ThisAddIn.Application.ActiveSheet.Cells(Row, 1).Top
                .Left = 0

                If AddDate Then
                    Dim tname As String = PutDate(.Top, GetExifDate(FileName))
                    ActiveSheet.Shapes.Range({tname, .Name}).Group()
                End If
            End With

            Row += 14
            If C = 3 Then
                C = 0
                Row += 1
            End If

            Kill(CopyFile)
        Next
    End Sub

    Private Function GetExifDate(fn As String) As Date
        Dim bmp As New Bitmap(fn)
        Dim item As PropertyItem
        Dim dt As Date = File.GetCreationTime(fn)
        For Each item In bmp.PropertyItems
            If item.Id = &H9003 And item.Type = 2 Then
                Dim val As String = Encoding.ASCII.GetString(item.Value).Trim(New Char() {ControlChars.NullChar})
                dt = DateTime.ParseExact(val, "yyyy:MM:dd HH:mm:ss", Nothing)
                Exit For
            End If
        Next item
        bmp.Dispose()

        Return dt
    End Function

    Public Sub PhotoResize()
        Try
            With Application.Selection
                If .ShapeRange.Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                    .ShapeRange.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse
                    .ShapeRange.Width = PhotoWidth
                    .ShapeRange.Height = PhotoHeight
                End If
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Function GetMaxRow() As Integer
        Dim sRow As Integer = (PageNum() - 1) * PageRows + 2
        Return 2 * PhotoRows + sRow + 2
    End Function

    Private Function NextRow(Row As Integer) As Integer
        Dim R As Integer = Row + PhotoRows + 1
        Dim P As Integer = Int(R / PageRows)

        Dim Cnt = Math.Min(Math.Ceiling((R - (P * PageRows)) / 14), 3) - 1
        Return P * PageRows + Cnt * 14 + 2
    End Function

    Private Function PreRow(Row As Integer) As Integer
        Dim R As Integer = Row - PhotoRows - 1
        Dim P As Integer = Int(R / PageRows)

        Dim Cnt = Math.Min(Math.Ceiling((R - (P * PageRows)) / 14), 3) - 1
        Return P * PageRows + Cnt * 14 + 2
    End Function

    Public Sub PhotoUp()
        Try
            If Application.Selection.ShapeRange.Type <> Microsoft.Office.Core.MsoShapeType.msoPicture Then
                Return
            End If

            Application.ScreenUpdating = False
            Application.Selection.ShapeRange.Top = ActiveSheet.Cells(PreRow(GetSelectRow()), 1).Top
        Catch ex As Exception
            MessageBox.Show("画像を選択して下さい。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Application.ScreenUpdating = True
        End Try
    End Sub

    Public Sub PhotoDown()
        Try
            If Application.Selection.ShapeRange.Type <> Microsoft.Office.Core.MsoShapeType.msoPicture Then
                Return
            End If

            Application.ScreenUpdating = False
            Dim sRow As Integer = GetSelectRow()
            If sRow >= GetMaxRow() Then
                PageAdd()
            End If
            Application.Selection.ShapeRange.Top = ActiveSheet.Cells(NextRow(sRow), 1).Top
        Catch ex As Exception
            MessageBox.Show("画像を選択して下さい。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Application.ScreenUpdating = True
        End Try
    End Sub

End Module
