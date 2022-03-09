Imports System.Collections
Imports System.Windows.Forms

Module Photo
    Public Const PhotoWidth As Double = 338
    Public Const PhotoHeight As Double = 253.5

    Public ReSize As Boolean = False
    Public Application As Excel.Application = Globals.ThisAddIn.Application
    Public ActiveSheet As Excel.Worksheet = Application.ActiveSheet

    Private SelectedPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)

    Public Function GetFiles() As ArrayList
        Dim ofd As New OpenFileDialog With {
            .InitialDirectory = SelectedPath,
            .Filter = "JPEG画像(*.jpg; *.jpeg)|*.jpg; *.jpeg",
            .Multiselect = True
        }

        Dim al As New ArrayList
        If ofd.ShowDialog() <> DialogResult.Cancel Then
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
        Dim Size As New Drawing.Rectangle With {
            .Width = 1024,
            .Height = 768
        }

        GetSelectCell(Row, C)

        For Each fn As String In fl
            C += 1
            If ReSize Then
                Dim ImgRe As New ClassJpegResize(Size, 90, BackupStyle.NewFile)
                ImgRe.Resize(fn)
                FileName = ImgRe.GetResizeFile
            Else
                FileName = fn
            End If

            With ActiveSheet.Shapes.AddPicture(FileName, False, True, 0, 0, 0, 0)
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
            End With

            Row = Row + 14
            If C = 3 Then
                C = 0
                Row = Row + 1
            End If

            If ReSize Then
                Kill(FileName)
            End If
        Next
    End Sub

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
