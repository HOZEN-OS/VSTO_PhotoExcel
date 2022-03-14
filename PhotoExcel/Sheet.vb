Imports System.Windows.Forms

Module Sheet
    Public Const COL_WIDTH_A As Double = 55.63
    Public Const COL_WIDTH_B As Double = 2.78
    Public Const COL_WIDTH_C As Double = 20.0#
    Public Const ROW_HEIGHT As Double = 19.5
    Public Const PageRows As Integer = 43
    Public Const PhotoRows As Integer = 13

    Public AllPageNum As Integer

    Private Pages As Integer

    Public Sub PageNew()
        If MessageBox.Show("シートの中身が消えますが、よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        Application.ScreenUpdating = False

        Clear()
        SetFontSize()
        SetCellWidth()
        SetPageStyle()
        SetPrintStyle()

        Application.ScreenUpdating = True
    End Sub

    Private Sub SetCellWidth()
        With ActiveSheet
            .Cells.RowHeight = ROW_HEIGHT
            .Columns("A").ColumnWidth = COL_WIDTH_A
            .Columns("B").ColumnWidth = COL_WIDTH_B
            .Columns("C").ColumnWidth = COL_WIDTH_C
        End With
    End Sub

    Public Sub PageAdd()
        If ActiveSheet.Columns("A").ColumnWidth = COL_WIDTH_A Then
            Application.ScreenUpdating = False
            SetPageStyle(PageNum())
            Application.ScreenUpdating = True
        Else
            PageNew()
        End If
    End Sub

    Public Sub PageModify()
        If MessageBox.Show("シートのスタイルと写真サイズを再設定します、よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        Application.ScreenUpdating = False

        SetFontSize()
        SetCellWidth()

        Pages = PageNum()
        With ActiveSheet
            For I As Integer = .Shapes.Count - 1 To 0 Step -1
                If .Shapes(I).Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                    If Math.Abs(.Shapes(I).Left) < 10 Then
                        .Shapes(I).Locked = Microsoft.Office.Core.MsoTriState.msoFalse
                        .Shapes(I).Placement = Excel.XlPlacement.xlMove
                        .Shapes(I).LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse
                        .Shapes(I).Left = 0
                        .Shapes(I).Top = GetTop(.Shapes(I).Top)
                        .Shapes(I).Height = PhotoHeight
                        .Shapes(I).Width = PhotoWidth
                        .Shapes(I).Locked = Microsoft.Office.Core.MsoTriState.msoTrue
                    End If
                End If
            Next
        End With

        SetPrintStyle()

        For I = 0 To Pages - 1
            SetPageStyle(I)
        Next I

        Application.ScreenUpdating = True
    End Sub

    Private Function GetTop(top As Double) As Double
        Dim P As Integer
        Dim R As Integer
        Dim R1 As Integer

        For P = 0 To Pages - 1
            For R = 0 To 2
                R1 = R * PhotoRows + (P * PageRows) + 2 + (R * 1)

                If Math.Abs(top - ActiveSheet.Cells(R1, 1).top) < 50 Then
                    Return ActiveSheet.Cells(R1, 1).top
                End If
            Next R
        Next P

        Return 0
    End Function

    Public Function PageNum() As Integer
        Dim ActiveWindow As Excel.Window = Application.ActiveWindow
        ActiveWindow.View = Excel.XlWindowView.xlPageBreakPreview
        Dim PNum As Integer = Application.ExecuteExcel4Macro("GET.DOCUMENT(50)")
        ActiveWindow.View = Excel.XlWindowView.xlNormalView
        Return PNum
    End Function

    Private Sub Clear()
        ActiveSheet.Cells.Delete()

        With ActiveSheet
            For I As Integer = .Shapes.Count To 1 Step -1
                .Shapes(I).Delete
            Next I
        End With
    End Sub

    Private Sub SetFontSize()
        With Application.ActiveWorkbook.Styles("Normal").Font
            If .Name <> "ＭＳ ゴシック" Then
                .Name = "ＭＳ ゴシック"
            End If
            If .Size <> 11 Then
                .Size = 11
            End If
        End With

        With ActiveSheet.Cells.Font
            If .Name.ToString <> "ＭＳ ゴシック" OrElse String.IsNullOrEmpty(.Name) Then
                .Name = "ＭＳ ゴシック"
            End If
            If .Size <> 11 Then
                .Size = 11
            End If
        End With
    End Sub

    Public Sub SetPageStyle(Optional Page As Integer = 0)
        Dim sRow As Integer
        Dim R As Integer
        Dim R1 As Integer
        Dim R2 As Integer
        Dim mRow As Integer

        sRow = Page * PageRows + 2

        For R = 0 To 2
            R1 = R * PhotoRows + sRow + (R * 1)
            R2 = R1 + PhotoRows - 1

            CellMerge(ActiveSheet.Range(ActiveSheet.Cells(R1, 1), ActiveSheet.Cells(R2, 1)))
            SetBorder(R1, R2)
        Next R

        mRow = (sRow - 1 + ((PhotoRows + 1) * 3))
        ActiveSheet.PageSetup.PrintArea = "$A$1:$C$" & mRow

        If Page > 0 Then
            Dim ActiveWindow As Excel.Window = Application.ActiveWindow
            ActiveWindow.View = Excel.XlWindowView.xlPageBreakPreview
            ActiveSheet.HPageBreaks(Page).Location = ActiveSheet.Range("C" & mRow - PageRows + 1)
            ActiveWindow.View = Excel.XlWindowView.xlNormalView
        End If
    End Sub

    Public Sub BlankAdd()
        Try
            Dim ActiveCell As Excel.Range = Application.ActiveCell
            If ActiveCell.Column = 1 Then
                Dim mRange As Excel.Range = ActiveSheet.Range(ActiveSheet.Cells(ActiveCell.Row, 1), ActiveSheet.Cells(ActiveCell.Row + PhotoRows - 1, 1))
                With mRange
                    If ActiveCell.Value = "" Then
                        ActiveCell.Value = "空白"
                        .Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                        .Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                        .Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.Constants.xlNone
                        .Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.Constants.xlNone

                        BorderDraw(mRange, Excel.XlBordersIndex.xlEdgeLeft)
                        BorderDraw(mRange, Excel.XlBordersIndex.xlEdgeTop)
                        BorderDraw(mRange, Excel.XlBordersIndex.xlEdgeBottom)
                        BorderDraw(mRange.Cells, Excel.XlBordersIndex.xlEdgeRight)

                        With mRange.Font
                            .Color = RGB(166, 166, 166)
                        End With
                    Else
                        .ClearContents()
                        .Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                        .Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                        .Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.Constants.xlNone
                        .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.Constants.xlNone
                        .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.Constants.xlNone
                        .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.Constants.xlNone
                        .Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.Constants.xlNone
                        .Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.Constants.xlNone
                    End If
                End With
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SetBorder(R1 As Integer, R2 As Integer)
        Dim mRange As Excel.Range = ActiveSheet.Range(ActiveSheet.Cells(R1, 3), ActiveSheet.Cells(R2, 3))
        With mRange
            .Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
            .Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
            .Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.Constants.xlNone
        End With

        BorderDraw(mRange, Excel.XlBordersIndex.xlEdgeLeft)
        BorderDraw(mRange, Excel.XlBordersIndex.xlEdgeTop)
        BorderDraw(mRange, Excel.XlBordersIndex.xlEdgeBottom)
        BorderDraw(mRange, Excel.XlBordersIndex.xlEdgeRight)
        BorderDraw(mRange, Excel.XlBordersIndex.xlInsideHorizontal)
    End Sub

    Private Sub BorderDraw(mRange As Excel.Range, BSide As Excel.XlBordersIndex)
        With mRange.Borders(BSide)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.349986266670736
            .Weight = Excel.XlBorderWeight.xlThin
        End With
    End Sub

    Private Sub CellMerge(mRange As Excel.Range)
        With mRange
            .HorizontalAlignment = Excel.Constants.xlCenter
            .VerticalAlignment = Excel.Constants.xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
        End With
        mRange.Merge()
    End Sub

    Private Sub SetPrintStyle()
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .PrintTitleRows = ""
            .PrintTitleColumns = ""
        End With
        Application.PrintCommunication = True

        With ActiveSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(1.18110236220472)      '3.0
            .RightMargin = Application.InchesToPoints(0.393700787401575)
            .TopMargin = Application.InchesToPoints(0.393700787401575)     '1.0
            .BottomMargin = Application.InchesToPoints(0.393700787401575)  '1.0
            .HeaderMargin = Application.InchesToPoints(0.31496062992126)
            .FooterMargin = Application.InchesToPoints(0.31496062992126)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = True
            .Orientation = Excel.XlPageOrientation.xlPortrait
            .Draft = False
            .PaperSize = Excel.XlPaperSize.xlPaperA4
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = False
            .Zoom = 100
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End With
    End Sub
End Module
