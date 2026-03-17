Public Class Binaris1001
    Dim n As Integer = 6
    Dim cells(n, n) As Button
    Dim board(n, n), solution(n, n) As Integer
    Dim start As Boolean

    'Rule 1: No triples allowed (three consecutive 0s or 1s)
    'Rule 2: Each row and column must contain an equal number of 0s and 1s
    'Rule 3: Any two rows or columns must be unique (cannot be exactly the same)

    Private Function InvalidRows(n As Integer, board(,) As Integer, row As Integer, col As Integer, bin As Integer) As List(Of Integer)
        Dim invalidRowsIdx As New List(Of Integer)
        If bin <> -1 Then
            Dim isCurrRowInvalid As Boolean = False
            Dim temp As Integer = board(row, col)

            board(row, col) = bin

            'Rule 1
            If col - 2 >= 0 Then
                If board(row, col - 2) = board(row, col - 1) And board(row, col - 1) = board(row, col) Then
                    isCurrRowInvalid = True
                End If
            End If
            If col + 2 <= n - 1 Then
                If board(row, col + 2) = board(row, col + 1) And board(row, col + 1) = board(row, col) Then
                    isCurrRowInvalid = True
                End If
            End If
            If col - 1 >= 0 And col + 1 <= n - 1 Then
                If board(row, col - 1) = board(row, col) And board(row, col) = board(row, col + 1) Then
                    isCurrRowInvalid = True
                End If
            End If

            'Rule 2
            Dim zerosRowCnt As Integer = 0, onesRowCnt As Integer = 0
            For j = 0 To n - 1
                If board(row, j) = 0 Then
                    zerosRowCnt += 1
                ElseIf board(row, j) = 1 Then
                    onesRowCnt += 1
                End If
            Next
            If zerosRowCnt + onesRowCnt = n Then
                If zerosRowCnt <> onesRowCnt Then
                    isCurrRowInvalid = True
                End If
            Else
                If zerosRowCnt > n \ 2 Or onesRowCnt > n \ 2 Then
                    isCurrRowInvalid = True
                End If
            End If

            'Rule 3
            For i = 0 To n - 1
                If i <> row Then
                    Dim isRowUnique As Boolean = False
                    For j = 0 To n - 1
                        If board(i, j) = -1 Or board(row, j) = -1 Then
                            isRowUnique = True
                            Exit For
                        End If
                        If board(i, j) <> board(row, j) Then
                            isRowUnique = True
                            Exit For
                        End If
                    Next
                    If Not isRowUnique Then
                        isCurrRowInvalid = True
                        invalidRowsIdx.Add(i)
                    End If
                End If
            Next

            If isCurrRowInvalid Then
                invalidRowsIdx.Add(row)
            End If

            board(row, col) = temp
        End If
        Return invalidRowsIdx
    End Function

    Private Function InvalidCols(n As Integer, board(,) As Integer, row As Integer, col As Integer, bin As Integer) As List(Of Integer)
        Dim invalidColsIdx As New List(Of Integer)
        If bin <> -1 Then
            Dim isCurrColInvalid As Boolean = False
            Dim temp As Integer = board(row, col)

            board(row, col) = bin

            'Rule 1
            If row - 2 >= 0 Then
                If board(row - 2, col) = board(row - 1, col) And board(row - 1, col) = board(row, col) Then
                    isCurrColInvalid = True
                End If
            End If
            If row + 2 <= n - 1 Then
                If board(row + 2, col) = board(row + 1, col) And board(row + 1, col) = board(row, col) Then
                    isCurrColInvalid = True
                End If
            End If
            If row - 1 >= 0 And row + 1 <= n - 1 Then
                If board(row - 1, col) = board(row, col) And board(row, col) = board(row + 1, col) Then
                    isCurrColInvalid = True
                End If
            End If

            'Rule 2
            Dim zerosColCnt As Integer = 0, onesColCnt As Integer = 0
            For i = 0 To n - 1
                If board(i, col) = 0 Then
                    zerosColCnt += 1
                ElseIf board(i, col) = 1 Then
                    onesColCnt += 1
                End If
            Next
            If zerosColCnt + onesColCnt = n Then
                If zerosColCnt <> onesColCnt Then
                    isCurrColInvalid = True
                End If
            Else
                If zerosColCnt > n \ 2 Or onesColCnt > n \ 2 Then
                    isCurrColInvalid = True
                End If
            End If

            'Rule 3
            For j = 0 To n - 1
                If j <> col Then
                    Dim isColUnique As Boolean = False
                    For i = 0 To n - 1
                        If board(i, j) = -1 Or board(i, col) = -1 Then
                            isColUnique = True
                            Exit For
                        End If
                        If board(i, j) <> board(i, col) Then
                            isColUnique = True
                            Exit For
                        End If
                    Next
                    If Not isColUnique Then
                        isCurrColInvalid = True
                        invalidColsIdx.Add(j)
                    End If
                End If
            Next

            If isCurrColInvalid Then
                invalidColsIdx.Add(col)
            End If

            board(row, col) = temp
        End If
        Return invalidColsIdx
    End Function

    Dim success As Boolean
    Private Function RandomizeBoardState(steps As Integer, n As Integer, board(,) As Integer)
        If steps = n ^ 2 Then
            success = True
        Else
            Dim rn As Random = New Random()
            Dim bin As Integer = rn.Next(2)
            Dim row As Integer = steps \ n, col As Integer = steps Mod n
            success = False
            If Not success Then
                If InvalidRows(n, board, row, col, bin).Count = 0 And InvalidCols(n, board, row, col, bin).Count = 0 Then
                    board(row, col) = bin
                    If Not RandomizeBoardState(steps + 1, n, board) Then
                        board(row, col) = -1
                    End If
                End If
            End If
            If Not success Then
                If InvalidRows(n, board, row, col, 1 - bin).Count = 0 And InvalidCols(n, board, row, col, 1 - bin).Count = 0 Then
                    board(row, col) = 1 - bin
                    If Not RandomizeBoardState(steps + 1, n, board) Then
                        board(row, col) = -1
                    End If
                End If
            End If
        End If
        Return success
    End Function

    Private Sub FillBoard(n As Integer, board(,) As Integer, cells(,) As Button)
        Dim rn As New Random()
        Dim emptyCellsCnt As Integer = 0
        Dim cellChoices As New List(Of KeyValuePair(Of Integer, Integer))
        For i = 0 To n - 1
            For j = 0 To n - 1
                cells(i, j).Text = ""
                cells(i, j).BackColor = Color.Silver
                cells(i, j).ForeColor = Color.Blue
                board(i, j) = -1
                solution(i, j) = -1
            Next
        Next
        RandomizeBoardState(0, n, board)
        For i = 0 To n - 1
            For j = 0 To n - 1
                solution(i, j) = board(i, j)
                cells(i, j).Text = board(i, j).ToString()
                cells(i, j).Tag = False
                cellChoices.Add(New KeyValuePair(Of Integer, Integer)(i, j))
            Next
        Next
        emptyCellsCnt = rn.Next(n ^ 2 \ 3, n ^ 2 * 2 \ 3 + 1)
        For i = 1 To emptyCellsCnt
            Dim index As Integer = rn.Next(cellChoices.Count)
            cells(cellChoices(index).Key, cellChoices(index).Value).Text = ""
            cells(cellChoices(index).Key, cellChoices(index).Value).Tag = True
            board(cellChoices(index).Key, cellChoices(index).Value) = -1
            cellChoices.RemoveAt(index)
        Next
    End Sub

    Private Sub Binaris1001_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For i = 0 To n ^ 2 - 1
            Dim rw As Integer = i \ n, cl As Integer = i Mod n
            If i < 10 Then
                cells(rw, cl) = DirectCast(Controls.Find("cell0" & i, False).First, Button)
            Else
                cells(rw, cl) = DirectCast(Controls.Find("cell" & i, False).First, Button)
            End If
            cells(rw, cl).BackColor = Color.Yellow
            cells(rw, cl).ForeColor = Color.Red
            cells(rw, cl).Text = ""
            board(rw, cl) = -1
            solution(rw, cl) = -1
        Next
        start = False
        FillBoard(n, board, cells)
        startBtn.Enabled = True
        restartBtn.Enabled = False
        randomBtn.Enabled = True
    End Sub

    Private Function IsSolved(n As Integer, board(,) As Integer, solution(,) As Integer)
        For i = 0 To n - 1
            For j = 0 To n - 1
                If board(i, j) <> solution(i, j) Then
                    Return False
                End If
            Next
        Next
        Return True
    End Function

    Private Sub button_Click(sender As Object, e As EventArgs) Handles cell35.Click, cell34.Click, cell33.Click, cell32.Click, cell31.Click, cell30.Click, cell29.Click, cell28.Click, cell27.Click, cell26.Click, cell25.Click, cell24.Click, cell23.Click, cell22.Click, cell21.Click, cell20.Click, cell19.Click, cell18.Click, cell17.Click, cell16.Click, cell15.Click, cell14.Click, cell13.Click, cell12.Click, cell11.Click, cell10.Click, cell09.Click, cell08.Click, cell07.Click, cell06.Click, cell05.Click, cell04.Click, cell03.Click, cell02.Click, cell01.Click, cell00.Click
        If start Then
            Dim cellName As String = CType(sender, Button).Name
            Dim cellNum As Integer = CType(cellName.Substring(4), Integer)
            Dim row As Integer = cellNum \ n, col As Integer = cellNum Mod n
            If cells(row, col).Tag Then
                For i = 0 To n - 1
                    For j = 0 To n - 1
                        cells(i, j).BackColor = Color.Silver
                    Next
                Next
                cells(row, col).ForeColor = Color.Black
                board(row, col) += 1
                If board(row, col) = 2 Then
                    cells(row, col).Text = ""
                    board(row, col) = -1
                Else
                    Dim invalidRowsIdx As List(Of Integer) = InvalidRows(n, board, row, col, board(row, col))
                    Dim invalidColsIdx As List(Of Integer) = InvalidCols(n, board, row, col, board(row, col))
                    cells(row, col).Text = board(row, col).ToString()
                    For Each rowIdx In invalidRowsIdx
                        For j = 0 To n - 1
                            cells(rowIdx, j).BackColor = Color.Red
                        Next
                    Next
                    For Each colIdx In invalidColsIdx
                        For i = 0 To n - 1
                            cells(i, colIdx).BackColor = Color.Red
                        Next
                    Next
                End If
                If IsSolved(n, board, solution) Then
                    start = False
                    MessageBox.Show("Solved!")
                End If
            End If
        End If
    End Sub

    Private Sub startBtn_Click(sender As Object, e As EventArgs) Handles startBtn.Click
        start = True
        startBtn.Enabled = False
        restartBtn.Enabled = True
    End Sub

    Private Sub restartBtn_Click(sender As Object, e As EventArgs) Handles restartBtn.Click
        For i = 0 To n - 1
            For j = 0 To n - 1
                cells(i, j).BackColor = Color.Silver
                cells(i, j).ForeColor = Color.Blue
                If cells(i, j).Tag Then
                    cells(i, j).Text = ""
                    board(i, j) = -1
                End If
            Next
        Next
        start = False
        startBtn.Enabled = True
        restartBtn.Enabled = False
    End Sub

    Private Sub randomBtn_Click(sender As Object, e As EventArgs) Handles randomBtn.Click
        If Not start Then
            FillBoard(n, board, cells)
        End If
    End Sub
End Class
