Public Class Form1
    'Samantha Zhang
    Dim intcount(6) As Integer
    Dim intboard(5, 6) As Integer
    Dim intplayer As Integer
    Dim picboard(5, 6) As PictureBox
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Integer
        For i = 0 To 6
            intcount(i) = 5
        Next
        'intcount is row
        'i is column/btn
        picboard(0, 0) = Pic00
        picboard(0, 1) = Pic01
        picboard(0, 2) = Pic02
        picboard(0, 3) = Pic03
        picboard(0, 4) = Pic04
        picboard(0, 5) = Pic05
        picboard(0, 6) = Pic06
        picboard(1, 0) = Pic10
        picboard(1, 1) = Pic11
        picboard(1, 2) = Pic12
        picboard(1, 3) = Pic13
        picboard(1, 4) = Pic14
        picboard(1, 5) = Pic15
        picboard(1, 6) = Pic16
        picboard(2, 0) = Pic20
        picboard(2, 1) = Pic21
        picboard(2, 2) = Pic22
        picboard(2, 3) = Pic23
        picboard(2, 4) = Pic24
        picboard(2, 5) = Pic25
        picboard(2, 6) = Pic26
        picboard(3, 0) = Pic30
        picboard(3, 1) = Pic31
        picboard(3, 2) = Pic32
        picboard(3, 3) = Pic33
        picboard(3, 4) = Pic34
        picboard(3, 5) = Pic35
        picboard(3, 6) = Pic36
        picboard(4, 0) = Pic40
        picboard(4, 1) = Pic41
        picboard(4, 2) = Pic42
        picboard(4, 3) = Pic43
        picboard(4, 4) = Pic44
        picboard(4, 5) = Pic45
        picboard(4, 6) = Pic46
        picboard(5, 0) = Pic50
        picboard(5, 1) = Pic51
        picboard(5, 2) = Pic52
        picboard(5, 3) = Pic53
        picboard(5, 4) = Pic54
        picboard(5, 5) = Pic55
        picboard(5, 6) = Pic56

        intplayer = 1


    End Sub

    Private Sub Btn_Click(sender As Object, e As EventArgs) Handles Btn0.Click, Btn1.Click, Btn2.Click, Btn3.Click, Btn4.Click, Btn5.Click, Btn6.Click
        Dim btn As Button = sender
        Dim intColumn As Integer = Val(btn.Tag)
        Dim IntRow As Integer = intcount(intColumn)

        'Column 0-6 intcount = 5 if intcount <0 then error

        If IntRow < 0 Then
            MessageBox.Show("Column full")
        Else
            If intplayer = 1 Then
                picboard(IntRow, intColumn).Image = Image.FromFile("Player1.bmp")

            Else
                picboard(IntRow, intColumn).Image = Image.FromFile("Player2.bmp")
            End If

            intcount(intColumn) = intcount(intColumn) - 1
            intboard(IntRow, intColumn) = intplayer

            If (isWinner(intboard)) Then
                MessageBox.Show("Game Over")
                End
            Else
                If intplayer = 1 Then
                    intplayer = 2
                Else
                    intplayer = 1
                End If
            End If

        End If
    End Sub
    Function isWinner(ByRef intboard(,) As Integer) As Boolean
        'check all rows() and then columns(3) and then diagnals(24)... 4 checks 0=1and2?
        Dim introw As Integer
        For introw = 0 To 5
            If (intboard(introw, 0) = intboard(introw, 1) And intboard(introw, 1) = intboard(introw, 2) And intboard(introw, 2) = intboard(introw, 3) And (intboard(introw, 0) = 1 Or intboard(introw, 0) = 2)) Then
                Return True 'Winner
            ElseIf (intboard(introw, 1) = intboard(introw, 2) And intboard(introw, 2) = intboard(introw, 3) And intboard(introw, 3) = intboard(introw, 4) And (intboard(introw, 1) = 1 Or intboard(introw, 1) = 2)) Then
                Return True
            ElseIf (intboard(introw, 2) = intboard(introw, 3) And intboard(introw, 3) = intboard(introw, 4) And intboard(introw, 4) = intboard(introw, 5) And (intboard(introw, 2) = 1 Or intboard(introw, 3) = 2)) Then
                Return True
            ElseIf (intboard(introw, 3) = intboard(introw, 4) And intboard(introw, 4) = intboard(introw, 5) And intboard(introw, 5) = intboard(introw, 6) And (intboard(introw, 3) = 1 Or intboard(introw, 3) = 2)) Then
                Return True

            End If
        Next introw

        Dim intCol As Integer
        For intCol = 0 To 6
            If (intboard(0, intCol) = intboard(1, intCol) And intboard(1, intCol) = intboard(2, intCol) And intboard(2, intCol) = intboard(3, intCol) And (intboard(0, intCol) = 1 Or intboard(0, intCol) = 2)) Then
                Return True 'Winner
            ElseIf (intboard(1, intCol) = intboard(2, intCol) And intboard(2, intCol) = intboard(3, intCol) And intboard(3, intCol) = intboard(4, intCol) And (intboard(1, intCol) = 1 Or intboard(1, intCol) = 2)) Then
                Return True
            ElseIf (intboard(2, intCol) = intboard(3, intCol) And intboard(3, intCol) = intboard(4, intCol) And intboard(4, intCol) = intboard(5, intCol) And (intboard(2, intCol) = 1 Or intboard(2, intCol) = 2)) Then
                Return True

            End If
        Next intCol

        If (intboard(0, 0) = intboard(1, 1) And intboard(1, 1) = intboard(2, 2) And intboard(2, 2) = intboard(3, 3) And (intboard(0, 0) = 1 Or intboard(0, 0) = 2)) Then
            Return True
        End If

        If (intboard(1, 1) = intboard(2, 2) And intboard(2, 2) = intboard(3, 3) And intboard(3, 3) = intboard(4, 4) And (intboard(1, 1) = 1 Or intboard(1, 1) = 2)) Then
            Return True
        End If

        If (intboard(2, 2) = intboard(3, 3) And intboard(3, 3) = intboard(4, 4) And intboard(4, 4) = intboard(5, 5) And (intboard(2, 2) = 1 Or intboard(2, 2) = 2)) Then
            Return True
        End If

        If (intboard(0, 1) = intboard(1, 2) And intboard(1, 2) = intboard(2, 3) And intboard(2, 3) = intboard(3, 4) And (intboard(0, 1) = 1 Or intboard(0, 1) = 2)) Then
            Return True
        End If

        If (intboard(1, 2) = intboard(2, 3) And intboard(2, 3) = intboard(3, 4) And intboard(3, 4) = intboard(4, 5) And (intboard(1, 2) = 1 Or intboard(1, 2) = 2)) Then
            Return True
        End If

        If (intboard(2, 3) = intboard(3, 4) And intboard(3, 4) = intboard(4, 5) And intboard(4, 5) = intboard(5, 6) And (intboard(2, 3) = 1 Or intboard(2, 3) = 2)) Then
            Return True
        End If

        If (intboard(1, 0) = intboard(2, 1) And intboard(2, 1) = intboard(3, 2) And intboard(3, 2) = intboard(4, 3) And (intboard(1, 0) = 1 Or intboard(1, 0) = 2)) Then
            Return True
        End If

        If (intboard(2, 1) = intboard(3, 2) And intboard(3, 2) = intboard(4, 3) And intboard(4, 3) = intboard(5, 4) And (intboard(2, 1) = 1 Or intboard(2, 1) = 2)) Then
            Return True
        End If

        If (intboard(2, 0) = intboard(3, 1) And intboard(3, 1) = intboard(4, 2) And intboard(4, 2) = intboard(5, 3) And (intboard(2, 0) = 1 Or intboard(2, 0) = 2)) Then
            Return True
        End If

        If (intboard(0, 2) = intboard(1, 3) And intboard(1, 3) = intboard(2, 4) And intboard(2, 4) = intboard(3, 5) And (intboard(0, 2) = 1 Or intboard(0, 2) = 2)) Then
            Return True
        End If

        If (intboard(1, 3) = intboard(2, 4) And intboard(2, 4) = intboard(3, 5) And intboard(3, 5) = intboard(4, 6) And (intboard(1, 3) = 1 Or intboard(1, 3) = 2)) Then
            Return True
        End If

        If (intboard(0, 3) = intboard(1, 4) And intboard(1, 4) = intboard(2, 5) And intboard(2, 5) = intboard(3, 6) And (intboard(0, 3) = 1 Or intboard(0, 3) = 2)) Then
            Return True
        End If



        If (intboard(0, 3) = intboard(1, 2) And intboard(1, 2) = intboard(2, 1) And intboard(2, 1) = intboard(3, 0) And (intboard(0, 3) = 1 Or intboard(0, 3) = 2)) Then
            Return True
        End If

        If (intboard(0, 4) = intboard(1, 3) And intboard(1, 3) = intboard(2, 2) And intboard(2, 2) = intboard(3, 1) And (intboard(0, 4) = 1 Or intboard(0, 4) = 2)) Then
            Return True
        End If

        If (intboard(0, 5) = intboard(1, 4) And intboard(1, 4) = intboard(2, 3) And intboard(2, 3) = intboard(3, 2) And (intboard(0, 5) = 1 Or intboard(0, 5) = 2)) Then
            Return True
        End If

        If (intboard(0, 6) = intboard(1, 5) And intboard(1, 5) = intboard(2, 4) And intboard(2, 4) = intboard(3, 3) And (intboard(0, 6) = 1 Or intboard(0, 6) = 2)) Then
            Return True
        End If

        If (intboard(1, 3) = intboard(2, 2) And intboard(2, 2) = intboard(3, 1) And intboard(3, 1) = intboard(4, 0) And (intboard(1, 3) = 1 Or intboard(1, 3) = 2)) Then
            Return True
        End If

        If (intboard(1, 4) = intboard(2, 3) And intboard(2, 3) = intboard(3, 2) And intboard(3, 2) = intboard(4, 1) And (intboard(1, 4) = 1 Or intboard(1, 4) = 2)) Then
            Return True
        End If

        If (intboard(1, 5) = intboard(2, 4) And intboard(2, 4) = intboard(3, 3) And intboard(3, 3) = intboard(4, 2) And (intboard(1, 5) = 1 Or intboard(1, 5) = 2)) Then
            Return True
        End If

        If (intboard(1, 6) = intboard(2, 5) And intboard(2, 5) = intboard(3, 4) And intboard(3, 4) = intboard(4, 3) And (intboard(1, 6) = 1 Or intboard(1, 6) = 2)) Then
            Return True
        End If

        If (intboard(2, 6) = intboard(3, 5) And intboard(3, 5) = intboard(4, 4) And intboard(4, 4) = intboard(5, 3) And (intboard(2, 6) = 1 Or intboard(2, 6) = 2)) Then
            Return True
        End If

        If (intboard(2, 5) = intboard(3, 4) And intboard(3, 4) = intboard(4, 3) And intboard(4, 3) = intboard(5, 2) And (intboard(2, 5) = 1 Or intboard(2, 5) = 2)) Then
            Return True
        End If

        If (intboard(2, 4) = intboard(3, 3) And intboard(3, 3) = intboard(4, 2) And intboard(4, 2) = intboard(5, 1) And (intboard(2, 4) = 1 Or intboard(2, 4) = 2)) Then
            Return True
        End If

        If (intboard(5, 0) = intboard(4, 1) And intboard(4, 1) = intboard(3, 2) And intboard(3, 2) = intboard(2, 3) And (intboard(5, 0) = 1 Or intboard(5, 0) = 2)) Then
            Return True
        End If

        Return False
    End Function
End Class
