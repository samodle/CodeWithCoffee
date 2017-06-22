Public Class Snake
#Region "Variables and Constants and Structure"
    Private Structure structSnake
        Dim rect As System.Drawing.Rectangle
        Dim x As Integer
        Dim y As Integer
    End Structure

    Private Enum Direction
        Rightward
        Downward
        Leftward
        Upward
    End Enum

    Private Const INITIAL_SNAKE_RECT_COUNT As Integer = 15
    Private Const COLUMN_COUNT As Integer = 65
    Private Const ROW_COUNT As Integer = 47

    Private curRecCount As Integer
    Private rects(,) As System.Drawing.Rectangle
    Private isSnakePart(,) As Boolean
    Private snake As Collection
    Private blackBrush As System.Drawing.Brush = New System.Drawing.SolidBrush(System.Drawing.Color.Black) 'FromArgb(0, 255, 0))
    Private snakeBrush As System.Drawing.Brush = New System.Drawing.SolidBrush(System.Drawing.Color.Cyan) 'FromArgb(0, 255, 0))
    Private backBrush As System.Drawing.Brush = New System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(1, 1, 1))
    Private tokenBrush As System.Drawing.Brush = New System.Drawing.SolidBrush(System.Drawing.Color.Violet) 'Color.Red)
    Private curDirection As Direction
    Private buffer As System.Drawing.Bitmap
    Private columnCount As Integer
    Private rowCount As Integer
    Private snakePoints As Integer
    Private snakeSpeed As Double
    Private snakeLength As Integer
    Private token As System.Drawing.Rectangle
    Private forceClose As Boolean = False

    '////////////////////////////////////////////////////////////
    Private prstoryMode As Boolean = False
    Private prBrush As System.Drawing.Brush
    Private prBrush1 As System.Drawing.Brush = New System.Drawing.SolidBrush(System.Drawing.Color.Violet) 'Color.Red)
    Private prBrush2 As System.Drawing.Brush = New System.Drawing.SolidBrush(System.Drawing.Color.LightGray) 'Color.Red)
    Private drawingCounter As Integer = 0
    Private x_PR As Integer
    Private y_PR As Integer

    Private Snake_Speed As Integer = 100

    '////////////////////////////////////////////////////////////
#End Region

    Private Function xyToRectIndex(ByVal X As Integer, ByVal Y As Integer) As Integer
        Return (Y * (columnCount)) + X
    End Function

    Private Sub rectIndexToXY(ByVal Index As Integer, ByRef X As Integer, ByRef Y As Integer)
        X = Index Mod (columnCount)
        Y = Index \ (columnCount)
    End Sub

    Private Sub initSnake()

        Dim x As Integer
        Dim y As Integer
        Dim i As Integer
        Dim index As Integer
        Dim sSnake As structSnake
        snake = New Collection

        x = ((columnCount) - 10) \ 2
        y = ((rowCount) - 6) \ 2

        'APRIL FOOLS
        '  If isAPRILFOOLS And My.Settings.APRILFOOLS_RunCount < 10 Then 'Mod 2 = 0 Then
        '  blackBrush = New System.Drawing.SolidBrush(System.Drawing.Color.BlueViolet) 'FromArgb(0, 255, 0))
        '  snakeBrush = New System.Drawing.SolidBrush(System.Drawing.Color.Orange) 'FromArgb(0, 255, 0))
        '  backBrush = New System.Drawing.SolidBrush(System.Drawing.Color.HotPink)
        '  tokenBrush = New System.Drawing.SolidBrush(System.Drawing.Color.HotPink) 'Color.Red)
        '  prBrush1 = New System.Drawing.SolidBrush(System.Drawing.Color.Green) 'Color.Red)
        '  prBrush2 = New System.Drawing.SolidBrush(System.Drawing.Color.LightPink) 'Color.Red)
        ' snakeSpeed = 300
        '   End If
        ''''''''''''


        Dim snakePosition As Point = New Point(x, y)
        index = xyToRectIndex(x, y)

        For i = 1 To INITIAL_SNAKE_RECT_COUNT
            rectIndexToXY(index + (i - 1), x, y)
            sSnake.rect = rects(x, y)
            sSnake.x = x
            sSnake.y = y
            snake.Add(sSnake)
        Next
        x_PR = x - 24
        y_PR = y + 3

        snakeLength = INITIAL_SNAKE_RECT_COUNT
        snakeSpeed = 100
    End Sub

    Private Sub selectRectangles()

        Dim g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(My.Resources.back)
        Dim i As Integer
        Dim sSnake As structSnake

        For i = 1 To INITIAL_SNAKE_RECT_COUNT
            sSnake = snake(i)
            g.FillRectangle(snakeBrush, sSnake.rect)
            isSnakePart(sSnake.x, sSnake.y) = True
        Next

        buffer = New System.Drawing.Bitmap(My.Resources.back)

        g.Dispose()
        Refresh()

    End Sub

    Private Sub initRectangles()
        Dim i As Integer
        Dim j As Integer

        columnCount = COLUMN_COUNT
        rowCount = ROW_COUNT

        ReDim rects(columnCount, rowCount)
        ReDim isSnakePart(columnCount, rowCount)

        For j = 0 To rowCount
            For i = 0 To columnCount
                rects(i, j) = New System.Drawing.Rectangle((i * 10) + 1, (j * 10) + 1, 9, 9)
                isSnakePart(i, j) = False
            Next
        Next
        '   ss.Items("tss0").Text = "Screen Size: " & CStr(columnCount) & " X " & CStr(rowCount)
    End Sub

    Private Sub initialize()
        curRecCount = INITIAL_SNAKE_RECT_COUNT
        curDirection = Direction.Leftward
        snakePoints = 0
        initRectangles()
        initSnake()
        selectRectangles()
        setToken()
        tmr.Interval = 50
        tmr.Enabled = True
    End Sub

    Private Sub setToken()
        Randomize()
        Dim x As Integer
        Dim y As Integer
        Dim g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(buffer)

        x = CInt(Rnd() * columnCount)
        Do While x > columnCount Or isSnakePart(x, y) = True
            x = CInt(Rnd() * columnCount)
        Loop

        y = CInt(Rnd() * rowCount)
        Do While y > rowCount Or isSnakePart(x, y) = True
            y = CInt(Rnd() * rowCount)
        Loop

        token = rects(x, y)
        '  ss.Items("tss1").Text = "Token Location: ( " & CStr(x) & " , " & CStr(y) & " )"

        g.FillEllipse(tokenBrush, token)
        Refresh()
        g.Dispose()

    End Sub

    Private Sub main_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initialize()
    End Sub

    Private Sub doneWithSname() 'Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub main_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        Select Case e.KeyCode
            Case Forms.Keys.Down
                If Not (curDirection = Direction.Downward Or curDirection = Direction.Upward) Then
                    curDirection = Direction.Downward
                End If
            Case Forms.Keys.Left
                If Not (curDirection = Direction.Leftward Or curDirection = Direction.Rightward) Then
                    curDirection = Direction.Leftward
                End If
            Case Forms.Keys.Right
                If Not (curDirection = Direction.Rightward Or curDirection = Direction.Leftward) Then
                    curDirection = Direction.Rightward
                End If
            Case Forms.Keys.Up
                If Not (curDirection = Direction.Upward Or curDirection = Direction.Downward) Then
                    curDirection = Direction.Upward
                End If
        End Select

    End Sub

    Private Sub moveSnake()

        Dim sSnake As structSnake
        Dim x As Integer
        Dim y As Integer
        Dim rect As System.Drawing.Rectangle = New System.Drawing.Rectangle()
        Dim g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(buffer)

        tmr.Enabled = False

        ' Remove the last snake square.
        sSnake = snake(snake.Count)
        g.FillRectangle(backBrush, sSnake.rect)
        snake.Remove(snake.Count)
        isSnakePart(sSnake.x, sSnake.y) = False

        ' Get the index of the snake's first square.
        sSnake = snake.Item(1)

        x = sSnake.x
        y = sSnake.y

        Select Case curDirection
            Case Direction.Downward
                y = y + 1
                If y > rowCount Then y = 0
            Case Direction.Leftward
                x = x - 1
                If x < 0 Then x = columnCount
            Case Direction.Rightward
                x = x + 1
                If x > columnCount Then x = 0
            Case Direction.Upward
                y = y - 1
                If y < 0 Then y = rowCount
        End Select

        If isSnakePart(x, y) = True Then
            tmr.Enabled = False
            initializeForPR()
            If My.Computer.Keyboard.CtrlKeyDown Then
                setToken()
                '  Snake_Speed += 50
            Else
                prstoryMode = True
                Exit Sub
                Me.Close()
            End If

            ' If MessageBox.Show(msgBoxText, msgBoxTitle, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question, System.Windows.Forms.MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
            'Else

            ' End If

        End If

        rect = rects(x, y)

        sSnake.x = x
        sSnake.y = y
        sSnake.rect = rect
        isSnakePart(x, y) = True

        g.FillRectangle(snakeBrush, sSnake.rect)
        Me.BackgroundImage = buffer

        ' Add the snake square to the beginning of the collection.
        snake.Add(sSnake, , 1)

        If rects(x, y).Equals(CObj(token)) Then

            snakePoints += 1


            If snakePoints Mod 5 = 0 Then
                sSnake = snake.Item(snake.Count)
                Select Case curDirection
                    Case Direction.Downward
                        sSnake.y -= 1
                    Case (Direction.Leftward)
                        sSnake.x += 1
                    Case Direction.Rightward
                        sSnake.x -= 1
                    Case Direction.Upward
                        sSnake.y += 1
                End Select

                sSnake.rect = rects(sSnake.x, sSnake.y)
                g.FillRectangle(snakeBrush, sSnake.rect)
                Me.BackgroundImage = buffer
                snake.Add(sSnake, , , snake.Count)
                snakeLength = snake.Count '+ 3

                tmr.Interval -= 2
                If tmr.Interval < 0 Then tmr.Interval = 1

                snakeSpeed += 70
                snakeSpeed = Snake_Speed + (((50 - tmr.Interval) / 50) * 100) '+ 70

            End If

            setToken()

        End If

        Refresh()

        tmr.Enabled = True

    End Sub

    Private Sub tmr_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmr.Tick
        If prstoryMode Then
            drawPRstory()
        Else
            If shouldSnakeClose Then
                tmr.Enabled = False
                initializeForPR()
                prstoryMode = True
                forceClose = True
            End If
            moveSnake()
        End If

        System.Windows.Forms.Application.DoEvents()
    End Sub

    '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||\

    Private Sub drawPRstory()
        Dim g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(buffer)
        Dim isDraw As Boolean = True
        tmr.Enabled = False
        Select Case drawingCounter
            'P
            Case 0
                prBrush = prBrush1
            Case 1
                y_PR -= 1
            Case 2
                y_PR -= 1
            Case 3
                y_PR -= 1
            Case 4
                y_PR -= 1
            Case 5
                y_PR -= 1
            Case 6
                x_PR += 1
            Case 7
                x_PR += 1
            Case 8
                y_PR += 1
            Case 9
                y_PR += 1
            Case 10
                x_PR -= 1
                ''''''''
                'R
            Case 11
                y_PR -= 2
                x_PR += 3
            Case 12
                y_PR += 1
            Case 13
                y_PR += 1
            Case 14
                y_PR -= 2
                x_PR += 1
            Case 15
                x_PR += 1
                '''''''''''
                'S
            Case 16
                prBrush = prBrush2
                x_PR += 2
            Case 17
                x_PR += 1
            Case 18
                x_PR += 1
                downOne()
            Case 19
                x_PR -= 1
                downOne()
            Case 20
                x_PR -= 1
            Case 21
                upOne()
                upOne()
                upOne()
                leftOne()
            Case 22
                upOne()
                rightOne()
            Case 23
                rightOne()

            Case 24
                rightOne()
            Case 25
                downOne()
                downOne()
                downOne()
                downOne()
                leftOne()
                leftOne()
                leftOne()
                'T
            Case 26
                rightOne()
                rightOne()
                rightOne()
                rightOne()
                rightOne()

            Case 27
                upOne()
            Case 28
                upOne()
            Case 29
                upOne()
            Case 30
                leftOne()
            Case 31
                rightOne()
                rightOne()

            Case 32
                leftOne()
                upOne()

            Case 33
                upOne()
            Case 34
                downOne()
                downOne()
                downOne()
                rightOne()
                rightOne()

            Case 35
                rightOne()
            Case 36
                rightOne()
            Case 37
                downOne()
            Case 38
                downOne()
            Case 39
                leftOne()
            Case 40
                leftOne()
            Case 41
                upOne()
            Case 42
                x_PR += 4
                upOne()
            Case 43
                rightOne()
            Case 44
                rightOne()
            Case 45
                x_PR -= 2
                downOne()
            Case 46
                downOne()
                'Y
            Case 47
                x_PR += 4
                upOne()
                upOne()
            Case 48
                downOne()
                rightOne()
            Case 49
                downOne()
                rightOne()
            Case 50
                upOne()
                rightOne()

            Case 51
                upOne()
                rightOne()
            Case 52
                downOne()
                downOne()
                downOne()
                x_PR -= 3
            Case 53
                downOne()
                leftOne()
            Case 54
                downOne()
                leftOne()
            Case 55

                drawingCounter = -1
                prstoryMode = False
                isDraw = False
        End Select
        drawingCounter += 1
        If isDraw Then
            g.FillEllipse(prBrush, rects(x_PR, y_PR))
            Me.BackgroundImage = buffer
            Refresh()
            g.Dispose()
            tmr.Enabled = True
        Else
            If forceClose Then
                Me.Close()
            Else
                Me.Close()
            End If



        End If


    End Sub


    Private Sub initializeForPR()
        curRecCount = INITIAL_SNAKE_RECT_COUNT
        curDirection = Direction.Leftward
        snakePoints = 0
        initRectangles()
        selectRectangles2()
        tmr.Interval = 15
        tmr.Enabled = True
    End Sub
    Private Sub selectRectangles2()
        Dim g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(My.Resources.back)
        buffer = New System.Drawing.Bitmap(My.Resources.back)
        g.Dispose()
        Refresh()

    End Sub
    Private Sub leftOne()
        x_PR -= 1
    End Sub
    Private Sub rightOne()
        x_PR += 1
    End Sub
    Private Sub upOne()
        y_PR -= 1
    End Sub
    Private Sub downOne()
        y_PR += 1
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
End Class