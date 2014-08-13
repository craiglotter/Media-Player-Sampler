Imports System.Math

Public Class Image_Display
    Inherits System.Windows.Forms.Form


    Private interval As String
    Private fullscreen As Boolean
    Private filelist As ArrayList
    Private currentindex As Integer
    Private maxindex As Integer

    Private clipcomplete As Boolean
    Private cliprandomize As Boolean

    Dim cnt As Integer = 0

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal imagelist As ArrayList, ByVal timeinterval As String, ByVal showfullscreen As Boolean, ByVal clip_complete As Boolean, ByVal clip_randomize As Boolean)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        filelist = imagelist
        maxindex = filelist.Count() - 1
        currentindex = 0
        interval = timeinterval
        fullscreen = showfullscreen
        clipcomplete = clip_complete
        cliprandomize = clip_randomize


    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MediaPlayer1 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CurrentPositionTimer As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Image_Display))
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.MediaPlayer1 = New AxWMPLib.AxWindowsMediaPlayer
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.CurrentPositionTimer = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.MediaPlayer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem8, Me.MenuItem7, Me.MenuItem1, Me.MenuItem2, Me.MenuItem3, Me.MenuItem4, Me.MenuItem6, Me.MenuItem5})
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 0
        Me.MenuItem8.Text = ""
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 1
        Me.MenuItem7.Text = "-"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 2
        Me.MenuItem1.Text = "Next Video"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 3
        Me.MenuItem2.Text = "Previous Video"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 4
        Me.MenuItem3.Text = "Pause"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 5
        Me.MenuItem4.Text = "Play"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 6
        Me.MenuItem6.Text = "-"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 7
        Me.MenuItem5.Text = "Stop"
        '
        'Timer1
        '
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Black
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Location = New System.Drawing.Point(2, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(288, 262)
        Me.Panel1.TabIndex = 1
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.MediaPlayer1)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(0, 32)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(288, 230)
        Me.Panel3.TabIndex = 4
        '
        'MediaPlayer1
        '
        Me.MediaPlayer1.ContainingControl = Me
        Me.MediaPlayer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MediaPlayer1.Enabled = True
        Me.MediaPlayer1.Location = New System.Drawing.Point(0, 0)
        Me.MediaPlayer1.Name = "MediaPlayer1"
        Me.MediaPlayer1.OcxState = CType(resources.GetObject("MediaPlayer1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MediaPlayer1.Size = New System.Drawing.Size(288, 230)
        Me.MediaPlayer1.TabIndex = 1
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Black
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(288, 32)
        Me.Panel2.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 24)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "[  Exit  ]"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(80, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(200, 24)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "[  Right Click Here For More Options  ]"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CurrentPositionTimer
        '
        Me.CurrentPositionTimer.Interval = 1000
        '
        'Image_Display
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.DimGray
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.ContextMenu = Me.ContextMenu1
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Image_Display"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Image Display"
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        CType(Me.MediaPlayer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Const WM_NCHITTEST As Integer = &H84
    Private Const HTCLIENT As Integer = &H1
    Private Const HTCAPTION As Integer = &H2

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Select Case m.Msg
            Case WM_NCHITTEST
                MyBase.WndProc(m)
                If (m.Result.ToInt32 = HTCLIENT) Then
                    m.Result = IntPtr.op_Explicit(HTCAPTION)
                End If
                Exit Sub
        End Select

        MyBase.WndProc(m)
    End Sub

    Private Sub Error_Handler(ByVal message As String)
        Try
            Dim Display_Message1 As New Display_Message(message)
            Display_Message1.ShowDialog()
        Catch ex As Exception
            MsgBox("An error occurred in Media Player Sampler's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Try
            'Timer1.Stop()
            'If Not MediaPlayer1.currentMedia Is Nothing Then
            '    MediaPlayer1.Ctlcontrols.stop()
            'End If
            'MediaPlayer1.Dispose()
            ' PictureBox1.Image.Dispose()
            Me.Close()
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while closing the Image Display Screen. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Timer1.Stop()
            display_image(CStr(filelist.Item(currentindex)))
            
            Timer1.Start()
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while checking the applications' status. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub display_image(ByVal imagetoload As String)
        Timer1.Stop()
        CurrentPositionTimer.Stop()
        Try

            ' PictureBox1.Visible = False
            '            If Not (PictureBox1.Image Is Nothing) Then
            '           PictureBox1.Image.Dispose()
            '          PictureBox1.Image = Nothing
            '         End If
            'PictureBox1.Image = Image.FromFile(imagetoload)
            'PictureBox1.Height = PictureBox1.Image.Height
            'PictureBox1.Width = PictureBox1.Image.Width
            ' ContextMenu1.MenuItems(0).Text = imagetoload.Split("\")(imagetoload.Split("\").Length - 1)

            ' MsgBox(imagetoload)
            MediaPlayer1.URL = imagetoload
          

            '            MsgBox(MediaPlayer1.Ctlcontrols.currentPosition)
            'MediaPlayer1.currentMedia.duration()
            'MsgBox(MediaPlayer1.URL)
            ContextMenu1.MenuItems(0).Text = imagetoload.Split("\")(imagetoload.Split("\").Length - 1)

            If fullscreen = True Then
                Zoom_Perfect()
            End If
            '     reposition_image()
            '        PictureBox1.Visible = True

            '       PictureBox1.Refresh()
            currentindex = currentindex + 1
            If currentindex > maxindex Then
                currentindex = currentindex - maxindex - 1
            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while displaying the specified image file. The program will attempt to recover shortly.")
        End Try
        CurrentPositionTimer.Interval = 1000
        CurrentPositionTimer.Start()
        Timer1.Start()
    End Sub

    Private Sub reposition_image()
        MediaPlayer1.Top = Round(((Panel1.Height - MediaPlayer1.Height) / 2), 0)
        MediaPlayer1.Left = Round(((Panel1.Width - MediaPlayer1.Width) / 2), 0)
    End Sub

    Private Sub Zoom_Perfect()
        Try
            MediaPlayer1.stretchToFit = True
            'If PictureBox1.Image.Height <> Panel1.Height Then
            '    If PictureBox1.Image.Height >= PictureBox1.Image.Width Then
            '        PictureBox1.Height = Panel1.Height
            '        PictureBox1.Width = PictureBox1.Image.Width * (Panel1.Height / PictureBox1.Image.Height)
            '    Else
            '        PictureBox1.Width = Panel1.Width
            '        PictureBox1.Height = PictureBox1.Image.Height * (Panel1.Width / PictureBox1.Image.Width)
            '    End If

            'End If
            'PictureBox1.Top = 0
            'PictureBox1.Left = 0
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Image_Display_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Height = Screen.FromControl(Me).GetBounds(Me).Height
            Me.Width = Screen.FromControl(Me).GetBounds(Me).Width
            Panel1.Height = Screen.FromControl(Me).GetBounds(Me).Height - 8
            Panel1.Width = Screen.FromControl(Me).GetBounds(Me).Width - 8
            Me.Top = 0
            Me.Left = 0
            Panel1.Top = Round(((Me.Height - Panel1.Height) / 2), 0)
            Panel1.Left = Round(((Me.Width - Panel1.Width) / 2), 0)

            Me.Refresh()
            display_image(CStr(filelist.Item(currentindex)))
            currentindex = currentindex + 1
            If currentindex > maxindex Then
                currentindex = 0
            End If
            Timer1.Interval = CInt(interval) * 1000
            Timer1.Start()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            Timer1.Stop()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        Try
            Timer1.Start()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        display_image(CStr(filelist.Item(currentindex)))

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        currentindex = currentindex - 2
        If currentindex < 0 Then
            currentindex = currentindex + maxindex + 1
        End If
        display_image(CStr(filelist.Item(currentindex)))

    End Sub





    Protected Overrides Sub OnMouseWheel(ByVal e As System.Windows.Forms.MouseEventArgs)

        'this event is fired everytime the user moves the mouse scroll wheel
        'one click. So each notch the wheel moves, this event is fired.

        'e.Delta will return either negative 120 or positive 120(returns 
        'the number 120 on my computer at least)

        'check if delta returns a negative number, then decrease the number
        If e.Delta < 0 Then

            cnt -= 10
            'MsgBox(Screen.FromControl(Me).GetBounds(Me).Height + cnt & " grth " & PictureBox1.Height & " And " & Me.Width = Screen.FromControl(Me).GetBounds(Me).Width + cnt & " grth " & PictureBox1.Width)
            'If Screen.FromControl(Me).GetBounds(Me).Height + cnt > PictureBox1.Height And Me.Width = Screen.FromControl(Me).GetBounds(Me).Width + cnt > PictureBox1.Width Then
            If Screen.FromControl(Me).GetBounds(Me).Height + cnt > 20 And Screen.FromControl(Me).GetBounds(Me).Width + cnt > 20 Then

                Me.Height = Screen.FromControl(Me).GetBounds(Me).Height + cnt
                Me.Width = Screen.FromControl(Me).GetBounds(Me).Width + cnt
                Panel1.Height = Screen.FromControl(Me).GetBounds(Me).Height - 8 + cnt
                Panel1.Width = Screen.FromControl(Me).GetBounds(Me).Width - 8 + cnt

                Panel1.Top = Round(((Me.Height - Panel1.Height) / 2), 0)
                Panel1.Left = Round(((Me.Width - Panel1.Width) / 2), 0)

                Me.Refresh()
            Else
                cnt += 10
            End If
            'delta returns a positive number, so increase the number
        Else

            cnt += 10


            Me.Height = Screen.FromControl(Me).GetBounds(Me).Height + cnt
            Me.Width = Screen.FromControl(Me).GetBounds(Me).Width + cnt
            Panel1.Height = Screen.FromControl(Me).GetBounds(Me).Height - 8 + cnt
            Panel1.Width = Screen.FromControl(Me).GetBounds(Me).Width - 8 + cnt

            Panel1.Top = Round(((Me.Height - Panel1.Height) / 2), 0)
            Panel1.Left = Round(((Me.Width - Panel1.Width) / 2), 0)

            Me.Refresh()
        End If
        currentindex = currentindex - 1
        If currentindex < 0 Then
            currentindex = currentindex + maxindex + 1
        End If
        '  display_image(CStr(filelist.Item(currentindex)))

   
    End Sub

    'Handles key presses when the image control is in focus.
    Private Sub Navigation_Key_Handler(ByVal sender As System.Object, ByVal keyselected As System.Windows.Forms.KeyEventArgs) Handles Panel1.KeyDown, MyBase.KeyDown
        Try


            If keyselected.KeyCode = Keys.Escape Then
                Try
                    'Timer1.Stop()
                    'If Not MediaPlayer1.currentMedia Is Nothing Then
                    '    MediaPlayer1.Ctlcontrols.stop()
                    'End If
                    'MediaPlayer1.Dispose()
                    Me.Close()
                Catch ex As Exception
                    Error_Handler("An """ & ex.Message & """ error occurred while closing the Image Display Screen. The program will attempt to recover shortly.")
                End Try
                keyselected.Handled = True

                Exit Sub
            End If

            If keyselected.KeyCode = Keys.Right Then
                display_image(CStr(filelist.Item(currentindex)))
                keyselected.Handled = True

                Exit Sub
            End If

            If keyselected.KeyCode = Keys.Left Then
                currentindex = currentindex - 2
                If currentindex < 0 Then
                    currentindex = currentindex + maxindex + 1
                End If
                display_image(CStr(filelist.Item(currentindex)))
                keyselected.Handled = True

                Exit Sub
            End If
            keyselected.Handled = True

        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    'Handles mouse clicks when the image control is in focus.
    Private Sub Navigation_Mouse_Handler(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown, MyBase.MouseDown
        Try
            'If double left-click when image control is in focus, select next image
            If e.Clicks = 2 And e.Button = MouseButtons.Left Then
                display_image(CStr(filelist.Item(currentindex)))
            End If

            'If double right-click when image control is in focus, select previous image
            If e.Clicks = 2 And e.Button = MouseButtons.Right Then
                currentindex = currentindex - 2
                If currentindex < 0 Then
                    currentindex = currentindex + maxindex + 1
                End If
                display_image(CStr(filelist.Item(currentindex)))
            End If

        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    'Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click
    '    MsgBox(MediaPlayer1.Ctlcontrols.currentPosition & " of " & MediaPlayer1.currentMedia.duration)
    'End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Try

        ' PictureBox1.Image.Dispose()
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub CloseHandler(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            Timer1.Stop()
            If Not MediaPlayer1.currentMedia Is Nothing Then
                MediaPlayer1.Ctlcontrols.stop()
            End If
            MediaPlayer1.Dispose()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub CurrentPositionTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CurrentPositionTimer.Tick
        Try
            If cliprandomize = True Or clipcomplete = True Then
                If cliprandomize = True Then
                    If Not MediaPlayer1.currentMedia Is Nothing Then
                        'Timer1.Stop()
                        Randomize()
                        Dim r As Random = New Random
                        If (CInt(Math.Round(MediaPlayer1.currentMedia.duration)) - interval) > 0 Then
                            MediaPlayer1.Ctlcontrols.currentPosition = r.Next(0, CInt(Math.Round(MediaPlayer1.currentMedia.duration)) - interval)
                        End If

                        CurrentPositionTimer.Stop()
                        Timer1.Stop()
                        Timer1.Start()
                        '  End If
                    End If
                    'MediaPlayer1.Ctlcontrols.currentPosition = 10
                End If
                If clipcomplete = True Then
                    If Not MediaPlayer1.currentMedia Is Nothing Then
                        Timer1.Stop()
                        If CInt(Math.Round(MediaPlayer1.currentMedia.duration)) * 1000 > 0 Then
                            Timer1.Interval = (CInt(Math.Round(MediaPlayer1.currentMedia.duration)) * 1000) + 3000
                            'MsgBox(Timer1.Interval)
                            Timer1.Start()
                            CurrentPositionTimer.Stop()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub


End Class
