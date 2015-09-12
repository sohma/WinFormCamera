Public Class Form1

    Dim DrawFlg As Integer  'ドローモード
    Public Gr As Graphics      'グラフィックオブジェクト
    Dim intX As Integer     'マウスポインタX
    Dim intY As Integer     'マウスポインタY
    Dim intNewX As Integer  '新しいマウスポインタX
    Dim intNewY As Integer  '新しいマウスポインタY
    Private cam As WebcamSnap 'カメラ


    Private Sub Form1_Load(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles MyBase.Load
        PictureBox1.Image = New Bitmap(PictureBox1.Width, _
        PictureBox1.Height)

        'グラフィックオブジェクトを作成します。
        Gr = Graphics.FromImage(PictureBox1.Image)
        cam = New WebcamSnap
    End Sub


    Private Sub PictureBox1_MouseDown(ByVal sender As Object, _
    ByVal e As System.Windows.Forms.MouseEventArgs) _
    Handles PictureBox1.MouseDown
        Console.WriteLine("Mouse Down")
        'ドローモード
        DrawFlg = True
        intX = e.X
        intY = e.Y
        intNewX = e.X
        intNewY = e.Y

    End Sub


    Private Sub PictureBox1_MouseMove(ByVal sender As Object, _
                                      ByVal e As System.Windows.Forms.MouseEventArgs) _
                                      Handles PictureBox1.MouseMove

        intNewX = e.X
        intNewY = e.Y

        If DrawFlg = False Then
            Exit Sub

        End If

        'ドローモード
        Gr.DrawLine(Pens.Red, intX, intY, e.X, e.Y)

        '描画
        PictureBox1.Refresh()
        'ポインタをレジスト
        intX = e.X
        intY = e.Y


    End Sub

    Private Sub PictureBox1_MouseUp(ByVal sender As Object, _
    ByVal e As System.Windows.Forms.MouseEventArgs) _
    Handles PictureBox1.MouseUp

        'ドローモード初期化
        DrawFlg = False

    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, _
    ByVal e As System.ComponentModel.CancelEventArgs) _
    Handles MyBase.Closing

        'グラフィックオブジェクト解放
        Gr.Dispose()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        cam.TakePicture()
        PictureBox1.Refresh()
    End Sub

    'Private Sub PictureBox1_Paint(sender As Object, e As PaintEventArgs) Handles PictureBox1.Paint
    '    e.Graphics.DrawLine(Pens.Red, intX, intY, intNewX, intNewY)
    'End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim OutputPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        PictureBox1.Image.Save(OutputPath + "\\picture.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        cam.StopPicture()
    End Sub
End Class
