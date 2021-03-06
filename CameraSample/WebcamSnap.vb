﻿Option Explicit On 

Imports System.IO
Imports System.Runtime.InteropServices
''' <summary> 
''' スナップショットを取るクラス。ビューはできない。
''' </summary> 
''' <remarks></remarks> 
Public Class WebcamSnap
    ' ごにょごにょまじない。WinAPIのavicap32を呼ぶためのも。
    Const WM_CAP As Short = &H400S
    Const WS_CHILD = &H40000000
    Const WS_VISIBLE = &H10000000

    Const WM_CAP_DRIVER_CONNECT As Integer = WM_CAP + 10
    Const WM_CAP_DRIVER_DISCONNECT As Integer = WM_CAP + 11
    Const WM_CAP_EDIT_COPY As Integer = WM_CAP + 30
    Const WM_CAP_SEQUENCE = WM_CAP + 62
    Const WM_CAP_FILE_SAVEAS = WM_CAP + 23

    Const WM_CAP_SET_SCALE = WM_CAP + 53
    Const WM_CAP_SET_PREVIEWRATE = WM_CAP + 52
    Const WM_CAP_SET_PREVIEW = WM_CAP + 50


    Const SWP_NOMOVE = &H2S
    Const SWP_NOSIZE = 1
    Const SWP_NOZORDER = &H45
    Const HWND_BOTTOM = 1



    Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, _
        <MarshalAs(UnmanagedType.AsAny)> ByVal lParam As Object) As Integer

    Declare Function capCreateCaptureWindowA Lib "avicap32.dll" _
        (ByVal lpszWindowName As String, ByVal dwStyle As Integer, _
        ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, _
        ByVal nHeight As Short, ByVal hWndParent As Integer, _
        ByVal nID As Integer) As Integer

    Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Short, _
        ByVal lpszName As String, ByVal cbName As Integer, ByVal lpszVer As String, _
        ByVal cbVer As Integer) As Boolean

    'Declare Function SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer

    'Declare Function DestroyWindow Lib "user32" (ByVal hndw As Integer) As Boolean

    Private hWnd As Integer
    Private DriverVersion As Integer
    Private mypicture As PictureBox = New PictureBox()

    Public Devices As New List(Of String)

    ' 画像サイズ
    Public Height As Integer = 480
    Public Width As Integer = 640

    ' ファイル保存のパス
    Public OutputPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
    Public FilenamePrefix As String = "snapshot"

    Public Sub New()
        mLoadDeviceList()
    End Sub

    Private Sub mLoadDeviceList()
        Dim lsName As String = Space(100)
        Dim lsVers As String = Space(100)
        Dim lbReturn As Boolean
        Dim x As Integer = 0

        Do
            ' ドライバーの名前とバージョン取得
            lbReturn = capGetDriverDescriptionA(x, lsName, 80, lsVers, 80)

            ' ディバイスのリストを作る 
            If lbReturn Then Devices.Add(lsName.Trim)
            x += 1
            Console.WriteLine("Device No." + x.ToString + ", Name:" + lsName + ", Vers:" + lsVers)
        Loop Until lbReturn = False
    End Sub
    Public Sub TakePicture()
        For i = 0 To Me.Devices.Count - 1
            Dim lsFilename As String = Path.Combine(OutputPath, Me.FilenamePrefix & i & ".jpg")
            TakePicture(i, lsFilename)
        Next
    End Sub
    Public Sub TakePicture(ByVal iDevice As Integer)
        Me.TakePicture(iDevice, Path.Combine(OutputPath, Me.FilenamePrefix & ".jpg"))
    End Sub
    Public Sub TakePicture(ByVal iDevice As Integer, ByVal filename As String)

        Dim lhHwnd As Integer ' 親ウィンドウハンドル

        ' フォーム作成 
        Using loWindow As New System.Windows.Forms.Form

            ' チャプチャーウィンドウ作成 
            lhHwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, Me.Width, _
                       Me.Height, loWindow.Handle.ToInt32, 0)
            'lhHwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, 0, 0, _
            '                                 loWindow.Handle.ToInt32, 0)

            ' ディバイス呼び出し
            SendMessage(lhHwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0)
            'SendMessage(lhHwnd, WM_CAP_SEQUENCE, 0, 0)
            ' ディレイ 
            For i = 1 To 100
                Application.DoEvents()
            Next

            ' クリップボードにコピー 
            SendMessage(lhHwnd, WM_CAP_EDIT_COPY, 0, 0)
            Console.WriteLine("COPY" + iDevice.ToString)

            ' クリップボードをビットマップに変換 
            Dim loData As IDataObject = Clipboard.GetDataObject()
            If loData.GetDataPresent(GetType(System.Drawing.Bitmap)) Then

                ' デスクトップ保存するときはコメントアウト
                'Using loBitmap As Image = CType(loData.GetData(GetType(System.Drawing.Bitmap)), Image)
                'loBitmap.Save(filename, Imaging.ImageFormat.Jpeg)
                'Console.WriteLine("Save" + filename)
                'End Using

                ' Imageプロパティは使わない
                'Form1.PictureBox1.Image = CType(loData.GetData(GetType(System.Drawing.Bitmap)), Image)
                Dim loBitmap As Bitmap = loData.GetData(GetType(System.Drawing.Bitmap))
                Form1.Gr.DrawImage(loBitmap, 0, 0, loBitmap.Width, loBitmap.Height)

            End If

            SendMessage(lhHwnd, WM_CAP_DRIVER_DISCONNECT, iDevice, 0)

        End Using

    End Sub

End Class