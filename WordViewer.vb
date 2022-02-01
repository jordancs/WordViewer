Imports System.IO
Imports System.Threading
Imports System.Runtime.InteropServices

Public Class WordViewer

#Region "Win32 API's"
    Const MF_BYPOSITION As Integer = &H400
    Const MF_REMOVE As Integer = &H1000

    Const SWP_DRAWFRAME As Integer = &H20
    Const SWP_NOMOVE As Integer = &H2
    Const SWP_NOSIZE As Integer = &H1
    Const SWP_NOZORDER As Integer = &H4


    <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
    Private Shared Function SetParent _
        (ByVal hWndChild As IntPtr, _
        ByVal hWndNewParent As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
    Private Shared Function FindWindow _
        (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
    Private Shared Function SetWindowPos _
        (ByVal hwnd As IntPtr, _
        ByVal hWndInsertAfter As IntPtr, _
        ByVal x As Integer, _
        ByVal y As Integer, _
        ByVal cx As Integer, _
        ByVal cy As Integer, _
        ByVal uFlags As Integer) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
    Private Shared Function MoveWindow _
        (ByVal hwnd As IntPtr, _
        ByVal x As Integer, _
        ByVal y As Integer, _
        ByVal nWidth As Integer, _
        ByVal nHeight As Integer, _
        ByVal bRepaint As Boolean) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True, ExactSpelling:=True)> _
    Public Shared Function IsWindow(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

#End Region

    Public Enum ViewMode
        NormalView = 1
        OutlineView = 2
        PrintView = 3
        PrintPreview = 4
        MasterView = 5
        WebView = 6
        ReadingView = 7
    End Enum

    Dim hWordApp As IntPtr = IntPtr.Zero
    'Dim hParentPrev As IntPtr = IntPtr.Zero
    Dim oWordApp As Object = Nothing
    Dim oWordDoc As Object = Nothing
    Dim mutex As New Semaphore(1, 1)
    Dim t1 As Thread = Nothing
    Dim fileName As String
    Dim version As Double = 0.0

    Private Delegate Sub ThreadSetParentCallback(ByVal child As IntPtr)
    Private Delegate Sub ThreadSetWindowPosCallback(ByVal hWND As IntPtr)

    Private Sub ThreadSetParent(ByVal child As IntPtr)
        If (Not Me.InvokeRequired) Then

            'hParentPrev = SetParent(child, Me.Handle)
            SetParent(child, Me.Handle)
        Else
            Me.Invoke(New ThreadSetParentCallback(AddressOf Me.ThreadSetParent), New Object() {child})
        End If
    End Sub

    Private Sub ThreadSetWindowPos(ByVal hWND As IntPtr)
        If (Not Me.InvokeRequired) Then
            SetWindowPos(hWND, Me.Handle, 0, 0, Me.Bounds.Width, Me.Bounds.Height, _
                 SWP_NOZORDER Or SWP_NOMOVE Or SWP_DRAWFRAME Or SWP_NOSIZE)
        Else
            Me.Invoke(New ThreadSetWindowPosCallback(AddressOf Me.ThreadSetWindowPos), New Object() {hWND})
        End If
    End Sub

    Protected Overrides Sub OnResize(ByVal e As System.EventArgs)
        OnResizeControl()
    End Sub

    Private Sub OnResizeControl()
        Dim borderWidth As Integer = SystemInformation.Border3DSize.Width
        Dim borderHeight As Integer = SystemInformation.Border3DSize.Height
        Dim captionHeight As Integer = SystemInformation.CaptionHeight
        Dim statusHeight As Integer = SystemInformation.ToolWindowCaptionHeight
        'If (hWordApp <> IntPtr.Zero) Then

        If IsWindow(Me.hWordApp) Then
            If (version >= 12.0) Then
                MoveWindow(hWordApp,
                           0 * borderWidth,
                           0,
                           Me.Bounds.Width + 0 * borderWidth,
                           Me.Bounds.Height,
                           True)
            Else
                MoveWindow(hWordApp, _
                           -2 * borderWidth, _
                           -2 * borderHeight - captionHeight, _
                           Me.Bounds.Width + 4 * borderWidth, _
                           Me.Bounds.Height + captionHeight + 4 * borderHeight + statusHeight, _
                           True)
            End If
        End If
    End Sub

    Private Sub SetViewControls(ByVal visible As Boolean, ByVal viewType As ViewMode)
        'If Me.oWordApp Is Nothing Then
        If Not IsWindow(Me.hWordApp) Then
            Exit Sub
        End If

        If Me.oWordApp.Documents.Count > 0 Then
            Me.oWordApp.ActiveWindow.DisplayRightRuler = visible
            Me.oWordApp.ActiveWindow.DisplayVerticalRuler = visible
            Me.oWordApp.ActiveWindow.DisplayScreenTips = visible
            Me.oWordApp.ActiveWindow.ActivePane.DisplayRulers = visible
            Me.oWordApp.ActiveWindow.ActivePane.View.Type = viewType
        End If
    End Sub

    Private Sub SetCommandBarControls(ByVal visible As Boolean)
        If Not IsWindow(Me.hWordApp) Then
            Exit Sub
        End If

        Me.oWordApp.CommandBars.AdaptiveMenus = visible

        Dim app As Object = Me.oWordApp.Application

        For i As Integer = 1 To app.CommandBars.Count
            Dim cbName As String = app.CommandBars(i).Name
            If cbName.Equals("Formatting") Then
                app.CommandBars(i).Enabled = visible
            End If
            If cbName.Equals("Standard") Then
                'app.CommandBars(i).Enabled = visible
                For j As Integer = 1 To 2
                    app.CommandBars(i).Controls(j).Enabled = visible
                Next
            End If
            If cbName.Equals("Menu Bar") Then
                app.CommandBars(i).Enabled = visible
            End If
        Next
    End Sub

    Public Sub Restore()
        If IsWindow(Me.hWordApp) Then
            Me.oWordApp.Visible = False
            Me.SetViewControls(True, ViewMode.PrintView)
            Me.SetCommandBarControls(True)
            Me.oWordApp.Quit(Nothing, Nothing, Nothing)
        End If
        Me.oWordApp = Nothing
        Me.t1 = Nothing
        Me.hWordApp = IntPtr.Zero
    End Sub

    Public Sub LoadDocument(ByVal strFileName As String)
        Me.fileName = strFileName
        If (Not Me.t1 Is Nothing) Then
            If (Me.t1.IsAlive) Then
                Exit Sub
            End If
        End If
        Me.t1 = New Thread(New ThreadStart(AddressOf Me.LoadDocumentThread))
        Me.t1.IsBackground = True
        Me.t1.Priority = ThreadPriority.Highest
        Me.t1.Start()
    End Sub

    Private Sub SetWindowWord(ByVal caption As String)
        Me.oWordApp.Caption = String.Format("{0}_fromViteApp", caption)
        Me.hWordApp = FindWindow(Nothing, Me.oWordApp.Caption)
        'Me.hWordApp = FindWindow("OpusApp", Nothing)

        If (Me.hWordApp <> IntPtr.Zero) Then
            'If (Me.hParentPrev = IntPtr.Zero) Then
            'hParentPrev = SetParent(hWordApp, me.Handle)
            Me.ThreadSetParent(Me.hWordApp)
            'End If
        End If
    End Sub

    Private Function CreateObjWord() As Boolean
        Try
            Me.oWordApp = CreateObject("Word.Application")
            Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture
            version = Double.Parse(Me.oWordApp.Version)
            Return True
        Catch ex As Exception

        End Try
        Return False
    End Function

    Private Sub LoadDocumentThread()

        Me.mutex.WaitOne()

        If (Me.oWordApp Is Nothing) Then
            If Not Me.CreateObjWord() Then
                MessageBox.Show("Erro ao abrir o documento. É necessario instalar o Microsoft Word", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End If

        If (Me.hWordApp = IntPtr.Zero) Then
            Me.SetWindowWord(Path.GetFileName(fileName))
        End If

        If IsWindow(Me.hWordApp) Then

            If (Me.oWordApp.Documents.Count > 0) Then
                Me.oWordApp.Documents.Close(Nothing, Nothing, Nothing)
            End If

            If Not File.Exists(fileName) Then
                Me.SetCommandBarControls(False)
            Else
                'me.oWordApp.Documents.Add(fileName, False, 0, True)
                Me.oWordApp.Documents.Open(fileName)
                Me.SetViewControls(False, ViewMode.WebView)
                Me.SetCommandBarControls(False)
            End If

            Me.oWordApp.Visible = True
            'SetWindowPos(hWordApp, Me.Handle, 0, 0, Me.Bounds.Width, Me.Bounds.Height, _
            '    SWP_NOZORDER Or SWP_NOMOVE Or SWP_DRAWFRAME Or SWP_NOSIZE)
            Me.ThreadSetWindowPos(hWordApp)
            Me.OnResizeControl()
        End If

        Me.mutex.Release()
    End Sub


    Private Sub WordViewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
