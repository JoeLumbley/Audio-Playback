' Audio Playback

' Uses Windows Multimedia API for playback of multiple audio files simultaneously.

' MIT License
' Copyright(c) 2022 Joseph W. Lumbley

' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO AND THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE A NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.


Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Public Class Form1

    Private Audio As AudioPlayer

    Private WithEvents AudioRestartTimer As New Timer With {
        .Interval = 180000,
        .Enabled = True
    }

    Private loopShouldPlay As Boolean = True
    Private loopVolume As Integer = 100

    ' Windows 11 dark mode title bar support
    Private Const DWMWA_USE_IMMERSIVE_DARK_MODE As Integer = 20
    <DllImport("dwmapi.dll")>
    Private Shared Function DwmSetWindowAttribute(
        hWnd As IntPtr,
        attr As Integer,
        ByRef attrValue As Integer,
        attrSize As Integer
    ) As Integer
    End Function

    ' ============================================================
    ' Form Load / Init
    ' ============================================================
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CenterToScreen()
        Text = "Audio Playback - Code with Joe"

        If IsDarkMode() Then
            ApplyDarkTheme()
            ApplyDarkTitleBar(True)
        End If

        CreateSoundFiles()

        Audio = New AudioPlayer()
        LoadAndRegisterSounds()

        Debug.Print($"Running... {Now}")
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        PlayLoop(800)
    End Sub

    ' ============================================================
    ' Buttons
    ' ============================================================
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PlayOverlappingSound()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If loopShouldPlay Then
            loopShouldPlay = False
            FadeOutAndStopLoop(800)
            Button2.Text = "Play Loop"
        Else
            loopShouldPlay = True
            PlayLoop(800)
            Button2.Text = "Stop Loop"
        End If

    End Sub

    ' ============================================================
    ' Form Closing
    ' ============================================================
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        AudioRestartTimer?.Stop()
        AudioRestartTimer?.Dispose()

        Audio?.Dispose()
        Audio = Nothing

    End Sub

    ' ============================================================
    ' Audio Engine Restart
    ' ============================================================
    Private Sub AudioRestartTimer_Tick(sender As Object, e As EventArgs) Handles AudioRestartTimer.Tick
        RestartAudioEngine()
    End Sub

    Private Sub RestartAudioEngine()

        If Audio?.IsPlaying("loop") Then
            Audio?.FadeOutAndStop("loop", 800)
        End If

        Dim t As New Timer() With {.Interval = 900}

        AddHandler t.Tick,
            Sub()
                t.Stop()
                t.Dispose()

                Audio?.Dispose()
                Audio = Nothing

                Audio = New AudioPlayer()
                LoadAndRegisterSounds()

                If loopShouldPlay Then
                    PlayLoop(800)
                End If
            End Sub

        t.Start()
    End Sub

    ' ============================================================
    ' Sound File Creation
    ' ============================================================
    Private Sub CreateSoundFiles()

        CreateFileFromResource(Path.Combine(Application.StartupPath, "loop.mp3"),
                               My.Resources.Resource1.pause)

        CreateFileFromResource(Path.Combine(Application.StartupPath, "overlapping.mp3"),
                               My.Resources.Resource1.cashcollected)

    End Sub

    Private Sub CreateFileFromResource(filepath As String, resource As Byte())

        Try
            If Not File.Exists(filepath) Then
                File.WriteAllBytes(filepath, resource)
            End If
        Catch ex As Exception
            Debug.Print($"Error creating file: {ex.Message}")
        End Try

    End Sub

    ' ============================================================
    ' Audio Registration
    ' ============================================================
    Private Sub LoadAndRegisterSounds()

        Audio?.AddSound("loop", Path.Combine(Application.StartupPath, "loop.mp3"))
        Audio?.SetVolume("loop", loopVolume)

        Audio?.AddOverlapping("overlapping", Path.Combine(Application.StartupPath, "overlapping.mp3"))
        Audio?.SetVolumeOverlapping("overlapping", 200)

    End Sub

    ' ============================================================
    ' Playback Helpers
    ' ============================================================
    Private Sub PlayOverlappingSound()
        Audio?.PlayOverlapping("overlapping")
    End Sub

    Private Sub PlayLoop(durationMs As Integer)

        Audio?.SetVolume("loop", 0)
        Audio?.LoopSound("loop")
        Audio?.FadeVolume("loop", 0, loopVolume, durationMs)

    End Sub

    Private Sub FadeOutAndStopLoop(durationMs As Integer)
        If Audio?.IsPlaying("loop") Then
            Audio?.FadeOutAndStop("loop", durationMs)
        End If
    End Sub

    ' ============================================================
    ' Dark Mode Detection
    ' ============================================================
    Private Function IsDarkMode() As Boolean

        ' Windows 11+ uses Application.SystemColorMode
        If Environment.OSVersion.Version.Build >= 22000 Then
            Return Application.SystemColorMode = SystemColorMode.Dark
        End If

        Return IsSystemDarkMode_Win10()

    End Function

    Private Function IsSystemDarkMode_Win10() As Boolean
        Try
            Using key As RegistryKey =
                Registry.CurrentUser.OpenSubKey(
                    "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")

                If key Is Nothing Then Return False

                Dim value As Object = key.GetValue("AppsUseLightTheme", 1)
                Return CInt(value) = 0
            End Using
        Catch
            Return False
        End Try
    End Function

    ' ============================================================
    ' Dark Mode UI
    ' ============================================================
    Private Sub ApplyDarkTheme()

        Me.BackColor = Color.FromArgb(32, 32, 32)
        Me.ForeColor = Color.White

        For Each ctrl As Control In Me.Controls
            ctrl.ForeColor = Color.White

            If TypeOf ctrl Is Button Then
                ctrl.BackColor = Color.FromArgb(55, 55, 55)
            End If
        Next

    End Sub

    Private Sub ApplyDarkTitleBar(isDark As Boolean)

        ' Only apply dark title bar on Windows 11+
        If Environment.OSVersion.Version.Build < 22000 Then Exit Sub

        Dim value As Integer = If(isDark, 1, 0)

        DwmSetWindowAttribute(
            Me.Handle,
            DWMWA_USE_IMMERSIVE_DARK_MODE,
            value,
            Marshal.SizeOf(value))

    End Sub

End Class




' Copilot is our AI assistant.


' I also make coding videos on my YouTube channel: Code with Joe.
' https://www.youtube.com/@codewithjoe6074
