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
    'Private WithEvents AudioRestartTimer As New Timer With {
    '    .Interval = 15000,
    '    .Enabled = True
    '}


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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CenterToScreen()
        Text = "Audio Playback - Code with Joe"

        ' Apply dark mode if Windows is in dark mode
        Dim dark As Boolean = IsDarkMode()
        If dark Then
            ApplyDarkTheme()
            ApplyDarkTitleBar(dark)
        End If


        CreateSoundFiles()
        Audio = New AudioPlayer()
        LoadAndRegisterSounds()
        'Audio.LoopSound("loop")

        Debug.Print($"Running... {Now}")
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Start the audio engine when the form is shown

        PlayLoop(2000) ' Fade in over 2 seconds

        'Audio.SetVolume("loop", 0) ' Start with volume at 0 to fade in
        'Audio.LoopSound("loop") ' Start looping the sound
        'Audio.FadeVolume("loop", 0, loopVolume, 2000) ' Fade in over 2 seconds

    End Sub

    Private Sub PlayLoop(durationMs As Integer)

        ' Fade‑in loop
        Audio.SetVolume("loop", 0)
        Audio.LoopSound("loop")
        Audio.FadeVolume("loop", 0, loopVolume, durationMs)

    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Audio.PlayOverlapping("overlapping")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ' Toggle loop playback 
        If loopShouldPlay Then
            loopShouldPlay = False

            If Audio.IsPlaying("loop") = True Then Audio.FadeOutAndStop("loop", 2000)

            Button2.Text = "Play Loop"

        Else
            loopShouldPlay = True

            PlayLoop(2000)

            Button2.Text = "Pause Loop"

        End If


    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        Audio.CloseAll()

    End Sub



    Private Sub AudioRestartTimer_Tick(sender As Object, e As EventArgs) Handles AudioRestartTimer.Tick
        RestartAudioEngine()
    End Sub


    Private Sub CreateSoundFiles()

        CreateFileFromResource(Path.Combine(Application.StartupPath, "loop.mp3"), My.Resources.Resource1.pause)

        CreateFileFromResource(Path.Combine(Application.StartupPath, "overlapping.mp3"), My.Resources.Resource1.cashcollected)

    End Sub

    Private Sub LoadAndRegisterSounds()

        Audio.AddSound("loop", Path.Combine(Application.StartupPath, "loop.mp3"))
        Audio.SetVolume("loop", 100)

        Audio.AddOverlapping("overlapping", Path.Combine(Application.StartupPath, "overlapping.mp3"))
        Audio.SetVolumeOverlapping("overlapping", 200)

    End Sub



    Private Sub RestartAudioEngine()

        ' Fade-out and stop the loop if it's playing
        If Audio.IsPlaying("loop") Then
            Audio.FadeOutAndStop("loop", 2000)
        End If

        ' Wait for fade-out to complete before restarting engine
        Dim t As New Timer() With {.Interval = 2300}

        AddHandler t.Tick, Sub()
                               t.Stop()
                               t.Dispose()

                               ' Dispose old engine
                               Audio?.Dispose()

                               Audio = Nothing

                               ' Create new engine
                               Audio = New AudioPlayer()

                               ' Reload all sounds
                               LoadAndRegisterSounds()

                               ' Restore loop based on state
                               If loopShouldPlay Then
                                   Audio.LoopSound("loop")
                               End If

                           End Sub

        t.Start()

    End Sub


    Private Sub CreateFileFromResource(filepath As String, resource As Byte())

        Try

            If Not IO.File.Exists(filepath) Then

                IO.File.WriteAllBytes(filepath, resource)

            End If

        Catch ex As Exception

            Debug.Print($"Error creating file: {ex.Message}")

        End Try

    End Sub


    Private Function IsDarkMode() As Boolean
        If Environment.OSVersion.Version.Build >= 22000 Then
            ' Windows 11+
            Return System.Windows.Forms.Application.SystemColorMode = System.Windows.Forms.SystemColorMode.Dark
        Else
            ' Windows 10 fallback
            Return IsSystemDarkMode_Win10()
        End If
    End Function

    Private Function IsSystemDarkMode_Win10() As Boolean
        Try
            Dim key As RegistryKey =
            Registry.CurrentUser.OpenSubKey(
                "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")

            If key Is Nothing Then Return False

            Dim value As Object = key.GetValue("AppsUseLightTheme", 1)
            Return CInt(value) = 0
        Catch
            Return False
        End Try
    End Function


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
        If Environment.OSVersion.Version.Build < 22000 Then Exit Sub ' Only Windows 11+

        Dim value As Integer = If(isDark, 1, 0)
        DwmSetWindowAttribute(Me.Handle,
                              DWMWA_USE_IMMERSIVE_DARK_MODE,
                              value,
                              Marshal.SizeOf(value))

    End Sub



End Class

' Copilot is our AI assistant.


' I also make coding videos on my YouTube channel: Code with Joe.
' https://www.youtube.com/@codewithjoe6074


