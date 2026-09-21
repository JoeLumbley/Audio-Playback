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


    Private playLoop As Boolean = True


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CenterToScreen()

        Text = "Audio Playback - Code with Joe"

        CreateSoundFiles()

        Audio = New AudioPlayer()

        LoadAndRegisterSounds()

        Audio.LoopSound("loop")

        Debug.Print($"Running... {Now}")

    End Sub

    Private Sub LoadAndRegisterSounds()

        Audio.AddSound("loop", Path.Combine(Application.StartupPath, "loop.mp3"))
        Audio.SetVolume("loop", 100)

        Audio.AddOverlapping("overlapping", Path.Combine(Application.StartupPath, "overlapping.mp3"))
        Audio.SetVolumeOverlapping("overlapping", 200)

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Audio.PlayOverlapping("overlapping")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If Audio.IsPlaying("loop") = True Then

            playLoop = False

            Audio.PauseSound("loop")

            Button2.Text = "Play Loop"

        Else

            playLoop = True

            Audio.LoopSound("loop")

            Button2.Text = "Pause Loop"

        End If

    End Sub

    Private Sub Form1_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Audio.CloseAll()

    End Sub

    Private Sub CreateSoundFiles()

        CreateFileFromResource(Path.Combine(Application.StartupPath, "loop.mp3"), My.Resources.Resource1.pause)

        CreateFileFromResource(Path.Combine(Application.StartupPath, "overlapping.mp3"), My.Resources.Resource1.cashcollected)

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

    Private Sub AudioRestartTimer_Tick(sender As Object, e As EventArgs) Handles AudioRestartTimer.Tick
        RestartAudioEngine()
    End Sub

    Private Sub RestartAudioEngine()

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

                               ' Create new engine
                               Audio = New AudioPlayer()

                               ' Reload all sounds
                               LoadAndRegisterSounds()

                               ' Restore loop based on state
                               If playLoop Then
                                   Audio.LoopSound("loop")
                               End If

                           End Sub

        t.Start()

    End Sub



End Class

' Copilot is our AI assistant.


' I also make coding videos on my YouTube channel.
' https://www.youtube.com/@codewithjoe6074


