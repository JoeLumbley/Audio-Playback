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

' Level music by Joseph Lumbley Jr.

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

        Audio = New AudioPlayer()


        CreateSoundFiles()


        LoadAndRegisterSounds()


        Audio.LoopSound("Music")

        Debug.Print($"Running... {Now}")

    End Sub

    Private Sub LoadAndRegisterSounds()

        Audio.AddSound("Music", Path.Combine(Application.StartupPath, "level.mp3"))
        Audio.SetVolume("Music", 100)

        Audio.AddOverlapping("CashCollected", Path.Combine(Application.StartupPath, "CashCollected.mp3"))
        Audio.SetVolumeOverlapping("CashCollected", 300)

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Audio.PlayOverlapping("CashCollected")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If Audio.IsPlaying("Music") = True Then

            playLoop = False

            Audio.PauseSound("Music")

            Button2.Text = "Play Loop"

        Else

            playLoop = True

            Audio.LoopSound("Music")

            Button2.Text = "Pause Loop"

        End If

    End Sub

    Private Sub Form1_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Audio.CloseAll()

    End Sub

    Private Sub CreateSoundFiles()

        CreateFileFromResource(Path.Combine(Application.StartupPath, "level.mp3"), My.Resources.level)

        CreateFileFromResource(Path.Combine(Application.StartupPath, "CashCollected.mp3"), My.Resources.CashCollected)

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

        If Audio.IsPlaying("Music") Then
            Audio.FadeOutAndStop("Music", 2000)
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
                                   Audio.LoopSound("Music")
                               End If

                           End Sub

        t.Start()

    End Sub



End Class


' Windows Multimedia

' Windows Multimedia refers to the collection of technologies and APIs (Application Programming Interfaces)
' provided by Microsoft Windows for handling multimedia tasks on the Windows operating system.

' It includes components for playing audio and video, recording sound, working with MIDI devices, managing
' multimedia resources, and controlling multimedia hardware.

' Windows Multimedia APIs like DirectShow, DirectX, Media Control Interface, and others enable developers
' to create multimedia applications that can interact with various multimedia devices and perform tasks
' related to multimedia playback, recording, and processing.

' https://learn.microsoft.com/en-us/windows/win32/multimedia/windows-multimedia-start-page


' Media Control Interface

' The Media Control Interface (MCI) is a high-level programming interface provided by Microsoft Windows
' for controlling multimedia devices such as CD-ROM drives, audio and video devices, and other multimedia
' hardware.

' MCI provides a standard way for applications to interact with multimedia devices without needing to know
' the specific details of each device's hardware or communication protocols.

' By using MCI commands and functions, applications can play, record, pause, stop, and otherwise control
' multimedia playback and recording devices in a consistent and platform-independent manner.

' https://learn.microsoft.com/en-us/windows/win32/multimedia/mci


' mciSendStringW Function

' mciSendStringW is a function that is used to send a command string to an MCI device.

' The "W" at the end of the function name indicates that it is the wide-character version of the function,
' which means it accepts Unicode strings.

' This function allows applications to control multimedia devices and perform operations such as playing
' audio or video, recording sound, and managing multimedia resources by sending commands in the form of
' strings to MCI devices.

' https://learn.microsoft.com/en-us/previous-versions//dd757161(v=vs.85)


' open Command

' The "open" command is used in the Windows Multimedia API to open or initialize an MCI device for playback,
' recording or other multimedia operations.

' By sending an MCI command string with the "open" command using mciSendStringW, applications can specify
' the type of multimedia device to open (such as a CD-ROM drive, sound card, or video device), the file or
' resource to be accessed and any additional parameters required for the operation.

' This command is essential for preparing a multimedia device for use before performing playback, recording,
' or other actions on it.

' https://learn.microsoft.com/en-us/windows/win32/multimedia/open


' setaudio Command

' The "setaudio" command is used to set the audio parameters for a multimedia device.

' When sending an MCI command string with the "setaudio" command using the mciSendStringW function,
' applications can adjust settings such as volume, balance, speed, and other audio-related properties of the
' specified multimedia device.

' This command allows developers to control and customize the audio playback characteristics of the device
' to meet specific requirements or user preferences.

' https://learn.microsoft.com/en-us/windows/win32/multimedia/setaudio


' seek Command

' The "seek" command is used to move the current position of playback or recording to a specified location
' within a multimedia resource.

' When sending an MCI command string with the "seek" command using the mciSendStringW function,
' applications can specify the position or time where playback should start or resume within the multimedia
' content.

' This command allows developers to navigate to a specific point in audio or video playback, facilitating
' precise control over multimedia playback operations.

' https://learn.microsoft.com/en-us/windows/win32/multimedia/seek


' play Command

' The "play" command is used to start or resume playback of a multimedia resource.

' When sending an MCI command string with the "play" command using the mciSendStringW function, applications
' can instruct the multimedia device to begin playing the specified audio or video content from the current
' position.

' This command is essential for initiating playback of multimedia files, allowing developers to control the
' start and continuation of audio or video playback operations using MCI commands.

' https://learn.microsoft.com/en-us/windows/win32/multimedia/play


' status Command

' The "status" command is used to retrieve information about the current status of a multimedia device or
' resource.

' When sending an MCI command string with the "status" command using the mciSendStringW function,
' applications can query various properties and states of the specified multimedia device, such as playback
' position, volume level, mode (playing, paused, stopped), and other relevant information.

' This command allows developers to monitor and obtain real-time feedback on the status of multimedia
' playback or recording operations, enabling them to make informed decisions based on the device's current
' state.

' https://learn.microsoft.com/en-us/windows/win32/multimedia/status


' close Command

' The "close" command is used to close or release a multimedia device that was previously opened for
' playback, recording, or other operations.

' When sending an MCI command string with the "close" command using the mciSendStringW function,
' applications can instruct the multimedia device to release any resources associated with the device and
' prepare it for shutdown.

' This command is essential for properly closing and cleaning up after using a multimedia device, ensuring
' that resources are properly released and the device is no longer in use by the application.

' https://learn.microsoft.com/en-us/windows/win32/multimedia/close


' pause Command

' The pause command is used to temporarily halt the playback of media content, allowing the user to resume
' playback from the paused position at a later time.

' https://learn.microsoft.com/en-us/windows/win32/multimedia/pause

' Copilot is our AI assistant.


' I also make coding videos on my YouTube channel.
' https://www.youtube.com/@codewithjoe6074


