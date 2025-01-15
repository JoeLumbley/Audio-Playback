'Audio Playback

'Uses Windows Multimedia API for playback of multiple audio files simultaneously.

'MIT License
'Copyright(c) 2022 Joseph W. Lumbley

'Permission Is hereby granted, free Of charge, to any person obtaining a copy
'of this software And associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, And/Or sell
'copies of the Software, And to permit persons to whom the Software Is
'furnished to do so, subject to the following conditions:

'The above copyright notice And this permission notice shall be included In all
'copies Or substantial portions of the Software.

'THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY Of ANY KIND, EXPRESS Or
'IMPLIED, INCLUDING BUT Not LIMITED To THE WARRANTIES Of MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE And NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS Or COPYRIGHT HOLDERS BE LIABLE For ANY CLAIM, DAMAGES Or OTHER
'LIABILITY, WHETHER In AN ACTION Of CONTRACT, TORT Or OTHERWISE, ARISING FROM,
'OUT OF Or IN CONNECTION WITH THE SOFTWARE Or THE USE Or OTHER DEALINGS IN THE
'SOFTWARE.

'Level music by Joseph Lumbley Jr.

Imports System.Runtime.InteropServices
Imports System.Text
Imports System.IO


Public Class Form1

    <DllImport("winmm.dll", EntryPoint:="mciSendStringW")>
    Private Shared Function mciSendStringW(<MarshalAs(UnmanagedType.LPTStr)> ByVal lpszCommand As String,
                                           <MarshalAs(UnmanagedType.LPWStr)> ByVal lpszReturnString As StringBuilder,
                                           ByVal cchReturn As UInteger, ByVal hwndCallback As IntPtr) As Integer
    End Function


    Private Sounds() As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Text = "Audio Playback - Code with Joe"

        CreateSoundFileFromResource()

        Dim FilePath As String = Path.Combine(Application.StartupPath, "level.mp3")

        AddSound("Music", FilePath)

        SetVolume("Music", 600)

        FilePath = Path.Combine(Application.StartupPath, "CashCollected.mp3")

        AddOverlapping("CashCollected", FilePath)

        SetVolumeOverlapping("CashCollected", 900)

        LoopSound("Music")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        PlayOverlapping("CashCollected")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If IsPlaying("Music") = True Then

            PauseSound("Music")

            Button2.Text = "Play Loop"

        Else

            LoopSound("Music")

            Button2.Text = "Pause Loop"

        End If

    End Sub

    Private Sub Form1_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        CloseSounds()

    End Sub

#Region "Sound Management"

    Private Function AddSound(SoundName As String, FilePath As String) As Boolean

        'Do we have a name and does the file exist?
        If Not String.IsNullOrWhiteSpace(SoundName) AndAlso IO.File.Exists(FilePath) Then
            'Yes, we have a name and the file exists.

            Dim CommandOpen As String = $"open ""{FilePath}"" alias {SoundName}"

            Dim ReturnString As New StringBuilder(128)

            'Do we have sounds?
            If Sounds IsNot Nothing Then
                'Yes, we have sounds.

                'Is the sound in the array already?
                If Not Sounds.Contains(SoundName) Then
                    'No, the sound is not in the array.

                    'Did the sound file open?
                    If mciSendStringW(CommandOpen, ReturnString, 0, IntPtr.Zero) = 0 Then
                        'Yes, the sound file did open.

                        'Add the sound to the Sounds array.
                        Array.Resize(Sounds, Sounds.Length + 1)

                        Sounds(Sounds.Length - 1) = SoundName

                        Return True 'The sound was added.

                    End If

                End If

            Else
                'No, we do not have sounds.

                'Did the sound file open?
                If mciSendStringW(CommandOpen, ReturnString, 0, IntPtr.Zero) = 0 Then
                    'Yes, the sound file did open.

                    'Start the Sounds array with the sound.
                    ReDim Sounds(0)

                    Sounds(0) = SoundName

                    Return True 'The sound was added.

                End If

            End If

        End If

        Return False 'The sound was not added.

    End Function

    Private Function SetVolume(SoundName As String, Level As Integer) As Boolean

        'Do we have sounds?
        If Sounds IsNot Nothing Then
            'Yes, we have sounds.

            'Is the sound in the sounds array?
            If Sounds.Contains(SoundName) Then
                'Yes, the sound is the sounds array.

                'Is the level in the valid range?
                If Level >= 0 AndAlso Level <= 1000 Then
                    'Yes, the level is in range.

                    Dim CommandVolume As String = $"setaudio {SoundName} volume to {Level}"

                    Dim ReturnString As New StringBuilder(128)

                    'Was the volume set?
                    If mciSendStringW(CommandVolume, ReturnString, 0, IntPtr.Zero) = 0 Then

                        Return True 'The volume was set.

                    End If

                End If

            End If

        End If

        Return False 'The volume was not set.

    End Function

    Private Function LoopSound(SoundName As String) As Boolean

        ' Do we have sounds?
        If Sounds IsNot Nothing Then
            ' Yes, we have sounds.

            ' Is the sound in the array?
            If Not Sounds.Contains(SoundName) Then
                ' No, the sound is not in the array.

                Return False ' The sound is not playing.

            End If

            Dim CommandSeekToStart As String = $"seek {SoundName} to start"

            Dim ReturnString As New StringBuilder(128)

            mciSendStringW(CommandSeekToStart, ReturnString, 0, IntPtr.Zero)

            Dim CommandPlayRepete As String = $"play {SoundName} repeat"

            If mciSendStringW(CommandPlayRepete, ReturnString, 0, Me.Handle) <> 0 Then

                Return False ' The sound is not playing.

            End If

        End If

        Return True ' The sound is playing.

    End Function

    Private Function PlaySound(SoundName As String) As Boolean

        'Do we have sounds?
        If Sounds IsNot Nothing Then
            'Yes, we have sounds.

            'Is the sound in the array?
            If Sounds.Contains(SoundName) Then
                'Yes, the sound is in the array.

                Dim CommandSeekToStart As String = $"seek {SoundName} to start"

                Dim ReturnString As New StringBuilder(128)

                mciSendStringW(CommandSeekToStart, ReturnString, 0, IntPtr.Zero)

                Dim CommandPlay As String = $"play {SoundName} notify"

                If mciSendStringW(CommandPlay, ReturnString, 0, Me.Handle) = 0 Then

                    Return True 'The sound is playing.

                End If

            End If

        End If

        Return False 'The sound is not playing.

    End Function

    Private Function PauseSound(SoundName As String) As Boolean

        'Do we have sounds?
        If Sounds IsNot Nothing Then
            'Yes, we have sounds.

            'Is the sound in the array?
            If Sounds.Contains(SoundName) Then
                'Yes, the sound is in the array.

                Dim CommandPause As String = $"pause {SoundName} notify"

                Dim ReturnString As New StringBuilder(128)

                If mciSendStringW(CommandPause, ReturnString, 0, Me.Handle) = 0 Then

                    Return True 'The sound is paused.

                End If

            End If

        End If

        Return False 'The sound is not paused.

    End Function

    Private Function IsPlaying(SoundName As String) As Boolean

        Return GetStatus(SoundName, "mode") = "playing"

    End Function

    Private Sub AddOverlapping(SoundName As String, FilePath As String)

        For Each Suffix As String In {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"}

            AddSound(SoundName & Suffix, FilePath)

        Next

    End Sub

    Private Sub PlayOverlapping(SoundName As String)

        For Each Suffix As String In {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"}

            If Not IsPlaying(SoundName & Suffix) Then

                PlaySound(SoundName & Suffix)

                Exit Sub

            End If

        Next

    End Sub

    Private Sub SetVolumeOverlapping(SoundName As String, Level As Integer)

        For Each Suffix As String In {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"}

            SetVolume(SoundName & Suffix, Level)

        Next

    End Sub

    Private Function GetStatus(SoundName As String, StatusType As String) As String

        If Sounds IsNot Nothing Then

            If Sounds.Contains(SoundName) Then

                Dim CommandStatus As String = $"status {SoundName} {StatusType}"

                Dim StatusReturn As New StringBuilder(128)

                mciSendStringW(CommandStatus, StatusReturn, 128, IntPtr.Zero)

                Return StatusReturn.ToString.Trim.ToLower

            End If

        End If

        Return String.Empty

    End Function

    Private Sub CloseSounds()

        If Sounds IsNot Nothing Then

            For Each Sound In Sounds

                Dim CommandClose As String = $"close {Sound}"

                Dim ReturnString As New StringBuilder(128)

                mciSendStringW(CommandClose, ReturnString, 0, IntPtr.Zero)

            Next

        End If

    End Sub

    Private Sub CreateSoundFileFromResource()

        Dim FilePath As String = Path.Combine(Application.StartupPath, "level.mp3")

        If Not IO.File.Exists(FilePath) Then

            IO.File.WriteAllBytes(FilePath, My.Resources.level)

        End If

        FilePath = Path.Combine(Application.StartupPath, "CashCollected.mp3")

        If Not IO.File.Exists(FilePath) Then

            IO.File.WriteAllBytes(FilePath, My.Resources.CashCollected)

        End If

    End Sub

#End Region

End Class


'Windows Multimedia

'Windows Multimedia refers to the collection of technologies and APIs (Application Programming Interfaces)
'provided by Microsoft Windows for handling multimedia tasks on the Windows operating system.

'It includes components for playing audio and video, recording sound, working with MIDI devices, managing
'multimedia resources, and controlling multimedia hardware.

'Windows Multimedia APIs like DirectShow, DirectX, Media Control Interface, and others enable developers
'to create multimedia applications that can interact with various multimedia devices and perform tasks
'related to multimedia playback, recording, and processing.

'https://learn.microsoft.com/en-us/windows/win32/multimedia/windows-multimedia-start-page


'Media Control Interface

'The Media Control Interface (MCI) is a high-level programming interface provided by Microsoft Windows
'for controlling multimedia devices such as CD-ROM drives, audio and video devices, and other multimedia
'hardware.

'MCI provides a standard way for applications to interact with multimedia devices without needing to know
'the specific details of each device's hardware or communication protocols.

'By using MCI commands and functions, applications can play, record, pause, stop, and otherwise control
'multimedia playback and recording devices in a consistent and platform-independent manner.

'https://learn.microsoft.com/en-us/windows/win32/multimedia/mci


'mciSendStringW Function

'mciSendStringW is a function that is used to send a command string to an MCI device.

'The "W" at the end of the function name indicates that it is the wide-character version of the function,
'which means it accepts Unicode strings.

'This function allows applications to control multimedia devices and perform operations such as playing
'audio or video, recording sound, and managing multimedia resources by sending commands in the form of
'strings to MCI devices.

'https://learn.microsoft.com/en-us/previous-versions//dd757161(v=vs.85)


'open Command

'The "open" command is used in the Windows Multimedia API to open or initialize an MCI device for playback,
'recording or other multimedia operations.

'By sending an MCI command string with the "open" command using mciSendStringW, applications can specify
'the type of multimedia device to open (such as a CD-ROM drive, sound card, or video device), the file or
'resource to be accessed and any additional parameters required for the operation.

'This command is essential for preparing a multimedia device for use before performing playback, recording,
'or other actions on it.

'https://learn.microsoft.com/en-us/windows/win32/multimedia/open


'setaudio Command

'The "setaudio" command is used to set the audio parameters for a multimedia device.

'When sending an MCI command string with the "setaudio" command using the mciSendStringW function,
'applications can adjust settings such as volume, balance, speed, and other audio-related properties of the
'specified multimedia device.

'This command allows developers to control and customize the audio playback characteristics of the device
'to meet specific requirements or user preferences.

'https://learn.microsoft.com/en-us/windows/win32/multimedia/setaudio


'seek Command

'The "seek" command is used to move the current position of playback or recording to a specified location
'within a multimedia resource.

'When sending an MCI command string with the "seek" command using the mciSendStringW function,
'applications can specify the position or time where playback should start or resume within the multimedia
'content.

'This command allows developers to navigate to a specific point in audio or video playback, facilitating
'precise control over multimedia playback operations.

'https://learn.microsoft.com/en-us/windows/win32/multimedia/seek


'play Command

'The "play" command is used to start or resume playback of a multimedia resource.

'When sending an MCI command string with the "play" command using the mciSendStringW function, applications
'can instruct the multimedia device to begin playing the specified audio or video content from the current
'position.

'This command is essential for initiating playback of multimedia files, allowing developers to control the
'start and continuation of audio or video playback operations using MCI commands.

'https://learn.microsoft.com/en-us/windows/win32/multimedia/play


'status Command

'The "status" command is used to retrieve information about the current status of a multimedia device or
'resource.

'When sending an MCI command string with the "status" command using the mciSendStringW function,
'applications can query various properties and states of the specified multimedia device, such as playback
'position, volume level, mode (playing, paused, stopped), and other relevant information.

'This command allows developers to monitor and obtain real-time feedback on the status of multimedia
'playback or recording operations, enabling them to make informed decisions based on the device's current
'state.

'https://learn.microsoft.com/en-us/windows/win32/multimedia/status


'close Command

'The "close" command is used to close or release a multimedia device that was previously opened for
'playback, recording, or other operations.

'When sending an MCI command string with the "close" command using the mciSendStringW function,
'applications can instruct the multimedia device to release any resources associated with the device and
'prepare it for shutdown.

'This command is essential for properly closing and cleaning up after using a multimedia device, ensuring
'that resources are properly released and the device is no longer in use by the application.

'https://learn.microsoft.com/en-us/windows/win32/multimedia/close


'pause Command

'The pause command is used to temporarily halt the playback of media content, allowing the user to resume
'playback from the paused position at a later time.

'https://learn.microsoft.com/en-us/windows/win32/multimedia/pause



'Monica is our an AI assistant.
'https://monica.im/


'I also make coding videos on my YouTube channel.
'https://www.youtube.com/@codewithjoe6074
