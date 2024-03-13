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

'Monica is our an AI assistant.
'https://monica.im/

'I'm making a video to explain the code on my YouTube channel.
'https://www.youtube.com/@codewithjoe6074

Imports System.Runtime.InteropServices
Imports System.Text
Imports System.IO

Public Class Form1

    Private Enum MCI_NOTIFY As Integer
        SUCCESSFUL = &H1
        SUPERSEDED = &H2
        ABORTED = &H4
        FAILURE = &H8
    End Enum

    <DllImport("winmm.dll", EntryPoint:="mciSendStringW")>
    Private Shared Function mciSendStringW(<MarshalAs(UnmanagedType.LPTStr)> ByVal lpszCommand As String,
                                           <MarshalAs(UnmanagedType.LPWStr)> ByVal lpszReturnString As StringBuilder,
                                           ByVal cchReturn As UInteger, ByVal hwndCallback As IntPtr) As Integer
    End Function


    Private Sounds() As String

    Private AppPath As String

    Private FilePath As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        AppPath = Application.StartupPath

        CreateSoundFileFromResource()

        FilePath = Path.Combine(AppPath, "level.mp3")

        AddSound("Music", FilePath)

        SetVolume("Music", 50)

        FilePath = Path.Combine(AppPath, "CashCollected.mp3")

        AddOverlaping("CashCollected", FilePath)

        SetVolumeOverlaping("CashCollected", 500)

        LoopSound("Music")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        PlayOverlaping("CashCollected")

    End Sub

    Private Sub Form1_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        CloseSounds()

    End Sub

    Private Function AddSound(ByVal SoundName As String, ByVal FilePath As String) As Boolean

        Dim CommandOpen As String = "open " & Chr(34) & FilePath & Chr(34) & " alias " & SoundName

        'Do we have a name and does the file exist?
        If Not SoundName.Trim = String.Empty And IO.File.Exists(FilePath) Then
            'Yes, we have a name and the file exists.

            'Do we have sounds?
            If Sounds IsNot Nothing Then
                'Yes, we have sounds.

                'Is the sound in the array already?
                If Not Sounds.Contains(SoundName) Then
                    'No, the sound is not in the array.

                    'Did the sound file open?
                    If mciSendStringW(CommandOpen, Nothing, 0, IntPtr.Zero) = 0 Then
                        'Yes, the sound file did open.

                        'Add the sound to the Sounds array.
                        Array.Resize(Sounds, Sounds.Length + 1)
                        Sounds(Sounds.Length - 1) = SoundName

                        Return True

                    End If

                End If

            Else
                'No, we do not have sounds.

                'Did the sound file open?
                If mciSendStringW(CommandOpen, Nothing, 0, IntPtr.Zero) = 0 Then
                    'Yes, the sound file did open.

                    'Start the Sounds array with the sound.
                    ReDim Sounds(0)
                    Sounds(0) = SoundName

                    Return True

                End If

            End If

        End If

        Return False

    End Function

    Private Function SetVolume(ByVal SoundName As String, ByVal Level As Integer) As Boolean

        Dim CommandVolume As String = "setaudio " & SoundName & " volume to " & Level.ToString

        'Do we have sounds?
        If Sounds IsNot Nothing Then
            'Yes, we have sounds.

            'Is the sound in the sounds array?
            If Sounds.Contains(SoundName) Then
                'Yes, the sound is the sounds array.

                'Is the level in the valid range?
                If Level >= 0 And Level <= 1000 Then
                    'Yes, the level is in range.

                    'Was the volume set?
                    If mciSendStringW(CommandVolume, Nothing, 0, IntPtr.Zero) = 0 Then
                        'Yes, the volume was set.

                        Return True

                    End If

                End If

            End If

        End If

        Return False 'The volume was not set.

    End Function

    Private Function LoopSound(ByVal SoundName As String) As Boolean

        If Sounds IsNot Nothing Then

            If Not Sounds.Contains(SoundName) Then

                Return False

            End If

            mciSendStringW("seek " & SoundName & " to start", Nothing, 0, IntPtr.Zero)

            If mciSendStringW("play " & SoundName & " repeat", Nothing, 0, Me.Handle) <> 0 Then

                Return False

            End If

        End If

        Return True

    End Function

    Private Function PlaySound(ByVal SoundName As String) As Boolean

        Dim CommandFromStart As String = "seek " & SoundName & " to start"

        Dim CommandPlay As String = "play " & SoundName & " notify"

        If Sounds IsNot Nothing Then

            If Sounds.Contains(SoundName) Then

                'Play sound file from the start.
                mciSendStringW(CommandFromStart, Nothing, 0, IntPtr.Zero)

                If mciSendStringW(CommandPlay, Nothing, 0, Me.Handle) = 0 Then

                    Return True

                End If

            End If

        End If

        Return False

    End Function

    Private Function IsPlaying(ByVal SoundName As String) As Boolean

        Return (GetStatus(SoundName, "mode") = "playing")

    End Function

    Private Sub AddOverlaping(ByVal SoundName As String, ByVal FilePath As String)

        AddSound(SoundName & "A", FilePath)

        AddSound(SoundName & "B", FilePath)

        AddSound(SoundName & "C", FilePath)

        AddSound(SoundName & "D", FilePath)

    End Sub

    Private Sub SetVolumeOverlaping(ByVal SoundName As String, ByVal Level As Integer)

        SetVolume(SoundName & "A", Level)

        SetVolume(SoundName & "B", Level)

        SetVolume(SoundName & "C", Level)

        SetVolume(SoundName & "D", Level)

    End Sub

    Private Sub PlayOverlaping(ByVal SoundName As String)

        If IsPlaying(SoundName & "A") = False Then

            PlaySound(SoundName & "A")

        Else

            If IsPlaying(SoundName & "B") = False Then

                PlaySound(SoundName & "B")

            Else

                If IsPlaying(SoundName & "C") = False Then

                    PlaySound(SoundName & "C")

                Else

                    If IsPlaying(SoundName & "D") = False Then

                        PlaySound(SoundName & "D")

                    End If

                End If

            End If

        End If

    End Sub

    Private Function GetStatus(ByVal SoundName As String, ByVal StatusType As String) As String

        Dim CommandStatus As String = "status " & SoundName & " " & StatusType

        Dim StatusReturn As New System.Text.StringBuilder(128)

        If Sounds IsNot Nothing Then

            If Sounds.Contains(SoundName) Then

                mciSendStringW(CommandStatus, StatusReturn, 128, IntPtr.Zero)

                Return StatusReturn.ToString.Trim.ToLower

            End If

        End If

        Return String.Empty

    End Function

    Private Sub CloseSounds()

        If Sounds IsNot Nothing Then

            For Each Sound In Sounds

                mciSendStringW("close " & Sound, Nothing, 0, IntPtr.Zero)

            Next

        End If

        Sounds = Nothing

    End Sub

    Private Sub CreateSoundFileFromResource()

        FilePath = Path.Combine(AppPath, "level.mp3")

        If Not IO.File.Exists(FilePath) Then

            IO.File.WriteAllBytes(FilePath, My.Resources.level)

        End If

        FilePath = Path.Combine(AppPath, "CashCollected.mp3")

        If Not IO.File.Exists(FilePath) Then

            IO.File.WriteAllBytes(FilePath, My.Resources.CashCollected)

        End If

    End Sub

End Class

'Windows Multimedia
'https://learn.microsoft.com/en-us/windows/win32/multimedia/windows-multimedia-start-page
'Windows Multimedia refers to the collection of technologies and APIs (Application Programming Interfaces)
'provided by Microsoft Windows for handling multimedia tasks on the Windows operating system. It includes
'components for playing audio and video, recording sound, working with MIDI
'(Musical Instrument Digital Interface) devices, managing multimedia resources, and controlling multimedia
'hardware. Windows Multimedia APIs like DirectShow, DirectX, MCI (Media Control Interface), and others
'enable developers to create multimedia applications that can interact with various multimedia devices and
'perform tasks related to multimedia playback, recording, and processing.

'MCI
'https://learn.microsoft.com/en-us/windows/win32/multimedia/mci
'The Media Control Interface (MCI) is a high-level programming interface provided by Microsoft Windows
'for controlling multimedia devices such as CD-ROM drives, audio and video devices, and other multimedia
'hardware. MCI provides a standard way for applications to interact with multimedia devices without
'needing to know the specific details of each device's hardware or communication protocols.
'By using MCI commands and functions, applications can play, record, pause, stop, and otherwise control
'multimedia playback and recording devices in a consistent and platform-independent manner.

'mciSendStringW function
'https://learn.microsoft.com/en-us/previous-versions//dd757161(v=vs.85)
'mciSendStringW is a function in the Windows Multimedia API that is used to send a command string to an
'MCI (Media Control Interface) device. The "W" at the end of the function name indicates that it is the
'wide-character version of the function, which means it accepts Unicode strings.
'This function allows applications to control multimedia devices and perform operations such as playing
'audio or video, recording sound, and managing multimedia resources by sending commands in the form of
'strings to MCI devices.

'open command
'https://learn.microsoft.com/en-us/windows/win32/multimedia/open
'The mciSendStringW function with the "open" command is used in the Windows Multimedia API to open or
'initialize an MCI (Media Control Interface) device for playback, recording, or other multimedia
'operations. By sending an MCI command string with the "open" command using mciSendStringW, applications
'can specify the type of multimedia device to open (such as a CD-ROM drive, sound card, or video device),
'the file or resource to be accessed, and any additional parameters required for the operation.
'This command is essential for preparing a multimedia device for use before performing playback, recording,
'or other actions on it.

'setaudio command
'https://learn.microsoft.com/en-us/windows/win32/multimedia/setaudio

'seek command
'https://learn.microsoft.com/en-us/windows/win32/multimedia/seek

'play command
'https://learn.microsoft.com/en-us/windows/win32/multimedia/play

'status command
'https://learn.microsoft.com/en-us/windows/win32/multimedia/status

'close command
'https://learn.microsoft.com/en-us/windows/win32/multimedia/close