Imports System.Runtime.InteropServices 'For DllImport of winmm.dll
Imports System.Text
Imports System.IO

Public Class Form1

    Private Enum MCI_NOTIFY As Integer
        SUCCESSFUL = &H1
        SUPERSEDED = &H2
        ABORTED = &H4
        FAILURE = &H8
    End Enum

    'Import Windows Multimedia API for playback of multiple audio files simultaneously.
    <DllImport("winmm.dll", EntryPoint:="mciSendStringW")>
    Private Shared Function mciSendStringW(<MarshalAs(UnmanagedType.LPTStr)> ByVal lpszCommand As String,
                                           <MarshalAs(UnmanagedType.LPWStr)> ByVal lpszReturnString As StringBuilder,
                                           ByVal cchReturn As UInteger, ByVal hwndCallback As IntPtr) As Integer
    End Function
    'mciSendStringW is a function in the Windows Multimedia API that is used to send a command string to an
    'MCI (Media Control Interface) device. The "W" at the end of the function name indicates that it is the
    'wide-character version of the function, which means it accepts Unicode strings.
    'This function allows applications to control multimedia devices and perform operations such as playing
    'audio or video, recording sound, and managing multimedia resources by sending commands in the form of
    'strings to MCI devices.

    'Create array for sounds.
    Private Sounds() As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CreateSoundFileFromResource()

        AddSound("Music", Application.StartupPath & "level.mp3")

        SetVolume("Music", 50)

        AddOverlaping("CashCollected", Application.StartupPath & "CashCollected.mp3")

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

    Private Function IsPlaying(ByVal SoundName As String) As Boolean

        Return (GetStatus(SoundName, "mode") = "playing")

    End Function

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

    Private Sub CloseSounds()

        If Sounds IsNot Nothing Then

            For Each Sound In Sounds

                mciSendStringW("close " & Sound, Nothing, 0, IntPtr.Zero)

            Next

        End If

        Sounds = Nothing

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

    Private Sub CreateSoundFileFromResource()

        Dim file As String = Path.Combine(Application.StartupPath, "level.mp3")

        If Not IO.File.Exists(file) Then

            IO.File.WriteAllBytes(file, My.Resources.level)

        End If

        file = Path.Combine(Application.StartupPath, "CashCollected.mp3")

        If Not IO.File.Exists(file) Then

            IO.File.WriteAllBytes(file, My.Resources.CashCollected)

        End If

    End Sub

End Class

'Windows Multimedia
'https://learn.microsoft.com/en-us/windows/win32/multimedia/windows-multimedia-start-page

'MCI
'https://learn.microsoft.com/en-us/windows/win32/multimedia/mci

'mciSendString function
'https://learn.microsoft.com/en-us/previous-versions//dd757161(v=vs.85)

'open command
'https://learn.microsoft.com/en-us/windows/win32/multimedia/open

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