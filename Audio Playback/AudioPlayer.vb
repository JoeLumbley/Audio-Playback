Imports System.Runtime.InteropServices
Imports System.Text

Public Module AudioPlayer

    ' --- MCI API -------------------------------------------------------------

    <DllImport("winmm.dll", EntryPoint:="mciSendStringW")>
    Private Function mciSendStringW(
        <MarshalAs(UnmanagedType.LPWStr)> command As String,
        <MarshalAs(UnmanagedType.LPWStr)> returnString As StringBuilder,
        returnLength As UInteger,
        callback As IntPtr) As Integer
    End Function

    ' --- Internal State ------------------------------------------------------

    Private ReadOnly Aliases As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Private ReadOnly OverlapSuffixes As String() =
        Enumerable.Range(0, 24).Select(Function(i) Chr(Asc("A") + i).ToString()).ToArray()

    ' --- Helpers -------------------------------------------------------------

    Private Function Normalize(name As String) As String
        Return name.Trim().Replace(" ", "_")
    End Function

    Private Function Send(command As String) As Boolean
        Dim sb As New StringBuilder(256)
        Dim result = mciSendStringW(command, sb, CUInt(sb.Capacity), IntPtr.Zero)

        If result <> 0 Then
            Debug.Print($"MCI Error {result}: {command}")
            Return False
        End If

        Return True
    End Function

    Private Function Query(command As String) As String
        Dim sb As New StringBuilder(256)
        mciSendStringW(command, sb, CUInt(sb.Capacity), IntPtr.Zero)
        Return sb.ToString().Trim()
    End Function

    ' --- Core API ------------------------------------------------------------

    Public Function AddSound(soundName As String, filePath As String) As Boolean
        If String.IsNullOrWhiteSpace(soundName) OrElse Not IO.File.Exists(filePath) Then
            Debug.Print($"{soundName} not added.")
            Return False
        End If

        soundName = Normalize(soundName)

        If Aliases.Contains(soundName) Then Return True

        Dim cmd = $"open ""{filePath}"" alias {soundName}"

        If Send(cmd) Then
            Aliases.Add(soundName)
            Return True
        End If

        Debug.Print($"{soundName} failed to open.")
        Return False
    End Function

    Public Function PlaySound(soundName As String) As Boolean
        soundName = Normalize(soundName)

        If Not Aliases.Contains(soundName) Then
            Debug.Print($"{soundName} not playing.")
            Return False
        End If

        Return Send($"seek {soundName} to start") AndAlso
               Send($"play {soundName} notify")
    End Function

    Public Function LoopSound(soundName As String) As Boolean
        soundName = Normalize(soundName)

        If Not Aliases.Contains(soundName) Then
            Debug.Print($"{soundName} not looping.")
            Return False
        End If

        Return Send($"seek {soundName} to start") AndAlso
               Send($"play {soundName} repeat")
    End Function

    Public Function PauseSound(soundName As String) As Boolean
        soundName = Normalize(soundName)

        If Not Aliases.Contains(soundName) Then
            Debug.Print($"{soundName} not paused.")
            Return False
        End If

        Return Send($"pause {soundName} notify")
    End Function

    Public Function SetVolume(soundName As String, level As Integer) As Boolean
        soundName = Normalize(soundName)

        If Not Aliases.Contains(soundName) Then
            Debug.Print($"{soundName} volume not set.")
            Return False
        End If

        level = Math.Max(0, Math.Min(1000, level))

        Return Send($"setaudio {soundName} volume to {level}")
    End Function

    Public Function IsPlaying(soundName As String) As Boolean
        soundName = Normalize(soundName)
        If Not Aliases.Contains(soundName) Then Return False
        Return Query($"status {soundName} mode").Equals("playing", StringComparison.OrdinalIgnoreCase)
    End Function

    ' --- Overlapping Playback ------------------------------------------------

    Public Sub AddOverlapping(baseName As String, filePath As String)
        For Each suffix In OverlapSuffixes
            AddSound(baseName & suffix, filePath)
        Next
    End Sub

    Public Sub PlayOverlapping(baseName As String)
        For Each suffix In OverlapSuffixes
            Dim aliasName = Normalize(baseName & suffix)
            If Not IsPlaying(aliasName) Then
                PlaySound(aliasName)
                Exit Sub
            End If
        Next
    End Sub

    Public Sub SetVolumeOverlapping(baseName As String, level As Integer)
        For Each suffix In OverlapSuffixes
            SetVolume(baseName & suffix, level)
        Next
    End Sub

    ' --- Cleanup -------------------------------------------------------------

    Public Sub CloseAll()
        For Each aliasName In Aliases
            Send($"close {aliasName}")
        Next

        Aliases.Clear()
    End Sub

End Module
