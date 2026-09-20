'Imports System.Runtime.InteropServices
'Imports System.Text

'Public Module AudioPlayer

'    ' --- MCI API -------------------------------------------------------------

'    <DllImport("winmm.dll", EntryPoint:="mciSendStringW")>
'    Private Function mciSendStringW(
'        <MarshalAs(UnmanagedType.LPWStr)> command As String,
'        <MarshalAs(UnmanagedType.LPWStr)> returnString As StringBuilder,
'        returnLength As UInteger,
'        callback As IntPtr) As Integer
'    End Function

'    ' --- Internal State ------------------------------------------------------

'    Private ReadOnly Aliases As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

'    Private ReadOnly OverlapSuffixes As String() =
'        Enumerable.Range(0, 24).Select(Function(i) Chr(Asc("A") + i).ToString()).ToArray()

'    ' --- Helpers -------------------------------------------------------------

'    Private Function Normalize(name As String) As String
'        Return name.Trim().Replace(" ", "_")
'    End Function

'    Private Function Send(command As String) As Boolean
'        Dim sb As New StringBuilder(256)
'        Dim result = mciSendStringW(command, sb, CUInt(sb.Capacity), IntPtr.Zero)

'        If result <> 0 Then
'            Debug.Print($"MCI Error {result}: {command}")
'            Return False
'        End If

'        Return True
'    End Function

'    Private Function Query(command As String) As String
'        Dim sb As New StringBuilder(256)
'        mciSendStringW(command, sb, CUInt(sb.Capacity), IntPtr.Zero)
'        Return sb.ToString().Trim()
'    End Function

'    ' --- Core API ------------------------------------------------------------

'    Public Function AddSound(soundName As String, filePath As String) As Boolean
'        If String.IsNullOrWhiteSpace(soundName) OrElse Not IO.File.Exists(filePath) Then
'            Debug.Print($"{soundName} not added.")
'            Return False
'        End If

'        soundName = Normalize(soundName)

'        If Aliases.Contains(soundName) Then Return True

'        Dim cmd = $"open ""{filePath}"" alias {soundName}"

'        If Send(cmd) Then
'            Aliases.Add(soundName)
'            Return True
'        End If

'        Debug.Print($"{soundName} failed to open.")
'        Return False
'    End Function

'    Public Function PlaySound(soundName As String) As Boolean
'        soundName = Normalize(soundName)

'        If Not Aliases.Contains(soundName) Then
'            Debug.Print($"{soundName} not playing.")
'            Return False
'        End If

'        Return Send($"seek {soundName} to start") AndAlso
'               Send($"play {soundName} notify")
'    End Function

'    Public Function LoopSound(soundName As String) As Boolean
'        soundName = Normalize(soundName)

'        If Not Aliases.Contains(soundName) Then
'            Debug.Print($"{soundName} not looping.")
'            Return False
'        End If

'        Return Send($"seek {soundName} to start") AndAlso
'               Send($"play {soundName} repeat")
'    End Function

'    Public Function PauseSound(soundName As String) As Boolean
'        soundName = Normalize(soundName)

'        If Not Aliases.Contains(soundName) Then
'            Debug.Print($"{soundName} not paused.")
'            Return False
'        End If

'        Return Send($"pause {soundName} notify")
'    End Function

'    Public Function SetVolume(soundName As String, level As Integer) As Boolean
'        soundName = Normalize(soundName)

'        If Not Aliases.Contains(soundName) Then
'            Debug.Print($"{soundName} volume not set.")
'            Return False
'        End If

'        level = Math.Max(0, Math.Min(1000, level))

'        Return Send($"setaudio {soundName} volume to {level}")
'    End Function

'    Public Function IsPlaying(soundName As String) As Boolean
'        soundName = Normalize(soundName)
'        If Not Aliases.Contains(soundName) Then Return False
'        Return Query($"status {soundName} mode").Equals("playing", StringComparison.OrdinalIgnoreCase)
'    End Function

'    ' --- Overlapping Playback ------------------------------------------------

'    Public Sub AddOverlapping(baseName As String, filePath As String)
'        For Each suffix In OverlapSuffixes
'            AddSound(baseName & suffix, filePath)
'        Next
'    End Sub

'    Public Sub PlayOverlapping(baseName As String)
'        For Each suffix In OverlapSuffixes
'            Dim aliasName = Normalize(baseName & suffix)
'            If Not IsPlaying(aliasName) Then
'                PlaySound(aliasName)
'                Exit Sub
'            End If
'        Next
'    End Sub

'    Public Sub SetVolumeOverlapping(baseName As String, level As Integer)
'        For Each suffix In OverlapSuffixes
'            SetVolume(baseName & suffix, level)
'        Next
'    End Sub

'    ' --- Cleanup -------------------------------------------------------------

'    Public Sub CloseAll()
'        For Each aliasName In Aliases
'            Send($"close {aliasName}")
'        Next

'        Aliases.Clear()
'    End Sub

'End Module





Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports System.Threading
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.IO

Public Class AudioPlayer
    Implements IDisposable

    ' ============================================================
    ' MCI API
    ' ============================================================
    <DllImport("winmm.dll", EntryPoint:="mciSendStringW")>
    Private Shared Function mciSendStringW(
        <MarshalAs(UnmanagedType.LPWStr)> command As String,
        <MarshalAs(UnmanagedType.LPWStr)> returnString As StringBuilder,
        returnLength As UInteger,
        callback As IntPtr) As Integer
    End Function

    ' ============================================================
    ' Instance State
    ' ============================================================
    Private ReadOnly Aliases As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly SoundInfo As New Dictionary(Of String, (filePath As String, volume As Integer))(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly Looping As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly Cooldowns As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

    Private ReadOnly OverlapSuffixes As String() =
        {"a", "b", "c", "d", "e", "f", "g", "h"}

    'Private cleanupTimer As System.Threading.Timer
    'Private cleanupInitialized As Boolean

    Private ReadOnly syncRoot As New Object()

    ' ============================================================
    ' Constructor / Destructor
    ' ============================================================
    Public Sub New()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose

        CloseAll()
    End Sub


    ' ============================================================
    ' Helpers
    ' ============================================================
    Private Function Normalize(name As String) As String
        If name Is Nothing Then Return ""
        Return name.Trim().Replace(" ", "_")
    End Function

    Private Function ShouldLogError(command As String, code As Integer) As Boolean
        If code = 263 Then
            Dim c = command.Trim().ToLowerInvariant()
            If c.StartsWith("status ") OrElse c.StartsWith("stop ") OrElse c.StartsWith("close ") Then
                Return False
            End If
        End If
        Return True
    End Function

    Private Function Send(command As String) As Boolean
        Dim sb As New StringBuilder(256)
        Dim result = mciSendStringW(command, sb, CUInt(sb.Capacity), IntPtr.Zero)

        If result <> 0 AndAlso ShouldLogError(command, result) Then
            Debug.Print($"MCI Error {result}: {command}")
            Return False
        End If

        Return result = 0
    End Function

    Private Function Query(command As String) As String
        Dim sb As New StringBuilder(256)
        Dim result = mciSendStringW(command, sb, CUInt(sb.Capacity), IntPtr.Zero)

        If result <> 0 AndAlso ShouldLogError(command, result) Then
            Debug.Print($"MCI Error {result}: {command}")
            Return ""
        End If

        Return sb.ToString().Trim()
    End Function

    Private Function AnyPlaying() As Boolean
        Dim snapshot As List(Of String)

        SyncLock syncRoot
            snapshot = Aliases.ToList()
        End SyncLock

        For Each aliasName In snapshot
            If Query($"status {aliasName} mode").Equals("playing", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Function CooldownReady(soundName As String, ms As Integer) As Boolean
        Dim now = Environment.TickCount

        SyncLock syncRoot
            Dim last As Integer
            If Cooldowns.TryGetValue(soundName, last) Then
                If now - last < ms Then Return False
            End If
            Cooldowns(soundName) = now
        End SyncLock

        Return True
    End Function

    Private Function GetDeviceType(filePath As String) As String
        Dim ext = Path.GetExtension(filePath).ToLowerInvariant()
        Select Case ext
            Case ".wav" : Return "waveaudio"
            Case ".mp3" : Return "mpegvideo"
            Case Else : Return ""
        End Select
    End Function

    Private Function OpenSoundInternal(soundName As String, filePath As String, volume As Integer) As Boolean
        Dim deviceType = GetDeviceType(filePath)
        Dim ok As Boolean

        If String.IsNullOrEmpty(deviceType) Then
            ok = Send($"open ""{filePath}"" alias {soundName}")
        Else
            ok = Send($"open ""{filePath}"" type {deviceType} alias {soundName}")
        End If

        If ok Then
            SyncLock syncRoot
                Aliases.Add(soundName)
                SoundInfo(soundName) = (filePath, volume)
            End SyncLock
        End If

        Return ok
    End Function

    ' ============================================================
    ' Volume Fade (Async)
    ' ============================================================
    Private Async Function FadeVolumeAsync(soundName As String, startVol As Integer, endVol As Integer, durationMs As Integer) As Task
        Dim steps As Integer = Math.Max(1, durationMs \ 10)
        Dim delta As Double = (endVol - startVol) / steps
        Dim current As Double = startVol

        For i = 1 To steps
            current += delta
            SetVolume(soundName, CInt(current))
            Await Task.Delay(10)
        Next

        SetVolume(soundName, endVol)
    End Function

    Public Sub FadeVolume(soundName As String, startVol As Integer, endVol As Integer, durationMs As Integer)
        Dim fadeTask As Task = FadeVolumeAsync(soundName, startVol, endVol, durationMs)
    End Sub

    ' ============================================================
    ' Fade Out
    ' ============================================================
    Public Sub FadeOut(soundName As String, durationMs As Integer)
        soundName = Normalize(soundName)

        SyncLock syncRoot
            If Not Aliases.Contains(soundName) Then Exit Sub
        End SyncLock

        Dim info = SoundInfo(soundName)
        Dim currentVol = info.volume

        FadeVolume(soundName, currentVol, 0, durationMs)
    End Sub

    Public Sub FadeOutAndStop(soundName As String, durationMs As Integer)
        FadeOut(soundName, durationMs)

        Task.Run(Async Function()
                     Await Task.Delay(durationMs)
                     Send($"stop {Normalize(soundName)}")
                 End Function)
    End Sub


    ' ============================================================
    ' Core API
    ' ============================================================
    Public Function AddSound(soundName As String, filePath As String) As Boolean
        soundName = Normalize(soundName)

        If String.IsNullOrWhiteSpace(soundName) OrElse Not IO.File.Exists(filePath) Then
            Debug.Print($"{soundName} not added.")
            Return False
        End If

        SyncLock syncRoot
            If Aliases.Contains(soundName) Then Return True
        End SyncLock

        If OpenSoundInternal(soundName, filePath, 500) Then
            Return True
        End If

        Debug.Print($"{soundName} failed to open.")
        Return False
    End Function

    Public Function PlaySound(soundName As String) As Boolean
        soundName = Normalize(soundName)
        If Not CooldownReady(soundName, 40) Then Return False

        SyncLock syncRoot
            If Not Aliases.Contains(soundName) Then Return False
        End SyncLock

        Send($"stop {soundName}")
        Send($"seek {soundName} to start")

        Dim info = SoundInfo(soundName)

        SetVolume(soundName, 0)
        FadeVolume(soundName, 0, info.volume, 80)

        Return Send($"play {soundName}")
    End Function

    Public Function LoopSound(soundName As String) As Boolean
        soundName = Normalize(soundName)

        SyncLock syncRoot
            If Not Aliases.Contains(soundName) Then Return False
        End SyncLock

        Send($"stop {soundName}")
        Send($"seek {soundName} to start")

        Dim ok = Send($"play {soundName} repeat")
        If ok Then
            SyncLock syncRoot
                Looping.Add(soundName)
            End SyncLock
        End If

        Return ok
    End Function

    Public Function StopSound(soundName As String) As Boolean
        soundName = Normalize(soundName)

        SyncLock syncRoot
            If Not Aliases.Contains(soundName) Then Return False
        End SyncLock

        Return Send($"stop {soundName}")
    End Function

    Public Function PauseSound(soundName As String) As Boolean
        soundName = Normalize(soundName)

        SyncLock syncRoot
            If Not Aliases.Contains(soundName) Then Return False
        End SyncLock

        Return Send($"pause {soundName}")
    End Function

    Public Function SetVolume(soundName As String, level As Integer) As Boolean
        soundName = Normalize(soundName)
        level = Math.Max(0, Math.Min(1000, level))

        SyncLock syncRoot
            If Not Aliases.Contains(soundName) Then Return False
        End SyncLock

        Dim ok = Send($"setaudio {soundName} volume to {level}")

        If ok Then
            SyncLock syncRoot
                Dim info = SoundInfo(soundName)
                SoundInfo(soundName) = (info.filePath, level)
            End SyncLock
        End If

        Return ok
    End Function

    Public Function IsPlaying(soundName As String) As Boolean
        soundName = Normalize(soundName)

        SyncLock syncRoot
            If Not Aliases.Contains(soundName) Then Return False
        End SyncLock

        Return Query($"status {soundName} mode").Equals("playing", StringComparison.OrdinalIgnoreCase)
    End Function

    ' ============================================================
    ' Overlapping Playback
    ' ============================================================
    Public Sub AddOverlapping(baseName As String, filePath As String)
        For Each suffix In OverlapSuffixes
            AddSound(baseName & suffix, filePath)
        Next
    End Sub

    Public Sub PlayOverlapping(baseName As String)
        Dim list As List(Of String)

        SyncLock syncRoot
            list = OverlapSuffixes.Select(Function(s) Normalize(baseName & s)).
                                   Where(Function(a) Aliases.Contains(a)).
                                   ToList()
        End SyncLock

        For Each aliasName In list
            If Not Query($"status {aliasName} mode").Equals("playing", StringComparison.OrdinalIgnoreCase) Then

                If Not CooldownReady(aliasName, 40) Then Exit Sub

                Send($"stop {aliasName}")
                Send($"seek {aliasName} to start")
                Send($"play {aliasName}")
                Exit Sub
            End If
        Next
    End Sub

    Public Sub SetVolumeOverlapping(baseName As String, level As Integer)
        For Each suffix In OverlapSuffixes
            SetVolume(baseName & suffix, level)
        Next
    End Sub

    ' ============================================================
    ' Cleanup
    ' ============================================================
    Public Function CloseByAlias(aliasName As String) As Boolean
        aliasName = Normalize(aliasName)

        SyncLock syncRoot
            If Not Aliases.Contains(aliasName) Then Return False
        End SyncLock

        Send($"stop {aliasName}")

        Dim ok = Send($"close {aliasName}")

        If ok Then
            SyncLock syncRoot
                Aliases.Remove(aliasName)
                SoundInfo.Remove(aliasName)
                Looping.Remove(aliasName)
            End SyncLock
        End If

        Return ok
    End Function

    Public Sub CloseAll()
        Dim list As List(Of String)

        SyncLock syncRoot
            list = Aliases.ToList()
            Aliases.Clear()
            SoundInfo.Clear()
            Looping.Clear()
        End SyncLock

        For Each aliasName In list
            Send($"stop {aliasName}")
        Next

        For Each aliasName In list
            Send($"close {aliasName}")
        Next
    End Sub

End Class
