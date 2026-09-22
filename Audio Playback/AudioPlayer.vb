' Audio Player

' Uses Windows MCI API (winmm.dll) for playback of multiple audio files simultaneously.

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

Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Tasks
Imports System.IO
Imports System.Diagnostics

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
    ' Core API
    ' ============================================================
    Public Function AddSound(soundName As String, filePath As String) As Boolean
        soundName = Normalize(soundName)

        If String.IsNullOrWhiteSpace(soundName) OrElse Not File.Exists(filePath) Then
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
        Select Case Path.GetExtension(filePath).ToLowerInvariant()
            Case ".wav" : Return "waveaudio"
            Case ".mp3" : Return "mpegvideo"
            Case Else : Return ""
        End Select
    End Function

    Private Function OpenSoundInternal(soundName As String, filePath As String, volume As Integer) As Boolean
        Dim deviceType = GetDeviceType(filePath)
        Dim ok As Boolean

        If deviceType = "" Then
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
        Dim r = FadeVolumeAsync(soundName, startVol, endVol, durationMs)
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
        FadeVolume(soundName, info.volume, 0, durationMs)
    End Sub

    Public Sub FadeOutAndStop(soundName As String, durationMs As Integer)
        FadeOut(soundName, durationMs)

        Task.Run(Async Function()
                     Await Task.Delay(durationMs)
                     Send($"stop {Normalize(soundName)}")
                 End Function)
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


' Copilot is our AI assistant.


' I also make coding videos on my YouTube channel: Code with Joe.
' https://www.youtube.com/@codewithjoe6074

