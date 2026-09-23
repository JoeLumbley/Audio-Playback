# Audio Playback Engine (VB.NET / WinForms)

A lightweight, fully‑managed audio playback engine for VB.NET WinForms applications.  
This project demonstrates how to build a **multi‑sound audio system** using the Windows **Multimedia MCI API**, supporting:

- Looping background music  
- Overlapping sound effects  
- Volume control  
- Smooth fade‑in / fade‑out transitions  
- Automatic audio engine restart for long‑running stability  

The app is designed as a clean, approachable example for developers learning VB.NET, WinForms, or game‑style audio playback. 

<img width="1920" height="1080" alt="015" src="https://github.com/user-attachments/assets/8f91869c-abc9-4287-9cee-148bbf17719a" />

---

## Features

### **AudioPlayer Class**
A complete audio engine built around `mciSendStringW`, providing:

- **AddSound / PlaySound / StopSound / PauseSound**
- **LoopSound** for continuous background audio
- **Overlapping playback** using multiple aliases (e.g., `hit_a`, `hit_b`, …)
- **Volume control** per sound
- **Async fade‑in / fade‑out** using `Task`‑based transitions
- **Automatic cleanup** of aliases and resources
- **Cooldown system** to prevent rapid retriggering of sounds

The engine manages all MCI aliases internally and ensures safe, thread‑synchronized access.

### **Automatic Audio Engine Restart**
To avoid long‑session audio corruption (a common MCI issue), the app includes a timer that:

1. Fades out the looped audio  
2. Stops playback  
3. Disposes the entire audio engine  
4. Recreates a fresh engine  
5. Reloads all sounds  
6. Restores the loop state  

This keeps the audio stable even after hours of continuous use.

### **Demo UI**
The WinForms interface includes:

- A **Loop / Pause** toggle button  
- A **Play Overlapping Sound** button  
- Automatic extraction of embedded MP3 resources  
- Real‑time status logging via `Debug.Print`

---

## Project Structure

- **AudioPlayer.vb** — Full audio engine implementation  
- **Form1.vb** — Demo UI and engine lifecycle management  
- **Resources** — Embedded MP3 files extracted at runtime  
- **Timers** — Loop control and engine restart logic  

---

## Technologies Used

- **VB.NET (.NET Framework / WinForms)**
- **Windows MCI API (`winmm.dll`)**
- **Async / Await for fade transitions**
- **Thread‑safe audio alias management**
- **Embedded resource extraction**

---































---
---
---



# AudioPlayer class walkthrough

This section explains how the `AudioPlayer` class works and how to use it safely in your project.




## Overview

**Purpose:**  
`AudioPlayer` is a lightweight MCI-based sound engine for playing, looping, fading, and overlapping `.wav` and `.mp3` sounds using `winmm.dll`.

**Key features:**

- **Add and register sounds** by alias.
- **Play, loop, pause, stop** sounds.
- **Per-sound volume control** with fade-in and fade-out.
- **Overlapping playback** for rapid-fire sound effects.
- **Cooldowns** to avoid spamming MCI commands.
- **Thread-safe** internal state via `SyncLock`.

---


```vbnet


Public Class AudioPlayer
    Implements IDisposable


```





## MCI integration

```vbnet
<DllImport("winmm.dll", EntryPoint:="mciSendStringW")>
Private Shared Function mciSendStringW(
    <MarshalAs(UnmanagedType.LPWStr)> command As String,
    <MarshalAs(UnmanagedType.LPWStr)> returnString As StringBuilder,
    returnLength As UInteger,
    callback As IntPtr) As Integer
End Function
```

**What this does:**

- **Wraps `mciSendStringW`** from `winmm.dll` to send MCI commands as Unicode strings.
- `Send` and `Query` are thin helpers around this function:
  - **`Send(command As String)`**: fire-and-forget command, returns `Boolean` success.
  - **`Query(command As String)`**: sends a command and returns the trimmed response string.

Error handling:

```vbnet
Private Function ShouldLogError(command As String, code As Integer) As Boolean
    If code = 263 Then
        Dim c = command.Trim().ToLowerInvariant()
        If c.StartsWith("status ") OrElse c.StartsWith("stop ") OrElse c.StartsWith("close ") Then
            Return False
        End If
    End If
    Return True
End Function
```

- **Suppresses noisy error 263** for common status/stop/close calls.
- Other errors are logged via `Debug.Print("MCI Error ...")`.

---

## Internal state and thread safety

```vbnet
Private ReadOnly Aliases As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
Private ReadOnly SoundInfo As New Dictionary(Of String, (filePath As String, volume As Integer))(StringComparer.OrdinalIgnoreCase)
Private ReadOnly Looping As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
Private ReadOnly Cooldowns As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

Private ReadOnly OverlapSuffixes As String() =
    {"a", "b", "c", "d", "e", "f", "g", "h"}

Private ReadOnly syncRoot As New Object()
```

**Collections:**

- **`Aliases`**: all active MCI aliases.
- **`SoundInfo`**: maps alias → `(filePath, volume)`.
- **`Looping`**: tracks which aliases are currently set to loop.
- **`Cooldowns`**: last tick count per alias to throttle rapid calls.
- **`OverlapSuffixes`**: suffixes used for overlapping variants (e.g. `hit_a`, `hit_b`, ...).

All shared state is guarded by:

```vbnet
SyncLock syncRoot
    ' read/write collections here
End SyncLock
```

---

## Adding and opening sounds

```vbnet
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
```

**Flow:**

1. **Normalize** the name (`Trim`, replace spaces with `_`).
2. **Validate**: non-empty name and existing file.
3. **Skip if already added**.
4. **Open via `OpenSoundInternal`** with initial volume `500`.

Device type detection:

```vbnet
Private Function GetDeviceType(filePath As String) As String
    Select Case Path.GetExtension(filePath).ToLowerInvariant()
        Case ".wav" : Return "waveaudio"
        Case ".mp3" : Return "mpegvideo"
        Case Else : Return ""
    End Select
End Function
```

Opening:

```vbnet
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
```

---

## Playing, looping, pausing, stopping

### Play with fade-in

```vbnet
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
```

**Key points:**

- **Cooldown**: `CooldownReady(soundName, 40)` enforces a 40 ms minimum gap.
- **Reset playback**: `stop` then `seek ... to start`.
- **Fade-in**: volume goes from `0` to stored `info.volume` over `80 ms`.
- **Finally** sends `play` command.

### Looping

```vbnet
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
```

- Uses `play ... repeat` to loop.
- Tracks looping aliases in `Looping`.

### Pause and stop

```vbnet
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
```

### Status check

```vbnet
Public Function IsPlaying(soundName As String) As Boolean
    soundName = Normalize(soundName)

    SyncLock syncRoot
        If Not Aliases.Contains(soundName) Then Return False
    End SyncLock

    Return Query($"status {soundName} mode").Equals("playing", StringComparison.OrdinalIgnoreCase)
End Function
```

---

## Volume control and fading

### Direct volume set

```vbnet
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
```

- Clamps volume to **0–1000**.
- Updates `SoundInfo` on success.

### Async fade

```vbnet
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
```

- **Step size:** one volume update every `10 ms`.
- **Duration:** `durationMs` determines number of steps.
- `FadeVolume` is a fire-and-forget wrapper around the async method.

### Fade out helpers

```vbnet
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
```

- `FadeOut`: fades from current volume to `0`.
- `FadeOutAndStop`: fades, then stops after the fade duration.

---

## Overlapping playback

### Register overlapping variants

```vbnet
Public Sub AddOverlapping(baseName As String, filePath As String)
    For Each suffix In OverlapSuffixes
        AddSound(baseName & suffix, filePath)
    Next
End Sub
```

- Creates multiple aliases like `hit_a`, `hit_b`, ..., all pointing to the same file.

### Play overlapping

```vbnet
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
```

- Finds the first variant that is **not currently playing**.
- Applies **per-alias cooldown**.
- Plays that alias and exits—so only one overlapping instance is started per call.

### Overlapping volume

```vbnet
Public Sub SetVolumeOverlapping(baseName As String, level As Integer)
    For Each suffix In OverlapSuffixes
        SetVolume(baseName & suffix, level)
    Next
End Sub
```

---

## Cooldowns

```vbnet
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
```

- Prevents **rapid repeated commands** for the same alias.
- Uses `Environment.TickCount` to track last call time per alias.

---

## Cleanup and disposal

### Close a single alias

```vbnet
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
```

### Close all sounds

```vbnet
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
```

### IDisposable

```vbnet
Public Sub Dispose() Implements IDisposable.Dispose
    CloseAll()
End Sub
```

- Ensures all MCI devices are stopped and closed when the player is disposed.

---

## Usage examples

### Basic setup

```vbnet
Dim player As New AudioPlayer()

' Add sounds
player.AddSound("menu_music", "Assets\Music\menu.mp3")
player.AddSound("hit", "Assets\Sounds\hit.wav")

' Play once with fade-in
player.PlaySound("menu_music")

' Loop background music
player.LoopSound("menu_music")
```

### Overlapping sound effects

```vbnet
' Register overlapping variants
player.AddOverlapping("hit", "Assets\Sounds\hit.wav")

' In your game loop / input handler:
player.PlayOverlapping("hit")
```

### Fade out and cleanup

```vbnet
' Fade out music over 1 second and stop
player.FadeOutAndStop("menu_music", 1000)

' On game exit
player.CloseAll()
player.Dispose()
```

---

## Design notes

- **Normalization:** all sound names are normalized (`Trim`, spaces → `_`) to keep aliases consistent.
- **Thread safety:** `SyncLock syncRoot` protects shared collections from concurrent access.
- **Error noise reduction:** error 263 from MCI is selectively ignored for common status/stop/close commands to avoid log spam.





---





## Creator

This project is developed by **Joseph W. Lumbley**  
YouTube: [Code with Joe](https://www.youtube.com/@codewithjoe6074)




