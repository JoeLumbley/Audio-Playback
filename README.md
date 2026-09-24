# Audio Playback Engine (VB.NET / WinForms)

A lightweight, fully‑managed audio playback engine for VB.NET WinForms applications.  
This project demonstrates how to build a **multi‑sound audio system** using the Windows **Multimedia MCI API**, supporting:


- Looping background music  
- Overlapping sound effects  
- Volume control  
- Smooth fade‑in / fade‑out transitions  
- Automatic audio engine restart for long‑running stability  


<img width="1920" height="1080" alt="015" src="https://github.com/user-attachments/assets/8f91869c-abc9-4287-9cee-148bbf17719a" />

The app is designed as a clean, approachable example for developers learning VB.NET, WinForms, or game‑style audio playback. 


[Code Walkthrough](#audioplayer---full-code-walkthrough) 

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











---
---
---









---

# AudioPlayer - Full Code Walkthrough  
*A detailed line‑by‑line explanation of the AudioPlayer class and its subsystems.*

- [Imports](#imports)  

- [Class Declaration](#class-declaration)  

- [MCI API](#mci-api)

- [Instance State](#instance-state)

- [Constructor / Destructor](#constructor--destructor)


Constructor / Destructor

- [Public API](#public-api)
  
    - [Adding sounds](#addsound)
     
    - [Playing](#playsound)
     
    - [Looping](#loopsound)
     
    - [Stopping](#stopsound)
     
    - [Pausing](#pausesound)
     
    - [Volume control](#setvolume)
     
    - [Status checks](#isplaying)


[Top](#audio-playback-engine-vbnet--winforms)  

---

## Imports

```vbnet
Imports System.Runtime.InteropServices
```
Brings in interop services so you can call unmanaged (native) Windows APIs. This is required for the `DllImport` attribute used later to call `mciSendStringW` from `winmm.dll`.  



```vbnet
Imports System.Text
```
Imports text‑related types like `StringBuilder`, which is used to receive strings from the MCI API calls efficiently.  



```vbnet
Imports System.Threading.Tasks
```
Enables use of `Task` and `Async`/`Await` for asynchronous operations—here it’s used for non‑blocking volume fades and delayed stop operations.  



```vbnet
Imports System.IO
```
Provides file and path utilities (e.g., `File.Exists`, `Path.GetExtension`) used to validate sound files and determine device types (`waveaudio`, `mpegvideo`).  



```vbnet
Imports System.Diagnostics
```
Allows logging and debugging via `Debug.Print`, which is used to report MCI errors and failed sound registrations.  

[Walkthrough](#audioplayer---full-code-walkthrough)  

---






## Class Declaration

```vbnet
Public Class AudioPlayer
    Implements IDisposable
```

- **`Public Class AudioPlayer`** declares a reusable audio engine type that other parts of your program can instantiate to manage sounds.  
- **`Implements IDisposable`** signals that the class owns unmanaged resources (MCI devices) and provides a `Dispose` method so callers can cleanly release them—internally it calls `CloseAll()` to stop and close every open alias.  

[Walkthrough](#audioplayer---full-code-walkthrough)  

---















## MCI API




```vbnet
<DllImport("winmm.dll", EntryPoint:="mciSendStringW")>
```

This attribute tells .NET:

- You are calling a native function from the Windows multimedia library `winmm.dll`.
- The specific exported function you want is `mciSendStringW`, the Unicode version of the MCI command dispatcher.

This is the bridge between VB.NET and the Windows multimedia subsystem.  



```vbnet
Private Shared Function mciSendStringW(
```

Declares a **shared (static)** function inside your class.  
Imported native functions cannot be instance methods, so `Shared` is required.  



```vbnet
<MarshalAs(UnmanagedType.LPWStr)> command As String,
```

This parameter is the **MCI command string** you want Windows to execute.

`MarshalAs(UnmanagedType.LPWStr)` forces .NET to marshal the VB.NET `String` as a **wide (UTF‑16) C‑style string**, which is what `mciSendStringW` expects.  



```vbnet
<MarshalAs(UnmanagedType.LPWStr)> returnString As StringBuilder,
```

This is the **output buffer** where MCI writes its response.

Using `StringBuilder` is required because:

- MCI writes directly into the buffer.
- Strings are immutable in .NET, but `StringBuilder` is mutable.  



```vbnet
returnLength As UInteger,
```

Specifies the size of the output buffer.  
If the buffer is too small, MCI truncates the response.  



```vbnet
callback As IntPtr) As Integer
```

Pointer to a callback window handle for asynchronous notifications.  
You pass `IntPtr.Zero` because you are not using MCI notify callbacks.  
Returns an **Integer error code** (`0` = success).  



```vbnet
End Function
```

Closes the declaration of the imported native function.  

[Walkthrough](#audioplayer---full-code-walkthrough)  

---











## Instance State



### Aliases

```vbnet
Private ReadOnly Aliases As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
```

Tracks all active MCI aliases.  
Case‑insensitive, fast lookup, no duplicates.  



### SoundInfo

```vbnet
Private ReadOnly SoundInfo As New Dictionary(Of String, (filePath As String, volume As Integer))(StringComparer.OrdinalIgnoreCase)
```

Maps each alias to:

- its file path  
- its last known volume  

Used for restoring volume during fades.  



### Looping

```vbnet
Private ReadOnly Looping As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
```

Tracks which aliases are currently looping via:

```
play <alias> repeat
```  



### Cooldowns

```vbnet
Private ReadOnly Cooldowns As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
```

Stores last tick count per alias to prevent rapid‑fire MCI commands.  



### OverlapSuffixes

```vbnet
Private ReadOnly OverlapSuffixes As String() =
    {"a", "b", "c", "d", "e", "f", "g", "h"}
```

Defines suffixes for overlapping sound variants (`hit_a`, `hit_b`, …).  



### syncRoot

```vbnet
Private ReadOnly syncRoot As New Object()
```

Used with `SyncLock` to ensure thread‑safe access to shared collections.  


[Walkthrough](#audioplayer---full-code-walkthrough) 

---




















## Constructor / Destructor



### Constructor

```vbnet
Public Sub New()
End Sub
```

Empty constructor - all fields are already initialized inline.  



### Destructor / Dispose

```vbnet
Public Sub Dispose() Implements IDisposable.Dispose
    CloseAll()
End Sub
```

Implements `IDisposable`.  
Calls `CloseAll()` to stop and close every open alias, ensuring no dangling MCI devices remain.  

[Walkthrough](#audioplayer---full-code-walkthrough) 

---































## Public API  
*A complete line‑by‑line walkthrough of every public‑facing method in the AudioPlayer class.*

These are the functions your game or application will call directly:  


- [Adding sounds](#addsound)  
- [Playing](#playsound)  
- [Looping](#loopsound)  
- [Stopping](#stopsound)  
- [Pausing](#pausesound)  
- [Volume control](#setvolume)  
- [Status checks](#isplaying)



Each method includes validation, normalization, thread‑safety, and MCI command dispatching.


[Walkthrough](#audioplayer---full-code-walkthrough)  |  [Top](#audio-playback-engine-vbnet--winforms)



---

### AddSound

```vbnet
Public Function AddSound(soundName As String, filePath As String) As Boolean
```

Declares a public method that attempts to register a new sound with MCI.


```vbnet
soundName = Normalize(soundName)
```

Normalizes the alias name by trimming whitespace and replacing spaces with underscores.  
This ensures consistent alias formatting across the entire engine.


```vbnet
If String.IsNullOrWhiteSpace(soundName) OrElse Not File.Exists(filePath) Then
    Debug.Print($"{soundName} not added.")
    Return False
End If
```

Validates input:

- Rejects empty or whitespace names.  
- Rejects missing files.  
- Logs a debug message if invalid.


```vbnet
SyncLock syncRoot
    If Aliases.Contains(soundName) Then Return True
End SyncLock
```

Thread‑safe check:  
If the alias already exists, the function returns `True` because the sound is already loaded.


```vbnet
If OpenSoundInternal(soundName, filePath, 500) Then
    Return True
End If
```

Attempts to open the sound using MCI.  
Initial volume is set to **500** (mid‑range).  
If successful, return `True`.


```vbnet
Debug.Print($"{soundName} failed to open.")
Return False
```

Logs failure and returns `False`.

[Public API](#public-api)  

---


















### PlaySound

```vbnet
Public Function PlaySound(soundName As String) As Boolean
```

Plays a sound once, with a fade‑in effect.


```vbnet
soundName = Normalize(soundName)
If Not CooldownReady(soundName, 40) Then Return False
```

Normalizes the name and enforces a **40 ms cooldown** to prevent rapid MCI spam.


```vbnet
SyncLock syncRoot
    If Not Aliases.Contains(soundName) Then Return False
End SyncLock
```

Ensures the alias exists before attempting playback.


```vbnet
Send($"stop {soundName}")
Send($"seek {soundName} to start")
```

Resets playback position:

1. Stop the sound  
2. Seek to the beginning  

This ensures consistent behavior even if the sound was partially played earlier.


```vbnet
Dim info = SoundInfo(soundName)
```

Retrieves stored volume and file path.


```vbnet
SetVolume(soundName, 0)
FadeVolume(soundName, 0, info.volume, 80)
```

Implements a **fade‑in**:

- Start at volume `0`
- Fade to the stored volume over **80 ms**


```vbnet
Return Send($"play {soundName}")
```

Sends the MCI play command and returns success/failure.


[Public API](#public-api)  

---













### LoopSound

```vbnet
Public Function LoopSound(soundName As String) As Boolean
```

Loops a sound continuously.


```vbnet
soundName = Normalize(soundName)
```

Normalize alias.


```vbnet
SyncLock syncRoot
    If Not Aliases.Contains(soundName) Then Return False
End SyncLock
```

Ensure alias exists.


```vbnet
Send($"stop {soundName}")
Send($"seek {soundName} to start")
```

Reset playback position.


```vbnet
Dim ok = Send($"play {soundName} repeat")
```

Uses MCI’s built‑in looping:

```
play <alias> repeat
```


```vbnet
If ok Then
    SyncLock syncRoot
        Looping.Add(soundName)
    End SyncLock
End If
```

Tracks loop state internally.


```vbnet
Return ok
```

Return success/failure.


[Public API](#public-api)  

---












### StopSound

```vbnet
Public Function StopSound(soundName As String) As Boolean
```

Stops playback immediately.


```vbnet
soundName = Normalize(soundName)
```

Normalize alias.


```vbnet
SyncLock syncRoot
    If Not Aliases.Contains(soundName) Then Return False
End SyncLock
```

Ensure alias exists.


```vbnet
Return Send($"stop {soundName}")
```

Send MCI stop command.

[Public API](#public-api)  

---














### PauseSound

```vbnet
Public Function PauseSound(soundName As String) As Boolean
```

Pauses playback.


```vbnet
soundName = Normalize(soundName)
```

Normalize alias.


```vbnet
SyncLock syncRoot
    If Not Aliases.Contains(soundName) Then Return False
End SyncLock
```

Ensure alias exists.


```vbnet
Return Send($"pause {soundName}")
```

Send MCI pause command.

[Public API](#public-api)  

---










### SetVolume

```vbnet
Public Function SetVolume(soundName As String, level As Integer) As Boolean
```

Sets volume for a specific alias.


```vbnet
soundName = Normalize(soundName)
level = Math.Max(0, Math.Min(1000, level))
```

Normalize alias and clamp volume to **0–1000**.


```vbnet
SyncLock syncRoot
    If Not Aliases.Contains(soundName) Then Return False
End SyncLock
```

Ensure alias exists.


```vbnet
Dim ok = Send($"setaudio {soundName} volume to {level}")
```

Send MCI volume command.


```vbnet
If ok Then
    SyncLock syncRoot
        Dim info = SoundInfo(soundName)
        SoundInfo(soundName) = (info.filePath, level)
    End SyncLock
End If
```

Update stored volume on success.


```vbnet
Return ok
```

Return success/failure.

[Public API](#public-api)  

---

















### IsPlaying

```vbnet
Public Function IsPlaying(soundName As String) As Boolean
```

Checks whether a sound is currently playing.


```vbnet
soundName = Normalize(soundName)
```

Normalize alias.


```vbnet
SyncLock syncRoot
    If Not Aliases.Contains(soundName) Then Return False
End SyncLock
```

Ensure alias exists.


```vbnet
Return Query($"status {soundName} mode").Equals("playing", StringComparison.OrdinalIgnoreCase)
```

Queries MCI:

```
status <alias> mode
```

Returns `"playing"` if active.


[Public API](#public-api)  

---











































































## Creator

This project is developed by **Joseph W. Lumbley**  
YouTube: [Code with Joe](https://www.youtube.com/@codewithjoe6074)




