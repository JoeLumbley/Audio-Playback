# Audio Playback

Welcome to **Audio Playback**, a dynamic and versatile tool crafted for seamless audio management using the Windows Multimedia API. This application is packed with features that make it an essential asset for developers and audio enthusiasts alike.






![002](https://github.com/user-attachments/assets/4a0592e4-786c-4f07-8964-6489218bd567)








## Key Features:

- **Simultaneous Playback**: Unlock the full potential of the Windows Multimedia API by playing multiple audio files at the same time. This feature allows you to create rich, immersive audio experiences that captivate your audience.

- **Precision Volume Control**: Tailor the volume levels of individual audio tracks with precision. This ensures that you achieve the perfect audio balance tailored to your unique requirements.

- **Looping and Overlapping**: Effortlessly loop audio tracks and play overlapping sounds. This enables you to craft captivating and dynamic audio compositions that enhance your projects.

- **MCI Integration**: Harness the power of the Media Control Interface (MCI) to interact with multimedia devices. This provides a standardized and platform-independent approach to controlling multimedia hardware, making your development process smoother.

- **User-Friendly Interface**: Enjoy a clean and intuitive interface designed to simplify the management of audio playback operations. This makes it easy for users of all skill levels to navigate and utilize the application effectively.

With its robust functionality and seamless integration with the Windows Multimedia API, the Audio Playback application empowers you to create engaging multimedia applications effortlessly. Whether you’re a seasoned developer or an aspiring enthusiast, this tool is your gateway to unlocking the full potential of audio playback on the Windows platform.

**Clone the repository now and embark on a transformative audio playback journey!** Let's dive into the world of audio together!



---






# Code Walkthrough

Welcome to this detailed walkthrough of the `AudioPlayer` structure and the `Form1` class! We'll go through each line of code and explain its purpose. Let's dive in!

[Index](#index)

## Imports

```vb.net
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.IO
```
These imports bring in necessary namespaces for:
- `System.Runtime.InteropServices`: For interacting with unmanaged code.
- `System.Text`: For using the `StringBuilder` class.
- `System.IO`: For file input and output operations.

[Index](#index)

## AudioPlayer Structure

```vb.net
Public Structure AudioPlayer
```
This line defines a `Structure` named `AudioPlayer`. Structures in VB.NET are value types that can contain data and methods.

[Index](#index)

### DLL Import

```vb.net
<DllImport("winmm.dll", EntryPoint:="mciSendStringW")>
Private Shared Function mciSendStringW(<MarshalAs(UnmanagedType.LPWStr)> ByVal lpszCommand As String,
                                       <MarshalAs(UnmanagedType.LPWStr)> ByVal lpszReturnString As StringBuilder,
                                       ByVal cchReturn As UInteger, ByVal hwndCallback As IntPtr) As Integer
End Function
```
This imports the `mciSendStringW` function from the `winmm.dll` library. This function sends a command string to the Media Control Interface (MCI) to control multimedia devices.

[Index](#index)

### Sounds Array

```vb.net
Private Sounds() As String
```
This declares an array named `Sounds` to store the names of sounds that have been added.

[Index](#index)

### AddSound Method

```vb.net
Public Function AddSound(SoundName As String, FilePath As String) As Boolean
```
This method adds a sound to the player. It takes the name of the sound and the path to the sound file as parameters.

```vb.net
If Not String.IsNullOrWhiteSpace(SoundName) AndAlso IO.File.Exists(FilePath) Then
```
Checks if the sound name is not empty or whitespace and if the file exists.

```vb.net
Dim CommandOpen As String = $"open ""{FilePath}"" alias {SoundName}"
```
Creates a command string to open the sound file and assign it an alias.

```vb.net
If Sounds Is Nothing Then
```
Checks if the `Sounds` array is uninitialized.

```vb.net
If SendMciCommand(CommandOpen, IntPtr.Zero) Then
```
Sends the command to open the sound file.

```vb.net
ReDim Sounds(0)
Sounds(0) = SoundName
Return True
```
Initializes the `Sounds` array with the new sound and returns `True`.

```vb.net
ElseIf Not Sounds.Contains(SoundName) Then
```
Checks if the sound is not already in the array.

```vb.net
Array.Resize(Sounds, Sounds.Length + 1)
Sounds(Sounds.Length - 1) = SoundName
Return True
```
Adds the new sound to the `Sounds` array and returns `True`.

```vb.net
Debug.Print($"{SoundName} not added to sounds.")
Return False
```
Prints a debug message and returns `False` if the sound could not be added.

[Index](#index)

### SetVolume Method

```vb.net
Public Function SetVolume(SoundName As String, Level As Integer) As Boolean
```
This method sets the volume of a sound. It takes the sound name and volume level (0 to 1000) as parameters.

```vb.net
If Sounds IsNot Nothing AndAlso
   Sounds.Contains(SoundName) AndAlso
   Level >= 0 AndAlso Level <= 1000 Then
```
Checks if the `Sounds` array is not empty, contains the sound, and the volume level is valid.

```vb.net
Dim CommandVolume As String = $"setaudio {SoundName} volume to {Level}"
Return SendMciCommand(CommandVolume, IntPtr.Zero)
```
Creates and sends the command to set the volume.

```vb.net
Debug.Print($"{SoundName} volume not set.")
Return False
```
Prints a debug message and returns `False` if the volume could not be set.

[Index](#index)

### LoopSound Method

```vb.net
Public Function LoopSound(SoundName As String) As Boolean
```
This method loops a sound. It takes the sound name as a parameter.

```vb.net
If Sounds IsNot Nothing AndAlso Sounds.Contains(SoundName) Then
```
Checks if the `Sounds` array is not empty and contains the sound.

```vb.net
Dim CommandSeekToStart As String = $"seek {SoundName} to start"
Dim CommandPlayRepeat As String = $"play {SoundName} repeat"
Return SendMciCommand(CommandSeekToStart, IntPtr.Zero) AndAlso
       SendMciCommand(CommandPlayRepeat, IntPtr.Zero)
```
Creates and sends commands to seek to the start of the sound and play it in a loop.

```vb.net
Debug.Print(Debug.Print($"{SoundName} not looping.")
Return False
```
Prints a debug message and returns `False` if the sound could not be looped.

[Index](#index)

### PlaySound Method

```vb.net
Public Function PlaySound(SoundName As String) As Boolean
```
This method plays a sound. It takes the sound name as a parameter.

```vb.net
If Sounds IsNot Nothing AndAlso Sounds.Contains(SoundName) Then
```
Checks if the `Sounds` array is not empty and contains the sound.

```vb.net
Dim CommandSeekToStart As String = $"seek {SoundName} to start"
Dim CommandPlay As String = $"play {SoundName} notify"
Return SendMciCommand(CommandSeekToStart, IntPtr.Zero) AndAlso
 SendMciCommand(CommandPlay, IntPtr.Zero)

```
Creates and sends commands to seek to the start of the sound and play it.

```vb.net
Debug.Print($"{SoundName} not playing.")
Return False
```
Prints a debug message and returns `False` if the sound could not be played.

[Index](#index)

### PauseSound Method

```vb.net
Public Function PauseSound(SoundName As String) As Boolean
```
This method pauses a sound. It takes the sound name as a parameter.

```vb.net
If Sounds IsNot Nothing AndAlso Sounds.Contains(SoundName) Then
```
Checks if the `Sounds` array is not empty and contains the sound.

```vb.net
Dim CommandPause As String = $"pause {SoundName} notify"
Return SendMciCommand(CommandPause, IntPtr.Zero)
```
Creates and sends the command to pause the sound.

```vb.net
Debug.Print($"{SoundName} not paused.")
Return False
```
Prints a debug message and returns `False` if the sound could not be paused.

[Index](#index)

### IsPlaying Method

```vb.net
Public Function IsPlaying(SoundName As String) As Boolean
```
This method checks if a sound is playing. It takes the sound name as a parameter and returns a Boolean.

```vb.net
Return GetStatus(SoundName, "mode") = "playing"
```
Uses the `GetStatus` method to check if the sound is currently playing.

[Index](#index)

### AddOverlapping Method

```vb.net
Public Sub AddOverlapping(SoundName As String, FilePath As String)
```
This method adds multiple overlapping instances of a sound. It takes the sound name and file path as parameters.

```vb.net
For Each Suffix As String In {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"}
    AddSound(SoundName & Suffix, FilePath)
Next
```
Loops through a set of suffixes (A to L) and adds each sound instance with a unique name.

[Index](#index)

### PlayOverlapping Method

```vb.net
Public Sub PlayOverlapping(SoundName As String)
```
This method plays one instance of an overlapping sound. It takes the sound name as a parameter.

```vb.net
For Each Suffix As String In {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"}
    If Not IsPlaying(SoundName & Suffix) Then
        PlaySound(SoundName & Suffix)
        Exit Sub
    End If
Next
```
Loops through the set of suffixes and plays the first sound instance that is not already playing.

[Index](#index)

### SetVolumeOverlapping Method

```vb.net
Public Sub SetVolumeOverlapping(SoundName As String, Level As Integer)
```
This method sets the volume for all instances of an overlapping sound. It takes the sound name and volume level as parameters.

```vb.net
For Each Suffix As String In {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"}
    SetVolume(SoundName & Suffix, Level)
Next
```
Loops through the set of suffixes and sets the volume for each sound instance.

[Index](#index)

### SendMciCommand Method

```vb.net
Private Function SendMciCommand(command As String, hwndCallback As IntPtr) As Boolean
```
This method sends an MCI command. It takes the command string and a window handle for the callback as parameters.

```vb.net
Dim ReturnString As New StringBuilder(128)
Try
    Return mciSendStringW(command, ReturnString, 0, hwndCallback) = 0
Catch ex As Exception
    Debug.Print($"Error sending MCI command: {command} | {ex.Message}")
    Return False
End Try
```

Here, the `mciSendStringW` function is called with the command string. If the function returns `0`, it means the command was successfully sent. If an exception occurs, the error is printed, and the method returns `False`.

[Index](#index)

### GetStatus Method

```vb.net
Private Function GetStatus(SoundName As String, StatusType As String) As String
```
This method gets the status of a sound. It takes the sound name and the status type (e.g., "mode") as parameters and returns the status as a string.

```vb.net
If Sounds IsNot Nothing AndAlso Sounds.Contains(SoundName) Then
    Dim CommandStatus As String = $"status {SoundName} {StatusType}"
    Dim StatusReturn As New StringBuilder(128)
    mciSendStringW(CommandStatus, StatusReturn, 128, IntPtr.Zero)
    Return StatusReturn.ToString.Trim.ToLower
End If
```
Checks if the `Sounds` array is not empty and contains the sound. Creates and sends the command to get the status, stores the result in `StatusReturn`, and returns the status as a trimmed, lowercase string.

```vb.net
Catch ex As Exception
    Debug.Print($"Error getting status: {SoundName} | {ex.Message}")
End Try
Return String.Empty
```
If an exception occurs, the error is printed, and an empty string is returned.

[Index](#index)

### CloseSounds Method

```vb.net
Public Sub CloseSounds()
```
This method closes all sound files.

```vb.net
If Sounds IsNot Nothing Then
    For Each Sound In Sounds
        Dim CommandClose As String = $"close {Sound}"
        SendMciCommand(CommandClose, IntPtr.Zero)
    Next
End If
End Sub
```
Checks if the `Sounds` array is not empty. Loops through each sound and sends a command to close it.

[Index](#index)

## Form1 Class

```vb.net
Public Class Form1
```
This class defines a form in a Windows Forms application.

[Index](#index)

### Player Declaration

```vb.net
Private Player As AudioPlayer
```
This declares an instance of the `AudioPlayer` structure.

[Index](#index)

### Form Load Event

```vb.net
Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
```
This method handles the form's `Load` event.

```vb.net
Text = "Audio Playback - Code with Joe"
```
Sets the form's title.

[Index](#index)

### CreateSoundFiles Method

```vb.net
CreateSoundFiles()
```
Calls the `CreateSoundFiles` method to create the necessary sound files from embedded resources.

[Index](#index)

### Adding and Setting Up Sounds

```vb.net
Dim FilePath As String = Path.Combine(Application.StartupPath, "level.mp3")
Player.AddSound("Music", FilePath)
Player.SetVolume("Music", 600)
FilePath = Path.Combine(Application.StartupPath, "CashCollected.mp3")
Player.AddOverlapping("CashCollected", FilePath)
Player.SetVolumeOverlapping("CashCollected", 900)
Player.LoopSound("Music")
Debug.Print($"Running... {Now}")
```
Sets up the sound files by specifying their file paths, adding them to the player, setting their volume, and starting to loop the "Music" sound. It prints a debug message indicating the form is running.

[Index](#index)

### Button1 Click Event

```vb.net
Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
Player.PlayOverlapping("CashCollected")
End Sub
```
This method handles the `Click` event for `Button1`. It plays an overlapping instance of the "CashCollected" sound.

[Index](#index)

### Button2 Click Event

```vb.net
Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
If Player.IsPlaying("Music") = True Then
    Player.PauseSound("Music")
    Button2.Text = "Play Loop"
Else
    Player.LoopSound("Music")
    Button2.Text = "Pause Loop"
End If
End Sub
```
This method handles the `Click` event for `Button2`. It toggles between playing and pausing the "Music" sound and updates the button text accordingly.

[Index](#index)

### Form Closing Event

```vb.net
Private Sub Form1_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
Player.CloseSounds()
End Sub
```
This method handles the form's `Closing` event. It closes all sound files to release resources.

[Index](#index)

### CreateSoundFiles Method

```vb.net
Private Sub CreateSoundFiles()
Dim FilePath As String = Path.Combine(Application.StartupPath, "level.mp3")
CreateFileFromResource(FilePath, My.Resources.level)
FilePath = Path.Combine(Application.StartupPath, "CashCollected.mp3")
CreateFileFromResource(FilePath, My.Resources.CashCollected)
End Sub
```
This method creates sound files from embedded resources. It specifies the file paths and calls `CreateFileFromResource` to write the resource data to the file system.

[Index](#index)

### CreateFileFromResource Method

```vb.net
Private Sub CreateFileFromResource(filepath As String, resource As Byte())
Try
    If Not IO.File.Exists(filepath) Then
        IO.File.WriteAllBytes(filepath, resource)
    End If
Catch ex As Exception
    Debug.Print($"Error creating file: {ex.Message}")
End Try
End Sub
```
This method writes resource data to a file if it does not already exist. It handles exceptions by printing an error message.

[Index](#index)

---






## Adding Resources

To add a resource file to your Visual Studio project, follow these steps:

1. **Add a New Resource File**:
   - From the **Project** menu, select `Add New Item...`.
   - In the dialog that appears, choose `Resource File` from the list of templates.
   - Name your resource file (e.g., `Resource1.resx`) and click `Add`.

  
![010](https://github.com/user-attachments/assets/a7ecfd06-8c4a-4230-8110-e22aec1f16b5)



![006](https://github.com/user-attachments/assets/c70414fa-1563-4e71-8286-9c0de9c04db3)

2. **Open the Resource Editor**:
   - Double-click the newly created `.resx` file to open the resource editor.

![009](https://github.com/user-attachments/assets/be296067-b008-4efa-b981-bf4d7e88f3f2)


3. **Add Existing Files**:
   - In the resource editor, click on the **Green Plus Sign** or right-click in the resource pane and select `Add Resource`.
   - Choose `Add Existing File...` from the context menu.
   - Navigate to the location of the MP3 file (or any other resource file) you want to add, select it, and click `Open`.


![011](https://github.com/user-attachments/assets/d2cb9a9c-d395-4d91-8f7d-47e3c94b37bc)

4. **Verify the Addition**:
   - Ensure that your MP3 file appears in the list of resources in the resource editor. It should now be accessible via the Resource class in your code.

5. **Accessing the Resource in Code**:
   - You can access the added resource in your code using the following syntax:
     ```vb
     CreateFileFromResource(filePath, YourProjectNamespace.Resource1.YourResourceName)
     ```

6. **Save Changes**:
   - Don’t forget to save your changes to the `.resx` file.
  

![012](https://github.com/user-attachments/assets/6641c30d-d002-4e06-a783-d2d3761c8c9e)

By following these steps, you can easily add any existing MP3 file or other resources to your Visual Studio project and utilize them within your Audio Playback application.

---









## Index


[Imports](#imports)


[AudioPlayer Structure](#audioPlayer-structure)

[Dll Import](#dll-import)

[Sounds Array](#sounds-array)


[AddSound Method](#addSound-method)

[SetVolume Method](#setvolume-method)

[LoopSound Method](#loopsound-method)

[PlaySound Method](#playsound-method)

[PauseSound Method](#pausesound-method)


[IsPlaying Method](#isplaying-method)

[AddOverlapping Method](#addoverlapping-method)

[PlayOverlapping Method](#playoverlapping-method)


[SetVolumeOverlapping Method](#setVolumeoverlapping-method)


[SendMciCommand Method](#sendmcicommand-method)

[GetStatus Method](#getstatus-method)

[CloseSounds Method](#closesounds-method)

[Form1 Class](#form1-class)

[Player Declaration](#player-declaration)


[Form Load Event](#form-load-event)

[CreateSoundFiles Method](#createsoundfiles-method)

[Adding and Setting Up Sounds](#adding-and-setting-up-sounds)

[Button1 Click Event](#button1-click-event)


[Button2 Click Event](#button2-click-event)

[Form Closing Event](#form-closing-event)

[CreateSoundFiles Method](#createsoundfiles-method)

[CreateFileFromResource Method](#createfilefromresource-method)






---

This code defines an `AudioPlayer` structure that manages sound files, including adding, playing, pausing, looping, and overlapping sounds. The `Form1` class sets up the form, initializes the audio player, and handles events for playing and pausing sounds.

I hope this detailed explanation helps you understand the code better! If you have any questions or need further clarification, feel free to ask. 😊



