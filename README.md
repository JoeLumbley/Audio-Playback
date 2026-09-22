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

## Creator

This project is developed by **Joseph W. Lumbley**  
YouTube: [Code with Joe](https://www.youtube.com/@codewithjoe6074)

---



