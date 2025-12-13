# Gesture Controlled Volume System (Windows Core Audio + Pycaw)

## Project Title
Infosys_Springboard_Gesture_Controlled_Volume

## Timeline
11–13 November 2025

---

## Objective

This project implements a real-time hand gesture–based audio volume control system for Windows.  
The system uses computer vision to detect hand gestures through a webcam and maps those gestures to system-level audio controls using Windows Core Audio APIs accessed via Pycaw.

The primary interaction mechanism is the distance between the thumb and index finger, enabling touchless and intuitive volume adjustment.

---

## System Overview

The overall workflow of the system is as follows:

Webcam capture  
→ Hand detection and landmark extraction  
→ Finger distance computation  
→ Distance normalization and smoothing  
→ Volume mapping logic  
→ Windows Core Audio volume update  

---

## Hand Detection and Gesture Recognition

### Palm Detection Model

The system uses a BlazePalm-based palm detector.  
BlazePalm is a lightweight SSD-style object detection model optimized for real-time applications.

Key characteristics:
- Produces a bounding box around the detected palm
- Uses depthwise separable convolutions
- Optimized for low latency and CPU efficiency

### Hand Landmark Model

After palm detection, a CNN-based hand landmark model runs within the bounding box.

This model:
- Outputs 21 hand landmarks
- Accurately tracks fingertip positions
- Enables precise gesture measurement

---

## Gesture Definition

The gesture used for volume control is defined as the Euclidean distance between:
- Thumb tip landmark
- Index finger tip landmark

This distance directly represents the desired volume level.

---

## Live Camera Overlay Information

The live video feed displays the following information:

- Hand bounding box
- Hand landmarks (21 points)
- A line connecting thumb and index finger
- Numeric distance value
- Current volume percentage

The distance is initially calculated in pixels.  
This pixel distance is normalized to a predefined range and mapped to a volume scalar between 0.0 and 1.0.

Optional camera calibration can be applied to convert pixel distance into real-world units such as centimeters.

---

## Smooth Overlay Rendering

To ensure smooth visualization without flicker or lag, the following techniques are used:

- Frame-by-frame rendering using OpenCV
- Continuous redraw without clearing the frame abruptly
- Exponential moving average smoothing for distance values
- Avoidance of blocking operations in the main loop

Smoothing logic:

Distance_smooth = α × Current_distance + (1 − α) × Previous_distance

This prevents sudden jumps in volume and overlay jitter.

---

## Depth Variation Problem and Solution

### Problem

When the user moves their hand closer or farther from the camera, the absolute pixel distance between fingers changes even if the gesture remains the same.  
This causes inconsistent volume control.

### Solution

The system normalizes finger distance using hand size:

Normalized_distance = Finger_distance / Palm_width

This makes the volume control invariant to depth and ensures consistent behavior across different hand positions.

---

## Windows Core Audio System

Windows Core Audio is a low-level, COM-based API responsible for managing audio devices and sessions in the Windows operating system.

It provides access to:
- Hardware-connected audio devices
- System-wide volume control
- Per-application audio sessions
- Mute and unmute functionality

---

## Core Audio COM Interfaces

IMMDeviceEnumerator  
Enumerates all audio input and output devices connected to the system.

IAudioEndpointVolume  
Controls the master system volume and mute state.

IAudioSessionManager2  
Manages per-application audio sessions.

ISimpleAudioVolume  
Controls the volume of an individual application session.

All of these interfaces are implemented using COM (Component Object Model).

---

## COM and Pointer Handling

Windows Core Audio APIs are written in C++ and rely heavily on raw pointers.  
Python does not expose raw pointers directly for safety reasons.

Instead:
- comtypes manages COM interfaces
- ctypes handles low-level data types
- Pointer casting is handled internally and safely

This allows Python to interact with system-level APIs without unsafe memory access.

---

## Pycaw Overview

Pycaw (Python Core Audio Windows) is a Python wrapper around Windows Core Audio APIs.

Internally, it interfaces with:
- AudioSession.h
- EndpointVolume.h
- PolicyConfig.h

Using Pycaw, a Python program can:
- Access default playback devices
- Read and modify volume levels
- Mute or unmute system audio
- Control individual application audio sessions

---

## Gesture and Audio Integration Workflow

1. Capture a frame from the webcam  
2. Detect hand and extract landmarks  
3. Compute thumb–index finger distance  
4. Normalize and smooth the distance  
5. Convert distance to a volume scalar  
6. Send volume value to Pycaw  
7. Update system audio in real time  

---

## Installation

### Prerequisites

- Windows operating system
- Python 3.x
- Webcam

### Required Libraries

Install all dependencies using pip:

pip install pycaw comtypes opencv-python mediapipe

---

## Important Python Components

comtypes  
Used to access COM-based Windows APIs.

CLSCTX_ALL  
Specifies the execution context when activating COM objects.

ctypes  
Provides C-compatible data types and pointer handling.

POINTER  
Used to define and work with pointer-based interfaces.

---

## Master Volume Control Example

The following code demonstrates how to control the system master volume using Pycaw:

from comtypes import CLSCTX_ALL  
from ctypes import POINTER, cast  
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume  

devices = AudioUtilities.GetSpeakers()  

interface = devices.Activate(  
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None  
)  

volume = cast(interface, POINTER(IAudioEndpointVolume))  

current_volume_db = volume.GetMasterVolumeLevel()  
print(current_volume_db)  

Explanation:
- GetSpeakers retrieves the default playback device
- Activate loads the required COM interface
- IID uniquely identifies the interface
- cast converts the raw COM object into a usable pointer

---

## Identified Limitation

IAudioEndpointVolume controls only the master system volume.  
It does not control individual application audio such as YouTube, Chrome, or Spotify.

---

## Per-Application Volume Control

To control individual applications, the following interfaces are required:

- IAudioSessionManager2
- ISimpleAudioVolume

These allow enumeration and control of per-application audio sessions.

Implementation reference:
masterAudioControlKivy.py

---

## Future Enhancements

- Per-application gesture-based volume control
- Gesture-based mute and unmute
- Output device switching (speaker/headphones)
- AI-based gesture classification
- On-screen volume meter visualization

---

## References

Microsoft Windows Core Audio Documentation  
Pycaw: Python Core Audio Wrapper  
COM (Component Object Model) Fundamentals  
MediaPipe Hand Tracking Documentation  

---

## Author Notes

This project demonstrates the integration of computer vision, real-time systems, OS-level audio control, and COM-based API interaction.

It is suitable for internship evaluations, system-level AI demonstrations, and human–computer interaction projects.
