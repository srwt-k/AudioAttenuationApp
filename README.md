# Audio Attenuation App

A Windows desktop application that automatically *ducks* (attenuates) the volume of **all other applications** when a selected application produces sound, and restores volumes when it becomes silent.

Think of it as *smart media ducking*, similar to how iOS lowers background audio during calls — but this is process-specific.

---

## Features

- Select **one target process** to monitor
- Automatically **lower volume of other apps** when the target produces sound
- Restore volumes when the target is silent for a configurable duration
- Smooth volume fading (no sudden jumps)

---

## Requirements

- Windows 10 / 11
- .NET (WinForms)
- NAudio-compatible environment

> ⚠️ This app relies on Windows Core Audio APIs, so **Windows-only**.

---

## Configuration

### Adjustable Parameters

- **Silent threshold (ms)**
  - Time the target app must be silent before restoring volumes
  - Example: `1000` ms = 1 second

- **Lower volume limit**
  - Percentage to reduce other apps to (e.g. 30%)

---

## Typical Use Cases

Lower music/video when someone speak

---

## Known Limitations

- System sounds may not always be affected (Windows-managed)
- Some protected processes may not expose volume control
- Requires at least one active audio session to track

---
