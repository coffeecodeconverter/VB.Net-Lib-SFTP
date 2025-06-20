# VB.Net-Lib-SFTP
VB.Net Lib - SFTP Module - handles async download, upload, error handling, and real-time progress. 

A robust, asynchronous **SFTP (SSH File Transfer Protocol)** module for VB.NET, built on top of **SSH.NET**, designed to simplify file transfers across remote servers. Perfect for integrating into WinForms, WPF, or Console applications.

---

## 📌 Overview

This module provides a powerful, fault-tolerant interface for secure file transfers over SSH. Whether you're downloading a single file, syncing multiple files, or uploading data, this library ensures smooth operation even in unstable network environments.

---

## ✨ Features

- ✅ **Asynchronous operations** for non-blocking UI and smooth performance
- 📦 **Single & batch file downloads** with per-file status reporting
- 🔁 **Auto-resume on reconnect** — handles brief disconnects gracefully
- ⏳ **Real-time progress tracking** (percent, bytes, speed, ETA, elapsed time)
- ❌ **Doesn’t fail entire batch on a single error** — logs each file result
- 📤 **Asynchronous file uploads** with optional cancellation & callbacks
- 📃 **Clear error reporting** — great for debugging and end-user feedback
- 🧩 **Easy to integrate** with any VB.NET app using .NET Framework 4.7.2+

---

## 📦 Requirements

- **.NET Framework**: 4.7.2 or higher
- **NuGet Package**: [`SSH.NET 2025.0.0`](https://www.nuget.org/packages/SSH.NET/)

> 💡 Install via NuGet Package Manager:
> ```
> Install-Package SSH.NET -Version 2025.0.0
> ```

---

## 🛠️ Installation

1. Clone or download the source files for this module.
2. Add the `.vb` file(s) to your project.
3. Install the `SSH.NET` library from NuGet.
4. Start using the provided `SFTP_*` functions in your application.

---

## 🚀 Usage Examples

### 🔗 Test Connection

```vb.net
Dim result = Await SFTP_TestConnection("host.com", "username", "password", 22)
If result.Success Then
    Console.WriteLine("Connected successfully!")
Else
    Console.WriteLine("Connection failed: " & result.ErrorMessage)
End If
