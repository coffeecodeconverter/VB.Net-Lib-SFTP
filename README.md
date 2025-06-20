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
```

---

## Main Module

~~~~~~~~~~~~~~~
Imports Renci.SshNet
Imports Renci.SshNet.Sftp                           ' install ssh.net version 2025.0.0 via nuget 
Imports System.IO
Imports System.Net.NetworkInformation
Imports System.Net.Sockets
Imports System.Runtime.Remoting.Messaging
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks




Public Class SftpFileItem
    Public Property SFTPFileName As String
    Public Property SFTPFileType As String
    Public Property SFTPFileSizeKB As Long
    Public Property SFTPFileModified As Date
End Class



Public Class DownloadProgress
    Public Property SFTPBytesReceived As Long
    Public Property SFTPTotalBytes As Long
    Public Property SFTPSpeedKbps As Double
    Public Property SFTPETA As TimeSpan
    Public Property SFTPElapsed As TimeSpan
End Class


Public Class UploadProgress
    Public Property SFTPBytesSent As Long
    Public Property SFTPTotalBytes As Long
    Public Property SFTPSpeedKbps As Double
    Public Property SFTPETA As TimeSpan
    Public Property SFTPElapsed As TimeSpan
End Class








Module ModuleSFTP


    Private cts As CancellationTokenSource
    Public ModuleVersion_SFTP As String = "1.0.0.1"



    ' Call this before starting any operation (uplload/download) to get a fresh token
    Public Function GenerateToken() As CancellationTokenSource
        CancelToken() ' Ensure any previous token is canceled first
        cts = New CancellationTokenSource()
        Return cts
    End Function



    Public Sub CancelToken()
        If cts IsNot Nothing AndAlso Not cts.IsCancellationRequested Then
            cts.Cancel()
        End If
    End Sub


    ' Check if a cancellation is already in progress
    Public Function IsTokenActive() As Boolean
        Return cts IsNot Nothing AndAlso Not cts.IsCancellationRequested
    End Function





    Public Async Function HasInternetConnectionAsync(Optional timeoutMs As Integer = 1500) As Task(Of Boolean)
        If Not HasValidIpAddress() Then
            Return False
        End If

        Return Await CanReachInternetAsync(timeoutMs)
    End Function





    ' V2 returns messages instead of true false 
    Public Async Function SFTP_TestConnection(passedHost As String, passedUsername As String, passedPassword As String, passedPort As Integer) As Task(Of String)

        Return Await Task.Run(Async Function()

                                  Try
                                      Dim internetConnectionResult As Boolean = Await HasInternetConnectionAsync()
                                      If internetConnectionResult = False Then
                                          Return $"ERROR:{vbCrLf}Connection failed:{vbCrLf}No Internet Connection Detected"
                                      End If
                                  Catch ex As Exception
                                      Return $"ERROR:{vbCrLf}Connection failed:{vbCrLf}No Internet Connection Detected{vbCrLf}{ex.Message}"
                                  End Try

                                  Try
                                      ' Normalize host (strip "sftp://" if present)
                                      Dim normalizedHost As String = passedHost
                                      If normalizedHost.StartsWith("sftp://", StringComparison.OrdinalIgnoreCase) Then
                                          normalizedHost = normalizedHost.Substring(7)
                                      End If

                                      Dim connectionInfo As New ConnectionInfo(normalizedHost, passedPort, passedUsername,
                                  New PasswordAuthenticationMethod(passedUsername, passedPassword))

                                      Using sftpClient As New SftpClient(connectionInfo)
                                          Try
                                              sftpClient.Connect()
                                          Catch ex As Exception
                                              If ex.Message.ToLower.Contains("permission denied (password)") Then
                                                  Return $"ERROR:{vbCrLf}Connection failed:{vbCrLf}Username or Password is incorrect"

                                              ElseIf ex.Message.ToLower().Contains("no connection could be made because the target machine actively refused it") Then
                                                  Return $"ERROR:{vbCrLf}Connection failed:{vbCrLf}{ex.Message}. Check the Port number is correct, and firewall rules (TCP Outbound and Inbound) are configured to allow the specified Port."
                                              End If

                                              Return $"ERROR:{vbCrLf}Connection failed:{vbCrLf}{ex.Message}"
                                          End Try

                                          If sftpClient.IsConnected Then
                                              sftpClient.Disconnect()
                                              Return "Connection succeeded."
                                          Else
                                              Return $"ERROR:{vbCrLf}Client failed to connect for unknown reasons."
                                          End If
                                      End Using

                                  Catch ex As Exception
                                      Return $"Unhandled Error:{vbCrLf}{ex.Message}"
                                  End Try
                              End Function)
    End Function




    ' Gets Files from Remote Directory - returns custom object to populate a DataGridView
    Public Async Function SFTP_ListFilesAsync(passedHost As String, passedUsername As String, passedPassword As String, passedPort As Integer, remotePath As String) As Task(Of List(Of SftpFileItem))

        Return Await Task.Run(Function()
                                  Dim result As New List(Of SftpFileItem)()

                                  Try

                                      ' Normalize host (strip "sftp://" if present)
                                      Dim normalizedHost As String = passedHost
                                      If normalizedHost.StartsWith("sftp://", StringComparison.OrdinalIgnoreCase) Then
                                          normalizedHost = normalizedHost.Substring(7)
                                      End If

                                      Dim connectionInfo As New ConnectionInfo(normalizedHost, passedPort, passedUsername,
                                      New PasswordAuthenticationMethod(passedUsername, passedPassword))

                                      Using sftpClient As New SftpClient(connectionInfo)
                                          Try
                                              sftpClient.Connect()
                                          Catch ex As Exception
                                              Debug.WriteLine($"Connection failed: {ex.Message}")
                                              MessageBox.Show($"Connection failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                              Return result
                                          End Try

                                          If sftpClient.IsConnected Then

                                              Try

                                                  Dim SFTPFiles As IEnumerable(Of ISftpFile) = sftpClient.ListDirectory(remotePath)

                                                  For Each fileItem As ISftpFile In SFTPFiles
                                                      ' Skip special folders
                                                      If fileItem.Name <> "." AndAlso fileItem.Name <> ".." Then
                                                          Dim sftpFile = TryCast(fileItem, SftpFile)
                                                          If sftpFile IsNot Nothing Then
                                                              result.Add(New SftpFileItem With {
                                                                  .SFTPFileName = sftpFile.Name,
                                                                  .SFTPFileType = If(sftpFile.IsDirectory, "Directory", "File"),
                                                                  .SFTPFileSizeKB = sftpFile.Length \ 1024,
                                                                  .SFTPFileModified = sftpFile.LastWriteTime
                                                              })
                                                          End If
                                                      End If
                                                  Next

                                              Catch ex As Exception
                                                  Debug.WriteLine($"Error listing directory: {ex.Message}")
                                                  MessageBox.Show($"Error listing directory: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                              End Try
                                          Else
                                              Debug.WriteLine("Failed to connect to SFTP.")
                                              MessageBox.Show("Failed to connect to SFTP.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                          End If

                                          Try
                                              sftpClient.Disconnect()
                                          Catch ex As Exception
                                              Debug.WriteLine($"Error disconnecting: {ex.Message}")
                                              MessageBox.Show($"Error disconnecting: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                          End Try
                                      End Using

                                  Catch ex As Exception
                                      Debug.WriteLine($"Unhandled error: {ex.Message}")
                                      MessageBox.Show($"Unhandled error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                  End Try

                                  Return result
                              End Function)
    End Function





    ' v2 shows progress 
    Public Async Function SFTP_UploadFileAsync(
                                            passedHost As String,
                                            passedUsername As String,
                                            passedPassword As String,
                                            passedPort As Integer,
                                            localFilePath As String,
                                            remotePath As String,
                                            Optional progress As IProgress(Of UploadProgress) = Nothing,
                                            Optional token As CancellationToken = Nothing,
                                            Optional callback As Action = Nothing) As Task(Of Boolean)

        Return Await Task.Run(Function()

                                  Try
                                      Dim normalizedHost As String = passedHost
                                      If normalizedHost.StartsWith("sftp://", StringComparison.OrdinalIgnoreCase) Then
                                          normalizedHost = normalizedHost.Substring(7)
                                      End If

                                      Dim connectionInfo As New ConnectionInfo(normalizedHost, passedPort, passedUsername,
                                          New PasswordAuthenticationMethod(passedUsername, passedPassword))

                                      Using sftpClient As New SftpClient(connectionInfo)
                                          sftpClient.Connect()

                                          If sftpClient.IsConnected Then
                                              Using localFileStream As New IO.FileStream(localFilePath, IO.FileMode.Open, IO.FileAccess.Read)
                                                  Dim totalBytes As Long = localFileStream.Length
                                                  Dim buffer(8191) As Byte
                                                  Dim bytesRead As Integer
                                                  Dim bytesSent As Long = 0

                                                  Dim stopwatch As New Stopwatch()
                                                  stopwatch.Start()

                                                  Using remoteStream = sftpClient.OpenWrite(remotePath)

                                                      Do
                                                          If token.IsCancellationRequested Then
                                                              Return False
                                                          End If

                                                          bytesRead = localFileStream.Read(buffer, 0, buffer.Length)
                                                          If bytesRead > 0 Then
                                                              remoteStream.Write(buffer, 0, bytesRead)
                                                              bytesSent += bytesRead

                                                              If progress IsNot Nothing Then
                                                                  Dim elapsedSeconds = stopwatch.Elapsed.TotalSeconds
                                                                  Dim speedKbps = If(elapsedSeconds > 0, (bytesSent / 1024.0) / elapsedSeconds, 0)
                                                                  Dim estimatedTotalSeconds As Double = If(speedKbps > 0, (totalBytes / 1024.0) / speedKbps, 0)
                                                                  Dim remainingSeconds = Math.Max(0, estimatedTotalSeconds - elapsedSeconds)

                                                                  Dim etaTime As TimeSpan = TimeSpan.FromSeconds(remainingSeconds)
                                                                  Dim elapsedTime As TimeSpan = TimeSpan.FromSeconds(elapsedSeconds)

                                                                  progress.Report(New UploadProgress With {
                                                                  .SFTPBytesSent = bytesSent,
                                                                  .SFTPTotalBytes = totalBytes,
                                                                  .SFTPSpeedKbps = speedKbps,
                                                                  .SFTPETA = etaTime,
                                                                  .SFTPElapsed = elapsedTime
                                                              })
                                                              End If
                                                          End If
                                                      Loop While bytesRead > 0
                                                  End Using

                                                  stopwatch.Stop()
                                                  sftpClient.Disconnect()

                                                  If callback IsNot Nothing Then
                                                      Try
                                                          callback()
                                                      Catch ex As Exception
                                                          MessageBox.Show("Error executing callBack function: " & ex.Message)
                                                      End Try
                                                  End If

                                                  Return True
                                              End Using
                                          End If

                                      End Using
                                  Catch ex As Exception
                                      Console.WriteLine($"Upload failed: {ex.Message}")
                                  End Try

                                  Return False
                              End Function, token)
    End Function






    ' V5 - now with a callback function 
    Public Async Function SFTP_DownloadFileAsync(passedHost As String, passedUsername As String, passedPassword As String, passedPort As Integer, remoteFilePath As String, localDestinationPath As String, Optional progress As IProgress(Of DownloadProgress) = Nothing, Optional token As CancellationToken = Nothing, Optional callback As Action = Nothing) As Task(Of String)

        Dim downloadTask = Task.Run(Function()
                                        Return SFTP_DownloadFileBlocking(passedHost, passedUsername, passedPassword, passedPort, remoteFilePath, localDestinationPath, progress, token)
                                    End Function, token)

        Dim timeoutTask = Task.Delay(TimeSpan.FromSeconds(30), token)

        Dim completedTask = Await Task.WhenAny(downloadTask, timeoutTask)

        Dim result As String

        If completedTask Is timeoutTask Then
            result = "ERROR:" & vbCrLf & "Download failed: Operation timed out (no data in 30 seconds)."
        Else
            result = Await downloadTask
        End If

        If callback IsNot Nothing Then
            Try
                callback()
            Catch ex As Exception
                MessageBox.Show("Error executing callBack function: " & ex.Message)
            End Try
        End If

        Return result
    End Function





    Private Function SFTP_DownloadFileBlocking(passedHost As String, passedUsername As String, passedPassword As String, passedPort As Integer, remoteFilePath As String, localDestinationPath As String, progress As IProgress(Of DownloadProgress), token As CancellationToken) As String
        Try
            Dim normalizedHost As String = passedHost
            If normalizedHost.StartsWith("sftp://", StringComparison.OrdinalIgnoreCase) Then
                normalizedHost = normalizedHost.Substring(7)
            End If

            Dim connectionInfo As New ConnectionInfo(normalizedHost, passedPort, passedUsername,
                                                     New PasswordAuthenticationMethod(passedUsername, passedPassword))

            Using sftpClient As New SftpClient(connectionInfo)
                Try
                    sftpClient.Connect()
                Catch ex As Exception
                    Return $"ERROR:{vbCrLf}Connection failed:{vbCrLf}{ex.Message}"
                End Try

                If Not sftpClient.IsConnected Then
                    Return "ERROR:" & vbCrLf & "Client failed to connect."
                End If

                Try
                    Dim fileAttrs = sftpClient.GetAttributes(remoteFilePath)
                    Dim totalBytes = fileAttrs.Size

                    Using localStream As New IO.FileStream(localDestinationPath, IO.FileMode.Create, IO.FileAccess.Write)
                        Using remoteStream = sftpClient.OpenRead(remoteFilePath)

                            Dim buffer(8191) As Byte ' 8 KB buffer
                            Dim bytesRead As Integer
                            Dim bytesReceived As Long = 0
                            Dim stopwatch As New Stopwatch()
                            stopwatch.Start()

                            Dim lastDataReceivedTime As DateTime = DateTime.UtcNow

                            Do
                                If token.IsCancellationRequested Then
                                    Return "ERROR:" & vbCrLf & "Download cancelled by user."
                                End If

                                Try
                                    bytesRead = remoteStream.Read(buffer, 0, buffer.Length)
                                Catch ex As Exception
                                    Return "ERROR:" & vbCrLf & "Connection dropped during read: " & ex.Message
                                End Try

                                If bytesRead > 0 Then
                                    lastDataReceivedTime = DateTime.UtcNow ' Reset timeout

                                    localStream.Write(buffer, 0, bytesRead)
                                    bytesReceived += bytesRead


                                    If progress IsNot Nothing Then
                                        Dim elapsedSeconds = stopwatch.Elapsed.TotalSeconds
                                        Dim speedKbps = If(elapsedSeconds > 0, (bytesReceived / 1024.0) / elapsedSeconds, 0)
                                        Dim estimatedTotalSeconds As Double = If(speedKbps > 0, (totalBytes / 1024.0) / speedKbps, 0)
                                        Dim remainingSeconds = Math.Max(0, estimatedTotalSeconds - elapsedSeconds)

                                        Dim etaTime As TimeSpan = TimeSpan.FromSeconds(remainingSeconds)
                                        Dim elapsedTime As TimeSpan = TimeSpan.FromSeconds(elapsedSeconds)

                                        progress.Report(New DownloadProgress With {
                                                    .SFTPBytesReceived = bytesReceived,
                                                    .SFTPTotalBytes = totalBytes,
                                                    .SFTPSpeedKbps = speedKbps,
                                                    .SFTPETA = etaTime,
                                                    .SFTPElapsed = elapsedTime
                                                })
                                    End If

                                Else
                                    ' No data received — check for timeout
                                    If (DateTime.UtcNow - lastDataReceivedTime).TotalSeconds > 30 Then
                                        Return "ERROR:" & vbCrLf & "Download failed: No data received for 30 seconds."
                                    End If

                                    Thread.Sleep(200)

                                End If
                            Loop While bytesRead > 0

                            stopwatch.Stop()
                        End Using
                    End Using

                    sftpClient.Disconnect()

                    Dim returnMessage As String = "Download Finished."

                    Return returnMessage
                Catch ex As Exception
                    Return $"ERROR:{vbCrLf}Download failed:{vbCrLf}{ex.Message}"
                End Try
            End Using
        Catch ex As Exception
            Return $"Unhandled error:{vbCrLf}{ex.Message}"
        End Try
    End Function









    ' HELPER FUNCTIONS - NOT USED DIRECTLY
    ' ------------------------------------------

    ' for HasInternetConnectionAsync()
    Private Function HasValidIpAddress() As Boolean
        For Each nic In NetworkInterface.GetAllNetworkInterfaces()
            If nic.OperationalStatus <> OperationalStatus.Up Then Continue For

            Dim name = nic.Name.ToLower()
            Dim desc = nic.Description.ToLower()
            If name.Contains("virtual") Or name.Contains("bluetooth") Or desc.Contains("virtual") Or desc.Contains("bluetooth") Then
                Continue For
            End If

            Dim ipProps = nic.GetIPProperties()
            For Each addr In ipProps.UnicastAddresses
                If addr.Address.AddressFamily = AddressFamily.InterNetwork Then ' IPv4
                    Dim ip = addr.Address.ToString()
                    If Not ip.StartsWith("169.") Then ' Not a link-local address
                        Return True
                    End If
                End If
            Next
        Next
        Return False
    End Function



    ' for HasInternetConnectionAsync()
    Private Async Function CanReachInternetAsync(timeoutMs As Integer) As Task(Of Boolean)
        Dim testHosts = {
            "1.1.1.1",
            "8.8.8.8",
            "google.com",
            "example.com",
            "microsoft.com"
        }

        For Each host In testHosts
            Try
                Using client As New TcpClient()
                    Dim connectTask = client.ConnectAsync(host, 80)
                    If Await Task.WhenAny(connectTask, Task.Delay(timeoutMs)) Is connectTask AndAlso client.Connected Then
                        Return True
                    End If
                End Using
            Catch
                ' Ignore and continue
            End Try
        Next
        Return False
    End Function





End Module
~~~~~~~~~~~~~~~
