Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO.Compression
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Security
Imports IWshRuntimeLibrary

Public Class FileTools

    Dim dlls As New Dictionary(Of String, String)

    Private Declare Function OpenProcess Lib "kernel32&quot" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
    Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Integer, ByVal lpAddress As Integer, ByVal dwSize As Integer, ByVal flAllocationType As Integer, ByVal flProtect As Integer) As Integer
    Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Integer, ByVal lpBaseAddress As Integer, ByVal lpBuffer() As Byte, ByVal nSize As Integer, ByVal lpNumberOfBytesWritten As UInteger) As Boolean
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Integer, ByVal lpProcName As String) As Integer
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Integer
    Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Integer, ByVal lpThreadAttributes As Integer, ByVal dwStackSize As Integer, ByVal lpStartAddress As Integer, ByVal lpParameter As Integer, ByVal dwCreationFlags As Integer, ByVal lpThreadId As Integer) As Integer
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
    Shared Function WriteFile(ByVal dir As String, ByVal contents As String, Optional append As Boolean = False)
        Try
            My.Computer.FileSystem.WriteAllText(dir, contents, append)
        Catch ex As Exception : End Try
    End Function

    Shared Function ReadLine(ByVal dir As String, ByVal lineNum As Integer)
        Try
            Dim lines As String() = System.IO.File.ReadAllLines(dir)
            lineNum = lineNum - 1
            Dim line As String = lines(lineNum)
            Return line
        Catch ex As Exception : End Try
    End Function

    Shared Function FindStringInFile(ByVal dir As String, ByVal data As String)
        dir = ReadAll(dir)
        Return InStr(dir, data)
    End Function

    Shared Function MassExtensionChange(ByVal dir As String, ByVal extension As String, Optional ByVal replace As Boolean = False)
        Try
            Dim file As FileInfo
            For Each file In AllFilesInADirectory(dir)
                ChangeFileExtension(file.Directory.ToString, extension, replace)
            Next
        Catch ex As Exception : End Try
    End Function

    Shared Function MassCopyByExtension(ByVal dir As String, ByVal newDir As String, ByVal extension As String)
        Try
            Dim file As FileInfo
            For Each file In AllFilesInADirectory(dir)
                If GetExtension(file.Directory.ToString) = extension Then
                    
                    CopyFile(file.Directory.ToString, newDir)
                Else
                    

                End If
            Next
        Catch ex As Exception : End Try
    End Function
    Shared Function AllFilesInADirectory(ByVal dir As String)
        Try
            Dim di As New DirectoryInfo(dir)
            Dim Files As FileInfo() = di.GetFiles()
            Return Files
        Catch ex As Exception : End Try
    End Function

    Shared Function WriteLine(ByVal dir As String, ByVal lineNum As Integer, ByVal contents As String)
        Try
            Dim lines() As String = System.IO.File.ReadAllLines(dir)
            lines(lineNum - 1) = contents
            System.IO.File.WriteAllLines(dir, lines)
        Catch ex As Exception : End Try
    End Function

    Shared Function CountLines(ByVal dir As String)
        Dim allLines As String() = ReadAll(dir)
        Dim lines As Integer = 0
        For Each line As String In allLines
            lines = lines + 1
        Next
        Return lines
    End Function
    Shared Function InjectDLL(ByVal pID As Integer, ByVal dllLocation As String) As Boolean
        Dim hProcess As Integer = OpenProcess(&H1F0FFF, 1, pID)
        If hProcess = 0 Then Return False
        Dim dllBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(dllLocation)
        Dim allocAddress As Integer = VirtualAllocEx(hProcess, 0, dllBytes.Length, &H1000, &H4)
        If allocAddress = Nothing Then Return False
        Dim kernelMod As Integer = GetModuleHandle("kernel32.dll")
        Dim loadLibAddr = GetProcAddress(kernelMod, "LoadLibraryA")
        If kernelMod = 0 OrElse loadLibAddr = 0 Then Return False
        WriteProcessMemory(hProcess, allocAddress, dllBytes, dllBytes.Length, 0)
        Dim libThread As Integer = CreateRemoteThread(hProcess, 0, 0, loadLibAddr, allocAddress, 0, 0)
        If libThread = 0 Then
            Return False
        Else
            WaitForSingleObject(libThread, 5000)
            CloseHandle(libThread)
        End If
        CloseHandle(hProcess)
    End Function

    Shared Function ReadAll(ByVal dir As String)
        Try
            Dim fileReader As String
            fileReader = My.Computer.FileSystem.ReadAllText(dir)
            Return fileReader
        Catch ex As Exception : End Try
    End Function

    Public Function ExportFileIcon(ByVal dir As String, ByVal newDir As String)
        Try
            Return "This is not working yet"
        Catch ex As Exception : End Try
    End Function

    Public Function ChangeIcon(ByVal dir As String, ByVal icon As String)
        Try
            IconInjector.InjectIcon(dir, icon)
        Catch ex As Exception : End Try
    End Function

    Shared Function FilesExtensionIs(ByVal dir As String, ByVal extension As String)
        If GetExtension(FormatDirectory(dir)) = extension Then
            Return True
        Else
            Return False
        End If
    End Function

    Shared Function FileIsImage(ByVal dir As String)
        If GetExtension(FormatDirectory(dir)) = ".bmp" Or ".gif" Or ".jpeg" Or ".jpg" Or ".png" Or ".tif" Then
            Return True
        Else
            Return False
        End If
    End Function

    Shared Function CreateShortcut(ByVal dir As String, ByVal newDir As String, ByVal Desc As String)
        Try
            Dim myShortcut As IWshShortcut
            Dim wsh As New WshShell
            Dim str As String = My.Computer.FileSystem.GetFileInfo(dir).Name.Replace(My.Computer.FileSystem.GetFileInfo(dir).Extension, Nothing)
            myShortcut = CType(wsh.CreateShortcut(newDir & "\" & str & ".lnk"), IWshShortcut)
            With myShortcut
                .TargetPath = dir
                .WindowStyle = 1
                .Description = Desc
                .IconLocation = dir & " , 0"
                .Save()
            End With
        Catch ex As Exception : End Try
    End Function
    Shared Function GetFileName(ByVal dir As String, Optional withExtension As Boolean = True)
        Try
            Dim result As String
            If withExtension = True Then
                
                result = Path.GetFileName(dir)
                Return result
            Else
                
                result = Path.GetFileNameWithoutExtension(dir)
                Return result
            End If

        Catch ex As Exception : End Try
    End Function


    Shared Function MassDirectoryCreatorFromTextFile(ByVal dir As String, ByVal rootDir As String)
        Try
            rootDir = GetDirectory(rootDir)
            rootDir = FormatDirectory(rootDir)
            Dim tempDir As String
            Dim count As Integer = 0
            Dim allLines As String() = IO.File.ReadAllLines(dir)
            Return allLines
            For Each line As String In IO.File.ReadAllLines(dir)
                CreateDirectory(rootDir + line, False)
            Next
        Catch ex As Exception

        End Try
    End Function
    Shared Function CopyFile(ByVal dir As String, ByVal newDir As String, Optional ByVal keepName As Boolean = True)
        Try
            dir = FormatDirectory(dir)
            If keepName = True Then
                Dim temp1 As String = ReadAll(dir)
                Dim temp2 As String = newDir
                WriteFile(dir, temp1, False)
            Else
                My.Computer.FileSystem.CopyFile(dir, newDir)
            End If
        Catch ex As Exception : End Try
    End Function
    
    
    Shared Function FormatDirectory(ByVal dir As String)
        Dim s As String = dir
        Dim e As Char = s(s.Length - 1)
        If e <> "\" Then
            dir = dir + "\"
        End If
        Return dir
    End Function
    Shared Function MoveFile(ByVal dir As String, ByVal newDir As String, Optional ByVal overwrite As Boolean = False)
        Try
            My.Computer.FileSystem.MoveFile(dir, newDir, overwrite)
        Catch ex As Exception : End Try
    End Function

    Shared Function DeleteFile(ByVal dir As String)
        Try
            My.Computer.FileSystem.DeleteFile(dir)
        Catch ex As Exception : End Try
    End Function

    Shared Function StartProgram(ByVal dir As String, Optional ByVal parameters As String = "")
        Try
            Process.Start(dir, parameters)
        Catch ex As Exception : End Try
    End Function

    Shared Function CloseProgram(ByVal name As String)
        Try
            For Each prog As Process In Process.GetProcesses
                If prog.ProcessName = name Then
                    prog.Kill()
                End If
            Next
        Catch ex As Exception : End Try
    End Function

    Shared Function RunFileAsAdmin(ByVal dir As String)
        Dim procStartInfo As New ProcessStartInfo
        Dim procExecuting As New Process

        With procStartInfo
            .UseShellExecute = True
            .FileName = dir
            .WindowStyle = ProcessWindowStyle.Normal
            .Verb = "runas" 
        End With

        procExecuting = Process.Start(procStartInfo)
    End Function

    Shared Function GetExtension(ByVal dir As String)
        Try
            Dim result As String
            result = Path.GetExtension(dir)
            Return result
        Catch ex As Exception : End Try
    End Function

    Shared Function ChangeFileExtension(ByVal dir As String, ByVal extension As String, Optional replace As Boolean = False)
        Try
            If replace = True Then
                
                Dim originalFile As String = dir
                Dim newName As String = Path.ChangeExtension(originalFile, extension)
                MoveFile(originalFile, newName)
            Else
                
                Dim originalFile As String = dir
                Dim newName As String = Path.ChangeExtension(originalFile, extension)
                CopyFile(originalFile, newName)
            End If
        Catch ex As Exception : End Try
    End Function


    Shared Function GetFileSize(ByVal dir As String)
        Try
            Dim infoReader As System.IO.FileInfo
            infoReader = My.Computer.FileSystem.GetFileInfo(dir)
            Return infoReader.Length
        Catch ex As Exception : End Try
    End Function

    Shared Function GetDirectory(ByVal dir As String)
        Try
            Dim result As String = Path.GetDirectoryName(dir)
            Return result
        Catch ex As Exception : End Try
    End Function

    Shared Function CountFilesInADirectory(ByVal dir As String)
        Try
            dir = FormatDirectory(dir)
            Dim FilesInFolder = Directory.GetFiles(dir).Count()
            Return FilesInFolder
        Catch ex As Exception

        End Try
    End Function
    Shared Function GetDriveLetter(ByVal dir As String)
        Try
            Dim result As String = Path.GetPathRoot(dir)
            Return result
        Catch ex As Exception : End Try
    End Function

    Shared Function GetMD5FromFile(ByVal Filename As String) As String
        Try
            Dim MD5 = System.Security.Cryptography.MD5.Create
            Dim Hash As Byte()
            Dim sb As New System.Text.StringBuilder
            Using st As New IO.FileStream(Filename, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
                Hash = MD5.ComputeHash(st)
            End Using

            For Each b In Hash
                sb.Append(b.ToString("X2"))
            Next
            Return sb.ToString
        Catch ex As Exception : End Try
    End Function

    Shared Function VerifyMD5WithFile(ByVal dir As String, ByVal hash As String)
        Try
            If GetMD5FromFile(dir) = hash Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception : End Try
    End Function

    Shared Function EncryptMD5(ByVal value As String, ByVal key As String)
        Try
            Dim DES As New TripleDESCryptoServiceProvider
            Dim MD5 As New MD5CryptoServiceProvider
            DES.Key = MD5.ComputeHash(ASCIIEncoding.ASCII.GetBytes(key))
            DES.Mode = CipherMode.ECB
            Dim buffer As Byte() = ASCIIEncoding.ASCII.GetBytes(value)
            Return Convert.ToBase64String(DES.CreateEncryptor().TransformFinalBlock(buffer, 0, buffer.Length))
        Catch ex As Exception : End Try
    End Function

    Shared Function DecryptMD5(ByVal value As String, ByVal key As String)
        Try
            Dim DES As New TripleDESCryptoServiceProvider
            Dim MD5 As New MD5CryptoServiceProvider
            DES.Key = MD5.ComputeHash(ASCIIEncoding.ASCII.GetBytes(key))
            DES.Mode = CipherMode.ECB
            Dim buffer As Byte() = Convert.FromBase64String(value)
            Return ASCIIEncoding.ASCII.GetString(DES.CreateDecryptor().TransformFinalBlock(buffer, 0, buffer.Length))
        Catch ex As Exception : End Try
    End Function

    Shared Function DeleteDirectory(ByVal dir As String, Optional ByVal Recycle As Boolean = True)
        Try
            Dim ShouldRecycle
            If Recycle = True Then
                ShouldRecycle = FileIO.RecycleOption.SendToRecycleBin
            Else
                ShouldRecycle = FileIO.RecycleOption.DeletePermanently
            End If
            My.Computer.FileSystem.DeleteDirectory(dir, FileIO.DeleteDirectoryOption.DeleteAllContents, ShouldRecycle)
        Catch ex As Exception : End Try
    End Function

    Shared Function RenameFile(ByVal dir As String, ByVal name As String)
        Try
            My.Computer.FileSystem.RenameFile(dir, name)
        Catch ex As Exception : End Try
    End Function

    Shared Function RenameFolder(ByVal dir As String, ByVal name As String)
        Try
            My.Computer.FileSystem.RenameDirectory(Dir, name)
        Catch ex As Exception : End Try
    End Function

    Shared Function HideFile(ByVal dir As String)
        Try
            IO.File.SetAttributes(dir, IO.FileAttributes.Hidden)
        Catch ex As Exception : End Try
    End Function

    Shared Function ShowFile(ByVal dir As String)
        Try
            IO.File.SetAttributes(dir, IO.FileAttributes.Normal)
        Catch ex As Exception : End Try
    End Function

    Shared Function HideDirectory(ByVal dir As String)
        Try
            IO.File.SetAttributes(dir, IO.FileAttributes.Hidden)
        Catch ex As Exception : End Try
    End Function

    Shared Function ShowDirectory(ByVal dir As String)
        Try
            IO.File.SetAttributes(dir, IO.FileAttributes.Normal)
        Catch ex As Exception : End Try
    End Function

    Private Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long

    Shared Function RenameDrive(ByVal drive As String, ByVal name As String)
        Try
            SetVolumeLabel(drive, name)
        Catch ex As Exception : End Try
    End Function
    Shared Function CreateDirectory(ByVal dir As String, Optional ByVal hidden As Boolean = False)
        Try
            Dim isHidden
            If hidden = True Then
                My.Computer.FileSystem.CreateDirectory(dir)
            Else
                My.Computer.FileSystem.CreateDirectory(dir)
            End If

        Catch ex As Exception : End Try
    End Function

    Shared Function CopyDirectory(ByVal dir As String, ByVal newDir As String)
        Try
            My.Computer.FileSystem.CopyDirectory(dir, newDir)
        Catch ex As Exception : End Try
    End Function

    Shared Function MoveDirectory(ByVal dir As String, ByVal newDir As String, Optional ByVal overwrite As Boolean = False)
        Try
            My.Computer.FileSystem.MoveDirectory(dir, newDir, overwrite)
        Catch ex As Exception : End Try
    End Function

    Shared Function FileExists(ByVal dir As String)
        Try
            Return My.Computer.FileSystem.FileExists(dir)
        Catch ex As Exception : End Try
    End Function

    Shared Function DirectoryExists(ByVal dir As String)
        Try
            Return My.Computer.FileSystem.DirectoryExists(dir)
        Catch ex As Exception : End Try
    End Function

    Shared Function DriveExists(ByVal drive As String)
        Try
            Return DirectoryExists(drive)
        Catch ex As Exception : End Try
    End Function

    Shared Function DownloadString(ByVal url As String)
        Try
            Dim client As WebClient = New WebClient
            Dim reply As String = client.DownloadString(url)
            Return reply
        Catch ex As Exception : End Try
    End Function

    Shared Function DownloadFile(ByVal url As String, ByVal dir As String)
        Try
            My.Computer.Network.DownloadFile(url, dir)
        Catch ex As Exception : End Try
    End Function

    Shared Function DownloadHTMLSourceCode(ByVal url As String)
        Return "This is not working yet"
    End Function

    Shared Function MakeFileReadOnly(ByVal dir As String)
        IO.File.SetAttributes(dir, FileAttributes.ReadOnly)
    End Function

    Shared Function MakeFileReadWrite(ByVal dir As String)
        IO.File.SetAttributes(dir, FileAttributes.Normal)
    End Function

    Shared Function SameFile(ByVal file1 As String, ByVal file2 As String)
        Try
            file1 = ReadAll(file1)
            file2 = ReadAll(file2)
            If file1 = file2 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception : End Try
    End Function

#Region "Zipping"
    
    
    
    
    
    

    Shared Function ZipFiles(ByVal dir As String, ByVal newDir As String)
        
    End Function
#End Region
End Class
