Imports System.IO

Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public path As ArrayList
    Public basesourcepath As String
    Public targetdir As String
    Public basefolder As String
    Public RecycleBinPath As String
    Public CancelAction As Boolean
    Private Overwrite As Boolean

    Public Event WorkerStatusMessage(ByVal message As String)
    Public Event WorkerErrorEncountered(ByVal exc As Exception, ByVal message As String)
    Public Event WorkerComplete()




#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        path = New ArrayList
        basesourcepath = ""
        targetdir = ""
        basefolder = ""
        RecycleBinPath = ""
        CancelAction = False
        Overwrite = False
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region



    Private Sub Error_Handler(ByVal exc As Exception, ByVal message As String)
        Try
            If (Not WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                RaiseEvent WorkerErrorEncountered(exc, message)
            End If
        Catch ex As Exception
            MsgBox("An error occurred in File System Contorls' error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub



    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            CancelAction = False
            Overwrite = False
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    WorkerThread = New System.Threading.Thread(AddressOf CopyHandler)
                    WorkerThread.Start()
                Case 2
                    WorkerThread = New System.Threading.Thread(AddressOf DeleteHandler)
                    WorkerThread.Start()
            End Select
        Catch ex As Exception
            Error_Handler(ex, "ChooseThreads")
        End Try
    End Sub


    Private Sub CopyHandler()
        RaiseEvent WorkerStatusMessage("Copying Files")
        Try
            Overwrite = False
            If CancelAction = False Then


                Dim fils As String = ""
                For Each str As String In path
                    If Not str = RecycleBinPath Then
                        fils = fils.Replace(basefolder, ".\").Replace("\\", "\") & "- " & str & vbCrLf
                    End If

                Next
                fils = fils.Trim

                If path.Count <= 10 And path.Count > 0 Then
                    If MsgBox("Do you wish to copy the following files and folders to your hand-in folder?" & vbCrLf & vbCrLf & fils, MsgBoxStyle.OKCancel, "Marked for Copy") = MsgBoxResult.OK Then
                        Dim counter As Integer = 0
                        For Each str As String In path
                            If CancelAction = False Then
                                counter = counter + 1
                                'MsgBox(basesourcepath & "  " & targetdir)
                                Dim target As String = str.Replace(basesourcepath, (targetdir & "\"))
                                target = target.Replace("\\", "\")
                                'MsgBox(str & "  -  " & target)
                                RaiseEvent WorkerStatusMessage("Copying Files (" & counter & " of " & path.Count & ")")
                                Copy_FileObject(str, target)
                            Else
                                Exit For
                            End If
                        Next
                    End If
                Else
                    If path.Count > 0 Then
                        If MsgBox("Do you wish to copy the files and folders you have selected to your hand-in folder?", MsgBoxStyle.OKCancel, "Marked for Copy") = MsgBoxResult.OK Then
                            Dim counter As Integer = 0
                            For Each str As String In path
                                If CancelAction = False Then
                                    counter = counter + 1
                                    Dim target As String = str.Replace(basesourcepath, targetdir)
                                    'MsgBox(str & "  -  " & target)
                                    RaiseEvent WorkerStatusMessage("Copying Files (" & counter & " of " & path.Count & ")")
                                    Copy_FileObject(str, target)
                                Else
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            Else
                '        RaiseEvent WorkerComplete()
            End If

        Catch ex As Exception
            Error_Handler(ex, "CopyHandler")

        End Try
        RaiseEvent WorkerComplete()
    End Sub

    Private Sub Copy_FileObject(ByVal sourcepath As String, ByVal targetpath As String)
        Try


            Dim doaction As Boolean = True
            Dim target As FileInfo = New FileInfo(targetpath)
            Dim continued As Boolean = True
            Dim testdir As DirectoryInfo
            Dim arr As ArrayList = New ArrayList
            testdir = target.Directory
            While continued = True
                If testdir.Exists = True Then
                    continued = False
                Else
                    arr.Insert(0, testdir.FullName)
                    If testdir.Parent Is Nothing Then
                        continued = False
                    Else
                        testdir = testdir.Parent
                    End If
                End If
            End While
            For Each stri As String In arr
                Directory.CreateDirectory(stri)
            Next
            arr.Clear()
            arr = Nothing
            testdir = Nothing

            Dim source As FileInfo = New FileInfo(sourcepath)
            If source.Exists = False Then
                MsgBox(sourcepath.Replace(basefolder, ".\").Replace("\\", "\") & " cannot be located on your system.", MsgBoxStyle.OKOnly, "File Copy")
                doaction = False
            End If
            If target.Exists = True Then
                If Overwrite = False Then
                    If MsgBox(targetpath.Replace(basefolder, ".\").Replace("\\", "\") & " already exists on your system. Do you wish to overwrite it?", MsgBoxStyle.OKCancel, "File Copy") = MsgBoxResult.OK Then
                        If Overwrite = False Then
                            If path.Count > 1 Then
                                If MsgBox("Do you wish to overwrite all other files that may already exist?", MsgBoxStyle.OKCancel, "File Copy") = MsgBoxResult.OK Then
                                    Overwrite = True
                                End If
                            Else
                                Overwrite = True
                            End If
                        End If

                        doaction = True
                    Else
                        If path.Count > 1 Then
                            If MsgBox("Do you wish to continue processing the rest of the input files?", MsgBoxStyle.OKCancel, "File Copy") = MsgBoxResult.OK Then
                                CancelAction = False
                            Else
                                CancelAction = True
                            End If
                        Else
                            CancelAction = True
                        End If

                        doaction = False
                    End If
                Else
                    doaction = True
                End If

            End If
            If doaction = True Then
                source.CopyTo(targetpath, True)
            End If
            target = Nothing
            source = Nothing
        Catch ex As Exception
            Error_Handler(ex, "CopyFileObject")

        End Try
    End Sub

    Private Sub DeleteHandler()
        RaiseEvent WorkerStatusMessage("Deleting Files")

        Try
            Generate_RecycleBin()
            If CancelAction = False Then


                Dim fils As String = ""
                For Each str As String In path
                    If Not str = RecycleBinPath Then
                        fils = fils.Replace(basefolder, ".\").Replace("\\", "\") & "- " & str.Replace(basefolder, ".\").Replace("\\", "\") & vbCrLf
                    End If

                Next
                fils = fils.Trim

                If path.Count <= 10 And path.Count > 0 Then
                    If MsgBox("Do you wish to remove the following files and folders from your system?" & vbCrLf & vbCrLf & fils, MsgBoxStyle.OKCancel, "Marked for Deletion") = MsgBoxResult.OK Then
                        Dim counter As Integer = 0
                        For Each str As String In path
                            If CancelAction = False Then
                                counter = counter + 1
                                RaiseEvent WorkerStatusMessage("Deleting Files (" & counter & " of " & path.Count & ")")
                                Delete_FileObject(str)
                            Else
                                Exit For
                            End If
                        Next
                        '    


                    End If
                Else
                    If path.Count > 0 Then
                        If MsgBox("Do you wish to remove the files and folders you have selected from your system?", MsgBoxStyle.OKCancel, "Marked for Deletion") = MsgBoxResult.OK Then
                            Dim counter As Integer = 0
                            For Each str As String In path
                                If CancelAction = False Then
                                    counter = counter + 1
                                    RaiseEvent WorkerStatusMessage("Deleting Files (" & counter & " of " & path.Count & ")")
                                    Delete_FileObject(str)
                                Else
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "DeleteHandler")
        Finally
            RaiseEvent WorkerComplete()
        End Try
        RaiseEvent WorkerComplete()
    End Sub

    Private Sub Delete_FileObject(ByVal sourcepath As String)
        Try

            If Not sourcepath = RecycleBinPath Then
                Dim dinfo As DirectoryInfo = New DirectoryInfo(sourcepath)
                If dinfo.Exists = True Then
                    dinfo.MoveTo((RecycleBinPath & "\" & "del" & Format(Now, "yyyyMMddHHmmss") & dinfo.Name).Replace("\\", "\"))
                Else
                    Dim finfo As FileInfo = New FileInfo(sourcepath)
                    finfo.MoveTo((RecycleBinPath & "\" & "del" & Format(Now, "yyyyMMddHHmmss") & finfo.Name).Replace("\\", "\"))
                End If
            End If

        Catch ex As Exception
            Error_Handler(ex, "Delete_FileObject")

        End Try
    End Sub


    Private Sub Generate_RecycleBin()
        Try


            Dim dinfo As DirectoryInfo
            dinfo = New DirectoryInfo(RecycleBinPath)
            If dinfo.Exists = False Then
                dinfo.Create()
                dinfo.Attributes = FileAttributes.System Or FileAttributes.Hidden
            End If
            dinfo = Nothing
            If File.Exists((RecycleBinPath & "\Desktop.ini").Replace("\\", "\")) = False Then
                Dim fs As New IO.FileStream((RecycleBinPath & "\Desktop.ini").Replace("\\", "\"), IO.FileMode.Create)
                Dim tw As New IO.StreamWriter(fs)
                tw.WriteLine("[.ShellClassInfo]")
                tw.WriteLine("IconFile=%SystemRoot%\system32\SHELL32.dll")
                tw.WriteLine("IconIndex = 31")
                tw.Close()
                fs.Close()
                tw = Nothing
                fs = Nothing
                File.SetAttributes((RecycleBinPath & "\Desktop.ini").Replace("\\", "\"), FileAttributes.Hidden)
            End If
        Catch ex As Exception
            Error_Handler(ex, "Generate_RecycleBin")
        End Try

    End Sub

    Public Function MapDrive(ByVal pathtomap As String) As String
        Dim resultdrive As String = ""
        Dim result As String = ""
        Try
            Dim continue As Boolean = True
            Dim letterlist As ArrayList = New ArrayList
            letterlist.Clear()
            Dim i As Integer
            For i = 65 To 90
                letterlist.Add(Chr(i).ToString)
            Next

            Dim runner As IEnumerator
            Dim fso As New Scripting.FileSystemObject
            runner = fso.Drives.GetEnumerator()
            While runner.MoveNext() = True
                Dim d As Scripting.Drive
                d = runner.Current()
                letterlist.RemoveAt((letterlist.IndexOf(d.DriveLetter.ToString)))
            End While
            If letterlist.Count > 1 Then
                letterlist.Reverse()
            End If
            resultdrive = ""
            For Each strp As String In letterlist
                resultdrive = strp
                Exit For
            Next
            If resultdrive.Trim = "" Then
                resultdrive = "Fail. No available drive letter can be found"
                continue = False
            End If

            If continue = True Then
                Dim apptorun As String
                'net use T: \\Comlab\Vol2\handin /PERSISTENT:NO
                'net use T: /DELETE
                If pathtomap.EndsWith("\") Then
                    pathtomap = pathtomap.Remove(pathtomap.Length - 1, 1)
                End If
                apptorun = "net use " & resultdrive.Trim.ToUpper & ": """ & pathtomap & """ /PERSISTENT:NO"

                result = DosShellCommand(apptorun)
                If Not result.IndexOf("The command completed successfully.") = -1 Then
                    resultdrive = resultdrive & ":\"
                Else
                    resultdrive = "Fail. Unable to map the give path to a directory"
                End If
            End If
        Catch ex As Exception
            Error_Logger(ex, "MapDrive")
            resultdrive = "Fail. Unknown Reason. Check the log files for further information."
        End Try
        Return resultdrive
    End Function

    Public Function UnMapDrive(ByVal pathtomap As String) As String
        Dim resultdrive As String = ""
        Try
            Dim apptorun As String
            'net use T: \\Comlab\Vol2\handin /PERSISTENT:NO
            'net use T: /DELETE

            If pathtomap.EndsWith("\") = True Then
                pathtomap = pathtomap.Remove(pathtomap.Length - 1, 1)
            End If
            If pathtomap.EndsWith(":") = True Then
                pathtomap = pathtomap.Remove(pathtomap.Length - 1, 1)
            End If
            apptorun = "net use " & pathtomap & ":  /DELETE"

            resultdrive = DosShellCommand(apptorun)
        Catch ex As Exception
            Error_Logger(ex, "UnMapDrive")
            resultdrive = "Fail. Unknown Reason. Check the log files for further information."
        End Try
        Return resultdrive
    End Function

    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False


            Dim sErr As StreamReader
            Dim sOut As StreamReader
            Dim sIn As StreamWriter


            myProcess.StartInfo.CreateNoWindow = True

            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True

            'myProcess.StartInfo.FileName = AppToRun

            myProcess.Start()
            sIn = myProcess.StandardInput
            sIn.AutoFlush = True

            sOut = myProcess.StandardOutput()
            sErr = myProcess.StandardError

            sIn.Write(AppToRun & System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()

            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If



            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()


        Catch ex As Exception
            Error_Logger(ex, "DosShellCommand")
            Error_Handler(ex, "DosShellCommand")
        End Try

        Return s
    End Function

    Private Sub Error_Logger(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then

                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                If identifier_msg = "" Then
                    filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & ex.ToString)
                Else
                    filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)
                End If

                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Student Handin System's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub
End Class
