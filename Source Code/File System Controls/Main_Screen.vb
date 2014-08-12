Imports ExpTreeLib
Imports ExpTreeLib.CShItem
Imports ExpTreeLib.SystemImageListManager
Imports System.IO
Imports System.Threading

Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False
    Private application_exit As Boolean = False
    Private filewalker As ArrayList

    Private BaseFolder As String = ""
    Private BaseTitle As String = ""
    Private RecycleBinPath As String = ""

    'avoid Globalization problem-- an empty timevalue
    Dim testTime As New DateTime(1, 1, 1, 0, 0, 0)
    Private LastSelectedCSI As CShItem
    Private Shared Event1 As New ManualResetEvent(True)

    Dim WithEvents Worker1 As Worker
    Private workerbusy As Boolean = False
    Public Delegate Sub WorkerComplete_h()
    Public Delegate Sub WorkerError_h(ByVal Message As Exception)
    Public Delegate Sub WorkerStatusMessage_h(ByVal message As String)

    '   Public Delegate Sub TRefHandler(ByVal Value As Double, ByVal _
    '     Calculations As Double)
    Public Delegate Sub TRefHandler()

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        SystemImageListManager.SetListViewImageList(lv1, True, False)
        SystemImageListManager.SetListViewImageList(lv1, False, False)
        filewalker = New ArrayList
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompletehandler
        AddHandler Worker1.WorkerErrorEncountered, AddressOf WorkerErrorHandler
        AddHandler Worker1.WorkerStatusMessage, AddressOf WorkerStatusMessageHandler
    End Sub

    Public Sub New(ByVal splash As Splash_Screen, Optional ByVal Base_Folder As String = "", Optional ByVal Base_Title As String = "")
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash
        BaseTitle = Base_Title
        BaseFolder = Base_Folder
        SystemImageListManager.SetListViewImageList(lv1, True, False)
        SystemImageListManager.SetListViewImageList(lv1, False, False)
        filewalker = New ArrayList
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompletehandler
        AddHandler Worker1.WorkerErrorEncountered, AddressOf WorkerErrorHandler
        AddHandler Worker1.WorkerStatusMessage, AddressOf WorkerStatusMessageHandler
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip

    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem


    Friend WithEvents ExpTree1 As ExpTreeLib.ExpTree
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewLargeIcons As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewSmallIcons As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewList As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewDetails As System.Windows.Forms.MenuItem
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents lv1 As System.Windows.Forms.ListView
    Friend WithEvents lv1_Name As System.Windows.Forms.ColumnHeader
    Friend WithEvents lv1_Size As System.Windows.Forms.ColumnHeader
    Friend WithEvents lv1_Type As System.Windows.Forms.ColumnHeader
    Friend WithEvents lv1_Date_Modified As System.Windows.Forms.ColumnHeader
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents ContextMenu2 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents sbr1 As System.Windows.Forms.StatusBar
    Friend WithEvents WorkLabel As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Button1 = New System.Windows.Forms.Button
        Me.lv1 = New System.Windows.Forms.ListView
        Me.lv1_Name = New System.Windows.Forms.ColumnHeader
        Me.lv1_Size = New System.Windows.Forms.ColumnHeader
        Me.lv1_Type = New System.Windows.Forms.ColumnHeader
        Me.lv1_Date_Modified = New System.Windows.Forms.ColumnHeader
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.MenuItem12 = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.ExpTree1 = New ExpTreeLib.ExpTree
        Me.ContextMenu2 = New System.Windows.Forms.ContextMenu
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem16 = New System.Windows.Forms.MenuItem
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.mnuViewLargeIcons = New System.Windows.Forms.MenuItem
        Me.mnuViewSmallIcons = New System.Windows.Forms.MenuItem
        Me.mnuViewList = New System.Windows.Forms.MenuItem
        Me.mnuViewDetails = New System.Windows.Forms.MenuItem
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.WorkLabel = New System.Windows.Forms.Label
        Me.sbr1 = New System.Windows.Forms.StatusBar
        Me.MenuItem13 = New System.Windows.Forms.MenuItem
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.lv1)
        Me.Panel1.Controls.Add(Me.Splitter1)
        Me.Panel1.Controls.Add(Me.ExpTree1)
        Me.Panel1.Location = New System.Drawing.Point(8, 52)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(856, 384)
        Me.Panel1.TabIndex = 30
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.IndianRed
        Me.Button1.Location = New System.Drawing.Point(712, 328)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(128, 40)
        Me.Button1.TabIndex = 24
        Me.Button1.Text = "Cancel"
        Me.Button1.Visible = False
        '
        'lv1
        '
        Me.lv1.AllowDrop = True
        Me.lv1.BackColor = System.Drawing.Color.White
        Me.lv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lv1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.lv1_Name, Me.lv1_Size, Me.lv1_Type, Me.lv1_Date_Modified})
        Me.lv1.ContextMenu = Me.ContextMenu1
        Me.lv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lv1.ForeColor = System.Drawing.Color.Black
        Me.lv1.Location = New System.Drawing.Point(203, 0)
        Me.lv1.Name = "lv1"
        Me.lv1.Size = New System.Drawing.Size(653, 384)
        Me.lv1.TabIndex = 23
        '
        'lv1_Name
        '
        Me.lv1_Name.Text = "Name"
        Me.lv1_Name.Width = 150
        '
        'lv1_Size
        '
        Me.lv1_Size.Text = "Size"
        '
        'lv1_Type
        '
        Me.lv1_Type.Text = "Type"
        '
        'lv1_Date_Modified
        '
        Me.lv1_Date_Modified.Text = "Date Modified"
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem8, Me.MenuItem1, Me.MenuItem13})
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 0
        Me.MenuItem8.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem9, Me.MenuItem10, Me.MenuItem12, Me.MenuItem11})
        Me.MenuItem8.Text = "View"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 0
        Me.MenuItem9.Text = "Large Icons"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 1
        Me.MenuItem10.Text = "Small Icons"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 2
        Me.MenuItem12.Text = "List"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 3
        Me.MenuItem11.Text = "Details"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.Text = "Delete"
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(200, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 384)
        Me.Splitter1.TabIndex = 22
        Me.Splitter1.TabStop = False
        '
        'ExpTree1
        '
        Me.ExpTree1.AllowDrop = True
        Me.ExpTree1.ContextMenu = Me.ContextMenu2
        Me.ExpTree1.Cursor = System.Windows.Forms.Cursors.Default
        Me.ExpTree1.Dock = System.Windows.Forms.DockStyle.Left
        Me.ExpTree1.Location = New System.Drawing.Point(0, 0)
        Me.ExpTree1.Name = "ExpTree1"
        Me.ExpTree1.ShowHiddenFolders = False
        Me.ExpTree1.ShowRootLines = False
        Me.ExpTree1.Size = New System.Drawing.Size(200, 384)
        Me.ExpTree1.StartUpDirectory = ExpTreeLib.ExpTree.StartDir.Desktop
        Me.ExpTree1.TabIndex = 21
        '
        'ContextMenu2
        '
        Me.ContextMenu2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem3, Me.MenuItem16})
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 0
        Me.MenuItem3.Text = "Refresh Tree"
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 1
        Me.MenuItem16.Text = "Delete"
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Name"
        Me.ColumnHeader1.Width = 147
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Size"
        Me.ColumnHeader2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeader2.Width = 73
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Type"
        Me.ColumnHeader3.Width = 131
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Date Modified"
        Me.ColumnHeader4.Width = 64
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = -1
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem5})
        Me.MenuItem4.Text = "File"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 0
        Me.MenuItem5.Text = "Exit"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem6, Me.MenuItem7})
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 0
        Me.MenuItem6.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.mnuExit})
        Me.MenuItem6.Text = "&File"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "Refresh Tree"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 1
        Me.mnuExit.Text = "E&xit"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 1
        Me.MenuItem7.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuViewLargeIcons, Me.mnuViewSmallIcons, Me.mnuViewList, Me.mnuViewDetails})
        Me.MenuItem7.Text = "&View"
        '
        'mnuViewLargeIcons
        '
        Me.mnuViewLargeIcons.Index = 0
        Me.mnuViewLargeIcons.Text = "Lar&ge Icons"
        '
        'mnuViewSmallIcons
        '
        Me.mnuViewSmallIcons.Index = 1
        Me.mnuViewSmallIcons.Text = "S&mall Icons"
        '
        'mnuViewList
        '
        Me.mnuViewList.Index = 2
        Me.mnuViewList.Text = "&List"
        '
        'mnuViewDetails
        '
        Me.mnuViewDetails.Index = 3
        Me.mnuViewDetails.Text = "&Details"
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(856, 36)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "The folder below is your personal hand-in folder for the project you selected. Yo" & _
        "u can drag and drop or copy and paste files from your computer to the explorer w" & _
        "indow below. Please note that due to file permission restrictions on the server " & _
        "for security purposes, not all file operations will be available to you."
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.WorkLabel)
        Me.Panel2.Controls.Add(Me.sbr1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 438)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(872, 24)
        Me.Panel2.TabIndex = 31
        '
        'WorkLabel
        '
        Me.WorkLabel.Location = New System.Drawing.Point(656, 8)
        Me.WorkLabel.Name = "WorkLabel"
        Me.WorkLabel.Size = New System.Drawing.Size(200, 16)
        Me.WorkLabel.TabIndex = 25
        Me.WorkLabel.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'sbr1
        '
        Me.sbr1.Location = New System.Drawing.Point(0, 7)
        Me.sbr1.Name = "sbr1"
        Me.sbr1.Size = New System.Drawing.Size(872, 17)
        Me.sbr1.SizingGrip = False
        Me.sbr1.TabIndex = 24
        Me.sbr1.Text = "Ready"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 2
        Me.MenuItem13.Text = "Paste"
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(872, 462)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.MinimumSize = New System.Drawing.Size(880, 496)
        Me.Name = "Main_Screen"
        Me.Text = "Main_Screen"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in File System Controls' error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Activity_Logger(ByVal identifier_msg As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg)
            filewriter.Flush()
            filewriter.Close()
        Catch ex As Exception
            Error_Handler(ex, "Activity Logger")
        End Try
    End Sub

#Region "   ExplorerTree Event Handling"
    Private Sub AfterNodeSelect(ByVal pathName As String, ByVal CSI As CShItem) Handles ExpTree1.ExpTreeNodeSelected
        Try


            Dim dirList As New ArrayList
            Dim fileList As New ArrayList
            Dim TotalItems As Integer
            LastSelectedCSI = CSI
            If CSI.DisplayName.Equals(CShItem.strMyComputer) Then
                dirList = CSI.GetDirectories 'avoid re-query since only has dirs
            Else
                dirList = CSI.GetDirectories
                fileList = CSI.GetFiles
            End If
            ' SetUpComboBox(CSI)
            TotalItems = dirList.Count + fileList.Count
            Event1.WaitOne()
            If TotalItems > 0 Then
                Dim item As CShItem
                dirList.Sort()
                fileList.Sort()
                'Me.Text = pathName
                sbr1.Text = pathName.Replace(BaseFolder, ".\").Replace("\\", "\") & "                 " & _
                        dirList.Count & " Directories " & fileList.Count & " Files"
                Dim combList As New ArrayList(TotalItems)
                combList.AddRange(dirList)
                combList.AddRange(fileList)

                'Build the ListViewItems & add to lv1
                lv1.BeginUpdate()
                lv1.Items.Clear()
                For Each item In combList
                    Dim lvi As New ListViewItem(item.DisplayName)
                    With lvi
                        If Not item.IsDisk And item.IsFileSystem And Not item.IsFolder Then
                            If item.Length > 1024 Then
                                .SubItems.Add(Format(item.Length / 1024, "#,### KB"))
                            Else
                                .SubItems.Add(Format(item.Length, "##0 Bytes"))
                            End If
                        Else
                            .SubItems.Add("")
                        End If
                        .SubItems.Add(item.TypeName)
                        If item.IsDisk Then
                            .SubItems.Add("")
                        Else
                            If item.LastWriteTime = testTime Then '"#1/1/0001 12:00:00 AM#" is empty
                                .SubItems.Add("")
                            Else
                                .SubItems.Add(item.LastWriteTime)
                            End If
                        End If
                        '.ImageIndex = SystemImageListManager.GetIconIndex(item, False)
                        .Tag = item
                    End With
                    lv1.Items.Add(lvi)
                Next
                lv1.EndUpdate()
                LoadLV1Images()
            Else
                lv1.Items.Clear()
                sbr1.Text = pathName.Replace(BaseFolder, ".\").Replace("\\", "\") & " Has No Items"
            End If
        Catch ex As Exception
            Error_Handler(ex, "AfterNodeSelect")
        End Try
    End Sub

#End Region

#Region "   ListView and ComboBox Event Handling"
    'Private BackList As ArrayList

    Private Sub lv1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lv1.MouseUp
        Try

            Dim lvi As ListViewItem = lv1.GetItemAt(e.X, e.Y)


            If IsNothing(lvi) = True Or IsNothing(lv1.SelectedItems) Or lv1.SelectedItems.Count < 1 Then
                MenuItem1.Enabled = False
                Exit Sub
            Else
                MenuItem1.Enabled = True
                Exit Sub
            End If

        Catch ex As Exception
            Error_Handler(ex, "lv1_MouseUp")
        End Try
    End Sub






#End Region








#Region "   IconIndex Loading Thread"
    Private Sub LoadLV1Images()
        'DoLoadLv()
        Try
            Dim ts As New ThreadStart(AddressOf DoLoadLv)
            Dim ot As New Thread(ts)
            ot.ApartmentState = ApartmentState.STA
            Event1.Reset()
            ot.Start()
        Catch ex As Exception
            Error_Handler(ex, "LoadLV1Images")
        End Try
    End Sub

    Private Sub DoLoadLv()
        Try
            If application_exit = False Then

                Dim lvi As ListViewItem
                For Each lvi In lv1.Items
                    lvi.ImageIndex = SystemImageListManager.GetIconIndex(lvi.Tag, False)
                Next
                Event1.Set()

            End If
        Catch ex As Exception
            Error_Handler(ex, "DoLoadLv")
        End Try
    End Sub
#End Region

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

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If BaseFolder.StartsWith("\\") Then
                BaseFolder = Worker1.MapDrive(BaseFolder)
            End If
            If BaseFolder.StartsWith("Fail") Then
                MsgBox("Unable to load BaseFolder")
                Me.Close()
            End If

            RecycleBinPath = (BaseFolder & "\_Recycle Bin").Replace("\\", "\")
            Generate_RecycleBin()
            Me.Text = BaseTitle
            Dim dinfo As DirectoryInfo = New DirectoryInfo(BaseFolder)
            If dinfo.Exists = False Then
                dinfo.Create()
            End If
            dinfo = Nothing
            Dim cs As CShItem = New CShItem(BaseFolder)
            ExpTree1.AllowDrop = False
            ExpTree1.RootItem = cs
            ExpTree1.RefreshTree(cs)


            'ft.RootFolder = BaseFolder
            'ft.Refresh()


            dataloaded = True
            splash_loader.Visible = False
        Catch ex As Exception
            Error_Handler(ex, "Load")
        End Try
    End Sub





    Private Sub mnuViewLargeIcons_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewLargeIcons.Click
        Try
            lv1.View = View.LargeIcon
        Catch ex As Exception
            Error_Handler(ex, "mnuViewLargeIcons_Click")
        End Try
    End Sub

    Private Sub mnuViewSmallIcons_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewSmallIcons.Click
        Try
            lv1.View = View.SmallIcon
        Catch ex As Exception
            Error_Handler(ex, "mnuViewSmallIcons_Click")
        End Try
    End Sub

    Private Sub mnuViewList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewList.Click
        Try
            lv1.View = View.List
        Catch ex As Exception
            Error_Handler(ex, "mnuViewList_Click")
        End Try
    End Sub

    Private Sub mnuViewDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewDetails.Click
        Try
            lv1.View = View.Details
            lv1.Columns.Item(0).Width = Math.Round(((lv1.Width - 4) / 5) * 2)
            lv1.Columns.Item(1).Width = Math.Round(((lv1.Width - 4) / 5))
            lv1.Columns.Item(2).Width = Math.Round(((lv1.Width - 4) / 5))
            lv1.Columns.Item(3).Width = Math.Round(((lv1.Width - 4) / 5))
        Catch ex As Exception
            Error_Handler(ex, "mnuViewDetails_Click")
        End Try
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex, "mnuExit_Click")
        End Try
    End Sub



    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click
        Try
            lv1.View = View.LargeIcon
        Catch ex As Exception
            Error_Handler(ex, "MenuItem9_Click")
        End Try
    End Sub

    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem10.Click
        Try
            lv1.View = View.SmallIcon
        Catch ex As Exception
            Error_Handler(ex, "MenuItem10_Click")
        End Try
    End Sub

    Private Sub MenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem12.Click
        Try
            lv1.View = View.List
        Catch ex As Exception
            Error_Handler(ex, "MenuItem12_Click")
        End Try
    End Sub

    Private Sub MenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem11.Click
        Try
            lv1.View = View.Details
            lv1.Columns.Item(0).Width = Math.Round(((lv1.Width - 4) / 5) * 2)
            lv1.Columns.Item(1).Width = Math.Round(((lv1.Width - 4) / 5))
            lv1.Columns.Item(2).Width = Math.Round(((lv1.Width - 4) / 5))
            lv1.Columns.Item(3).Width = Math.Round(((lv1.Width - 4) / 5))
        Catch ex As Exception
            Error_Handler(ex, "MenuItem11_Click")
        End Try
    End Sub


    Private Sub Main_Screen_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            application_exit = True
        Catch ex As Exception
            Error_Handler(ex, "Main_Screen_Closing")
        End Try
    End Sub


   


    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        Try
            Dim it As ListViewItem
            Dim fils As ArrayList = New ArrayList
            For Each it In lv1.SelectedItems
                If Not (ExpTree1.SelectedItem.Path & "\" & it.Text.ToString).Replace("\\", "\").StartsWith(RecycleBinPath) Then
                    fils.Add((ExpTree1.SelectedItem.Path & "\" & it.Text.ToString).Replace("\\", "\"))
                End If

            Next
            If fils.Count > 0 Then
                ForceWork(2, fils, "", "")
            End If
        Catch ex As Exception
            Error_Handler(ex, "MenuItem1_Click")
        End Try
    End Sub



    Private Sub exptree1_KeyTrapper(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ExpTree1.KeyDown
        Try

            e.Handled = False

            If e.KeyCode.Equals(System.Windows.Forms.Keys.Back) = True And e.Handled = False Then
                If Not ExpTree1.SelectedItem.Parent Is Nothing Then
                    If Not ExpTree1.SelectedItem.Parent.Parent Is Nothing Then
                        ExpTree1.SelectItem(ExpTree1.SelectedItem.Parent.Parent, ExpTree1.SelectedItem.Parent)
                    Else
                        ExpTree1.SelectItem(ExpTree1.RootItem, ExpTree1.RootItem)
                    End If
                End If
                e.Handled = True
            End If



            If e.KeyCode.Equals(System.Windows.Forms.Keys.Delete) = True And e.Handled = False Then
                Dim foldertodelete As String = ExpTree1.SelectedItem.ToString
                Dim continue As Boolean = False
                If Not ExpTree1.SelectedItem.Parent Is Nothing Then
                    If Not ExpTree1.SelectedItem.Parent.Parent Is Nothing Then
                        ExpTree1.SelectItem(ExpTree1.SelectedItem.Parent.Parent, ExpTree1.SelectedItem.Parent)
                        continue = True
                    Else
                        ExpTree1.SelectItem(ExpTree1.RootItem, ExpTree1.RootItem)
                        continue = True
                    End If
                End If
                If continue = True Then
                    Dim it As ListViewItem
                    Dim fils As ArrayList = New ArrayList
                    For Each it In lv1.Items
                        If it.Text.ToString = foldertodelete Then
                            fils.Add((ExpTree1.SelectedItem.Path & "\" & it.Text.ToString).Replace("\\", "\"))
                        End If
                    Next
                    If fils.Count > 0 Then
                        ForceWork(2, fils, "", "")
                    End If
                End If
            End If

        Catch ex As Exception
            Error_Handler(ex, "exptree1_KeyTrapper")
        End Try
    End Sub


    Private Sub lv1_KeyTrapper(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lv1.KeyDown
        Try
            e.Handled = False
            If e.KeyCode.Equals(System.Windows.Forms.Keys.Back) = True And e.Handled = False Then
                If Not ExpTree1.SelectedItem.Parent Is Nothing Then
                    If Not ExpTree1.SelectedItem.Parent.Parent Is Nothing Then
                        ExpTree1.SelectItem(ExpTree1.SelectedItem.Parent.Parent, ExpTree1.SelectedItem.Parent)
                    Else
                        ExpTree1.SelectItem(ExpTree1.RootItem, ExpTree1.RootItem)
                    End If
                End If
                e.Handled = True
            End If

            If e.KeyCode.Equals(System.Windows.Forms.Keys.Delete) = True And e.Handled = False Then
                Dim it As ListViewItem
                Dim fils As ArrayList = New ArrayList
                For Each it In lv1.SelectedItems
                    If Not (ExpTree1.SelectedItem.Path & "\" & it.Text.ToString).Replace("\\", "\").StartsWith(RecycleBinPath) Then
                        fils.Add((ExpTree1.SelectedItem.Path & "\" & it.Text.ToString).Replace("\\", "\"))
                    End If
                Next
                If fils.Count > 0 Then
                    ForceWork(2, fils, "", "")

                End If
                e.Handled = True
            End If

            If e.KeyData = (Keys.Control + Keys.V) And e.Handled = False Then
                PasteNeeded()
                e.Handled = True
            End If
        Catch ex As Exception
            Error_Handler(ex, "lv1_KeyTrapper")
        End Try
    End Sub

    Private Sub lv1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lv1.DragEnter
        Try
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                e.Effect = DragDropEffects.All
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Sub RecursiveFileWalker(ByVal path As String)
        Try
            Dim dinfo As DirectoryInfo = New DirectoryInfo(path)
            If dinfo.Exists = True Then
                For Each finfo As FileInfo In dinfo.GetFiles
                    filewalker.Add(finfo.FullName)
                Next
                For Each sdinfo As DirectoryInfo In dinfo.GetDirectories
                    RecursiveFileWalker(sdinfo.FullName)
                Next
            End If
            dinfo = Nothing
        Catch ex As Exception
            Error_Handler(ex, "RecursiveFileWalker")
        End Try
    End Sub

    Private Sub GetFilesFromFolder(ByVal path As String)
        Try
            filewalker.Clear()
            RecursiveFileWalker(path)
        Catch ex As Exception
            Error_Handler(ex, "GetFilesFromFolder")
        End Try
    End Sub

    Private Sub lv1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lv1.DragDrop
        Try

            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                Dim MyFiles() As String
                Dim i As Integer

                ' Assign the files to an array.
                MyFiles = e.Data.GetData(DataFormats.FileDrop)
                ' Loop through the array and add the files to the list.
                Dim fils As ArrayList = New ArrayList
                Dim basesourcepath As String
                For i = 0 To MyFiles.Length - 1
                    If MyFiles.Length > 0 Then
                        Dim finfo As FileInfo = New FileInfo(MyFiles(i))
                        Dim dinfo As DirectoryInfo = New DirectoryInfo(MyFiles(i))

                        If finfo.Exists = True Then
                            fils.Add(finfo.FullName)
                            basesourcepath = finfo.DirectoryName
                        Else
                            If dinfo.Exists = True Then
                                basesourcepath = dinfo.Parent.FullName
                                GetFilesFromFolder(dinfo.FullName)
                                For Each str As String In filewalker
                                    fils.Add(str)
                                Next
                            End If
                        End If
                        dinfo = Nothing
                        finfo = Nothing
                    End If
                Next
                If fils.Count > 0 Then
                    ForceWork(1, fils, basesourcepath, ExpTree1.SelectedItem.Path)
                    '                    CopyHandler(fils, basesourcepath, ExpTree1.SelectedItem.Path)
                End If

            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub



    Private Sub lv1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lv1.DoubleClick
        Try
            Dim it As ListViewItem
            If lv1.SelectedItems.Count = 1 Then
                it = lv1.SelectedItems(0)
                Dim dinfo As DirectoryInfo = New DirectoryInfo((ExpTree1.SelectedItem.Path & "\" & it.Text.ToString).Replace("\\", "\"))
                If dinfo.Exists = True Then
                    If Not ExpTree1.SelectedItem Is Nothing Then
                        If ExpTree1.SelectedItem.HasSubFolders = True Then
                            For Each dinf As CShItem In ExpTree1.SelectedItem.GetDirectories(True)
                                If (dinf.Path) = dinfo.FullName Then
                                    dinf.RefreshDirectories()
                                    If ExpTree1.SelectItem(ExpTree1.SelectedItem, dinf) = True Then
                                        ExpTree1.ExpandANode(ExpTree1.SelectedItem)
                                    End If
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                End If
                dinfo = Nothing
            End If

        Catch ex As Exception
            Error_Handler(ex, "lv1_DoubleClick")
        End Try
    End Sub



    Private Sub MenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem16.Click
        Try
            If Not ExpTree1.SelectedItem.Path = ExpTree1.RootItem.Path Then
                Dim foldertodelete As String = ExpTree1.SelectedItem.ToString

                Dim continue As Boolean = False
                If Not ExpTree1.SelectedItem.Parent Is Nothing Then
                    If Not ExpTree1.SelectedItem.Parent.Parent Is Nothing Then
                        ExpTree1.SelectItem(ExpTree1.SelectedItem.Parent.Parent, ExpTree1.SelectedItem.Parent)
                        continue = True
                    Else
                        ExpTree1.SelectItem(ExpTree1.RootItem, ExpTree1.RootItem)
                        continue = True
                    End If
                End If
                If continue = True Then
                    Dim it As ListViewItem
                    Dim fils As ArrayList = New ArrayList
                    For Each it In lv1.Items
                        If it.Text.ToString = foldertodelete Then
                            fils.Add((ExpTree1.SelectedItem.Path & "\" & it.Text.ToString).Replace("\\", "\"))
                        End If
                    Next
                    If fils.Count > 0 Then
                        ForceWork(2, fils, "", "")
                    End If
                End If
            Else
                MsgBox("You cannot delete your root hand-in folder.", MsgBoxStyle.Information, "Delete Request Denied")
            End If
        Catch ex As Exception
            Error_Handler(ex, "MenuItem16_Click")
        End Try
    End Sub


    Private Sub RefreshTree()
        Try
            If Not ExpTree1.RootItem Is Nothing Then
                ExpTree1.RefreshTree(ExpTree1.RootItem)
            End If
            If Not ExpTree1.SelectedItem Is Nothing Then
                ExpTree1.ExpandANode(ExpTree1.SelectedItem)
            End If
            ''
            lv1.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "Refresh Tree")
        End Try
    End Sub
    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        RefreshTree()
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        RefreshTree()
    End Sub


    Private Sub ApplicationExit_Handler(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            If Not Worker1 Is Nothing Then
                If BaseFolder.StartsWith("\\") Then
                    Worker1.UnMapDrive(BaseFolder)
                End If
                Worker1.Dispose()
            End If
        Catch ex As Exception
            Error_Handler(ex, "Application Exit")
        End Try
    End Sub


    Private Sub WorkerStatusMessageHandler(ByVal message As String)
        Try
            WorkLabel.Text = message
            WorkLabel.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "WorkerStatusMessageHandler")
        End Try
    End Sub
    Private Sub WorkerErrorHandler(ByVal exc As Exception, ByVal message As String)
        Try
            Error_Handler(exc, message)
        Catch ex As Exception
            Error_Handler(ex, "WorkerErrorHandler")
        End Try
    End Sub
    Private Sub WorkerCompletehandler()
        Try
            Button1.Visible = False
            workerbusy = False
            WorkLabel.Text = ""
            WorkLabel.Refresh()
            TreeRefHandler()
        Catch ex As Exception
            Error_Handler(ex, "WorkerCompletehandler")
        End Try
    End Sub



    Private Sub ForceWork(ByVal inputselection As Integer, ByVal path As ArrayList, ByVal basesourcepath As String, ByVal targetdir As String)
        Try
            If workerbusy = False Then
                Button1.Visible = True
                workerbusy = True
                Worker1.path.Clear()
                For Each str As String In path
                    Worker1.path.Add(str)
                Next
                Worker1.basesourcepath = basesourcepath
                Worker1.targetdir = targetdir
                Worker1.basefolder = BaseFolder
                Worker1.RecycleBinPath = RecycleBinPath
                Worker1.ChooseThreads(inputselection)
            Else
                MsgBox("Sorry. This application is still busy processing your previous request. Please await its completion first before submitting another request.", MsgBoxStyle.Information, "Worker Busy")
            End If
        Catch ex As Exception
            Error_Handler(ex, "ForceWork")
        End Try

    End Sub

    Protected Sub TreeRefHandler()
        ' BeginInvoke causes asynchronous execution to begin at the address
        ' specified by the delegate. Simply put, it transfers execution of 
        ' this method back to the main thread. Any parameters required by 
        ' the method contained at the delegate are wrapped in an object and 
        ' passed. 
        'Me.BeginInvoke(New TRefHandler(AddressOf FactHandler), New Object() _
        '   {Value, Calculations})
        Me.BeginInvoke(New TRefHandler(AddressOf TreeRefreshHandler))
    End Sub

    Public Sub TreeRefreshHandler()
        RefreshTree()
    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If Not Worker1 Is Nothing Then
                Worker1.CancelAction = True
                WorkerCompletehandler()
            End If
        Catch ex As Exception
            Error_Handler(ex, "Cancel Button")
        End Try
    End Sub

    Private Sub PasteNeeded()


        Try
            Dim data As IDataObject = Clipboard.GetDataObject()
            If data.GetDataPresent(DataFormats.FileDrop) Then
                Dim MyFiles() As String
                Dim i As Integer

                ' Assign the files to an array.
                MyFiles = data.GetData(DataFormats.FileDrop)
                ' Loop through the array and add the files to the list.
                Dim fils As ArrayList = New ArrayList
                Dim basesourcepath As String
                For i = 0 To MyFiles.Length - 1
                    If MyFiles.Length > 0 Then
                        Dim finfo As FileInfo = New FileInfo(MyFiles(i))
                        Dim dinfo As DirectoryInfo = New DirectoryInfo(MyFiles(i))

                        If finfo.Exists = True Then
                            fils.Add(finfo.FullName)
                            basesourcepath = finfo.DirectoryName
                        Else
                            If dinfo.Exists = True Then
                                basesourcepath = dinfo.Parent.FullName
                                GetFilesFromFolder(dinfo.FullName)
                                For Each str As String In filewalker
                                    fils.Add(str)
                                Next
                            End If
                        End If
                        dinfo = Nothing
                        finfo = Nothing
                    End If
                Next
                If fils.Count > 0 Then
                    ForceWork(1, fils, basesourcepath, ExpTree1.SelectedItem.Path)
                    '                    CopyHandler(fils, basesourcepath, ExpTree1.SelectedItem.Path)
                End If

            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try

    End Sub

    Private Sub MenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem13.Click
        PasteNeeded()
    End Sub
End Class
