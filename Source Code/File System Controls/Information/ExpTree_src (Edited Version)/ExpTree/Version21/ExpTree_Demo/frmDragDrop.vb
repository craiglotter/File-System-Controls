Imports ExpTreeLib
Imports ExpTreeLib.CShItem
Imports ExpTreeLib.SystemImageListManager
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading

Public Class frmDragDrop
    Inherits System.Windows.Forms.Form

    'avoid Globalization problem-- an empty timevalue
    Dim testTime As New DateTime(1, 1, 1, 0, 0, 0)

    Private LastSelectedCSI As CShItem

    Private Shared Event1 As New ManualResetEvent(True)

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        SystemImageListManager.SetListViewImageList(lv1, False, False)
        SystemImageListManager.SetListViewImageList(lv1, True, False)

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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ExpTree1 As ExpTreeLib.ExpTree
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents sbr1 As System.Windows.Forms.StatusBar
    Friend WithEvents lv1 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeaderName As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderSize As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderType As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderModifyDate As System.Windows.Forms.ColumnHeader
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewLargeIcons As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewSmallIcons As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewList As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewDetails As System.Windows.Forms.MenuItem
    Friend WithEvents ColumnHeaderAttributes As System.Windows.Forms.ColumnHeader
    Friend WithEvents txtDropOn As System.Windows.Forms.TextBox
    Friend WithEvents mnuChangeRoot As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRefreshTree As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSetToDesktop As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.mnuChangeRoot = New System.Windows.Forms.MenuItem()
        Me.mnuRefreshTree = New System.Windows.Forms.MenuItem()
        Me.mnuSetToDesktop = New System.Windows.Forms.MenuItem()
        Me.mnuExit = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.mnuViewLargeIcons = New System.Windows.Forms.MenuItem()
        Me.mnuViewSmallIcons = New System.Windows.Forms.MenuItem()
        Me.mnuViewList = New System.Windows.Forms.MenuItem()
        Me.mnuViewDetails = New System.Windows.Forms.MenuItem()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lv1 = New System.Windows.Forms.ListView()
        Me.ColumnHeaderName = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeaderAttributes = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeaderSize = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeaderType = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeaderModifyDate = New System.Windows.Forms.ColumnHeader()
        Me.Splitter1 = New System.Windows.Forms.Splitter()
        Me.ExpTree1 = New ExpTreeLib.ExpTree()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.sbr1 = New System.Windows.Forms.StatusBar()
        Me.txtDropOn = New System.Windows.Forms.TextBox()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuChangeRoot, Me.mnuRefreshTree, Me.mnuSetToDesktop, Me.mnuExit})
        Me.MenuItem1.Text = "&File"
        '
        'mnuChangeRoot
        '
        Me.mnuChangeRoot.Index = 0
        Me.mnuChangeRoot.Text = "&Change Root"
        '
        'mnuRefreshTree
        '
        Me.mnuRefreshTree.Index = 1
        Me.mnuRefreshTree.Text = "&RefreshTree"
        '
        'mnuSetToDesktop
        '
        Me.mnuSetToDesktop.Index = 2
        Me.mnuSetToDesktop.Text = "Set View to &Desktop"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 3
        Me.mnuExit.Text = "E&xit"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuViewLargeIcons, Me.mnuViewSmallIcons, Me.mnuViewList, Me.mnuViewDetails})
        Me.MenuItem2.Text = "&View"
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
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel2, Me.Splitter1, Me.ExpTree1})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(592, 311)
        Me.Panel1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.lv1})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(219, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(373, 311)
        Me.Panel2.TabIndex = 3
        '
        'lv1
        '
        Me.lv1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeaderName, Me.ColumnHeaderAttributes, Me.ColumnHeaderSize, Me.ColumnHeaderType, Me.ColumnHeaderModifyDate})
        Me.lv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lv1.Name = "lv1"
        Me.lv1.Size = New System.Drawing.Size(373, 311)
        Me.lv1.TabIndex = 5
        Me.lv1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeaderName
        '
        Me.ColumnHeaderName.Text = "Name"
        Me.ColumnHeaderName.Width = 180
        '
        'ColumnHeaderAttributes
        '
        Me.ColumnHeaderAttributes.Text = "Attributes"
        Me.ColumnHeaderAttributes.Width = 72
        '
        'ColumnHeaderSize
        '
        Me.ColumnHeaderSize.Text = "Size"
        Me.ColumnHeaderSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeaderSize.Width = 80
        '
        'ColumnHeaderType
        '
        Me.ColumnHeaderType.Text = "Type"
        Me.ColumnHeaderType.Width = 100
        '
        'ColumnHeaderModifyDate
        '
        Me.ColumnHeaderModifyDate.Text = "Modified"
        Me.ColumnHeaderModifyDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeaderModifyDate.Width = 80
        '
        'Splitter1
        '
        Me.Splitter1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Splitter1.Location = New System.Drawing.Point(212, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(7, 311)
        Me.Splitter1.TabIndex = 2
        Me.Splitter1.TabStop = False
        '
        'ExpTree1
        '
        Me.ExpTree1.AllowDrop = True
        Me.ExpTree1.Cursor = System.Windows.Forms.Cursors.Default
        Me.ExpTree1.Dock = System.Windows.Forms.DockStyle.Left
        Me.ExpTree1.Name = "ExpTree1"
        Me.ExpTree1.ShowRootLines = False
        Me.ExpTree1.Size = New System.Drawing.Size(212, 311)
        Me.ExpTree1.StartUpDirectory = ExpTreeLib.ExpTree.StartDir.Desktop
        Me.ExpTree1.TabIndex = 1
        '
        'cmdExit
        '
        Me.cmdExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdExit.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmdExit.Location = New System.Drawing.Point(528, 376)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(44, 22)
        Me.cmdExit.TabIndex = 1
        Me.cmdExit.Text = "Exit"
        '
        'sbr1
        '
        Me.sbr1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.sbr1.Location = New System.Drawing.Point(0, 405)
        Me.sbr1.Name = "sbr1"
        Me.sbr1.Size = New System.Drawing.Size(592, 18)
        Me.sbr1.TabIndex = 3
        Me.sbr1.Text = "Ready"
        '
        'txtDropOn
        '
        Me.txtDropOn.AllowDrop = True
        Me.txtDropOn.Location = New System.Drawing.Point(7, 326)
        Me.txtDropOn.Multiline = True
        Me.txtDropOn.Name = "txtDropOn"
        Me.txtDropOn.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtDropOn.Size = New System.Drawing.Size(388, 75)
        Me.txtDropOn.TabIndex = 5
        Me.txtDropOn.Text = ""
        '
        'frmDragDrop
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 423)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtDropOn, Me.sbr1, Me.cmdExit, Me.Panel1})
        Me.Menu = Me.MainMenu1
        Me.Name = "frmDragDrop"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Drag/Drop Demo"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Form Exit Methods"
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        mnuExit_Click(sender, e)
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub
#End Region

#Region "Form Load"
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    End Sub
#End Region

#Region "   ExplorerTree Event Handling"
    Private Sub AfterNodeSelect(ByVal pathName As String, ByVal CSI As CShItem) Handles ExpTree1.ExpTreeNodeSelected
        Dim dirList As New ArrayList()
        Dim fileList As New ArrayList()
        Dim TotalItems As Integer
        LastSelectedCSI = CSI
        If CSI.DisplayName.Equals(CShItem.strMyComputer) Then
            dirList = CSI.GetDirectories 'avoid re-query since only has dirs
        Else
            dirList = CSI.GetDirectories
            fileList = CSI.GetFiles
        End If
        Event1.WaitOne()
        TotalItems = dirList.Count + fileList.Count
        If TotalItems > 0 Then
            Dim item As CShItem
            dirList.Sort()
            fileList.Sort()
            Me.Text = pathName
            sbr1.Text = pathName & "                 " & _
                        dirList.Count & " Directories " & fileList.Count & " Files"
            Dim combList As New ArrayList(TotalItems)
            combList.AddRange(dirList)
            combList.AddRange(fileList)

            'Build the ListViewItems & add to lv1
            lv1.BeginUpdate()
            lv1.Items.Clear()
            lv1.Refresh()
            For Each item In combList
                Dim lvi As New ListViewItem(item.DisplayName)
                With lvi
                    If Not item.IsDisk And item.IsFileSystem Then
                        Dim attr As FileAttributes
                        attr = GetAttr(item.Path)
                        Dim SB As New StringBuilder()
                        If (attr And FileAttributes.System) = FileAttributes.System Then SB.Append("S")
                        If (attr And FileAttributes.Hidden) = FileAttributes.Hidden Then SB.Append("H")
                        If (attr And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then SB.Append("R")
                        If (attr And FileAttributes.Archive) = FileAttributes.Archive Then SB.Append("A")
                        .SubItems.Add(SB.ToString)
                    Else : .SubItems.Add("")
                    End If
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
            sbr1.Text = pathName & " Has No Items"
        End If
    End Sub

#End Region

#Region "   View Menu Event Handling"
    Private Sub mnuViewLargeIcons_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewLargeIcons.Click
        lv1.View = View.LargeIcon
    End Sub

    Private Sub mnuViewSmallIcons_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewSmallIcons.Click
        lv1.View = View.SmallIcon
    End Sub

    Private Sub mnuViewList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewList.Click
        lv1.View = View.List
    End Sub

    Private Sub mnuViewDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewDetails.Click
        lv1.View = View.Details
    End Sub
#End Region

#Region "   IconIndex Loading Thread"
    Private Sub LoadLV1Images()
        Dim ts As New ThreadStart(AddressOf DoLoadLv)
        Dim ot As New Thread(ts)
        ot.ApartmentState = ApartmentState.STA
        Event1.Reset()
        ot.Start()
    End Sub

    Private Sub DoLoadLv()
        Dim lvi As ListViewItem
        For Each lvi In lv1.Items
            lvi.ImageIndex = SystemImageListManager.GetIconIndex(lvi.Tag, False)
        Next
        Event1.Set()
    End Sub
#End Region

#Region "   Drag From Routines"
    Private Sub lv1_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles lv1.ItemDrag
        'Debug.WriteLine("Item Drag -- Item = " & e.Item.text)
        With lv1
            If .SelectedItems.Count > 0 Then
                Dim toDrag As New ArrayList()
                Dim lvItem As ListViewItem
                Dim strD(.SelectedItems.Count - 1) As String
                Dim i As Integer
                For Each lvItem In .SelectedItems
                    toDrag.Add(lvItem.Tag)
                    strD(i) = CType(lvItem.Tag, CShItem).Path
                    i += 1
                Next
                'NOTE: FileDrop allowing auto conversion will generate
                ' a Shell IDList Array on demand... but in some cases, the
                ' resultant PIDLs can be different from what we want, so
                ' do our own.
                Dim Dobj As New DataObject()
                Dim ms As MemoryStream
                ms = CProcDataObject.MakeShellIDArray(toDrag)
                With Dobj
                    If Not ms Is Nothing Then
                        .SetData("Shell IDList Array", True, ms)
                    End If
                    .SetData("FileDrop", True, strD)
                    .SetData(toDrag)
                End With
                Dim dEff As DragDropEffects
                If e.Button = MouseButtons.Right Then
                    dEff = DragDropEffects.Copy Or DragDropEffects.Move Or DragDropEffects.Link
                Else
                    dEff = DragDropEffects.Copy Or DragDropEffects.Move
                End If
                Dim res As DragDropEffects = .DoDragDrop(Dobj, dEff)
                'Debug.WriteLine(res)
                'Debug.WriteLine("")
            End If
        End With
    End Sub

#End Region

    Private Sub lv1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lv1.DoubleClick
        Dim csi As CShItem = lv1.SelectedItems(0).Tag
        'csi.DebugDump()
        'csi.DumpPidl(csi.PIDL)
        If csi.IsFolder Then
            ExpTree1.ExpandANode(csi)
        Else
            Try
                Process.Start(csi.Path)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OKOnly, "Error in starting application")
            End Try
        End If
    End Sub


    Private Sub txtDropOn_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtDropOn.DragEnter
        If e.Data.GetDataPresent("FileDrop", True) And _
           ((e.AllowedEffect And DragDropEffects.Copy) = DragDropEffects.Copy) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub


    Private Sub txtDropOn_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtDropOn.DragDrop
        Dim fList() As String = e.Data.GetData("FileDrop", True)
        txtDropOn.Text = ""
        Dim S As String
        For Each S In fList
            txtDropOn.Text += S & vbCrLf
        Next
        e.Effect = DragDropEffects.None
    End Sub

    Private Sub SAY(ByVal S As String)
        txtDropOn.Text += S & vbCrLf
        Debug.WriteLine(S)
    End Sub


    Private Sub mnuRefreshTree_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefreshTree.Click
        ExpTree1.RefreshTree()
    End Sub

    Private Sub mnuChangeRoot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuChangeRoot.Click
        ExpTree1.RootItem = ExpTree1.SelectedItem
    End Sub

    Private Sub mnuSetToDesktop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetToDesktop.Click
        ExpTree1.RootItem = CShItem.GetDeskTop
    End Sub

#Region "       Test Routines"
    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    If Not IsNothing(lv1.SelectedItems) Then
    '        Dim L As New ArrayList()
    '        Dim P As CShItem = lv1.SelectedItems(0).Tag.parent
    '        Dim Item As ListViewItem
    '        For Each Item In lv1.SelectedItems
    '            L.Add(Item.Tag)
    '        Next
    '        Dim CM As New ComDataObject(P, L)
    '        CM.Dispose()

    '    End If
    'End Sub

    Private Sub mnuShowSpecial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim eNames() As String = [Enum].GetNames(GetType(ShellDll.CSIDL))
        'Dim eNums() As ShellDll.CSIDL = [Enum].GetValues(GetType(ShellDll.CSIDL))
        Dim eNames() As String = [Enum].GetNames(GetType(ExpTree.StartDir))
        Dim eNums() As ShellDll.CSIDL = [Enum].GetValues(GetType(ExpTree.StartDir))
        Dim CSI As CShItem
        Dim i As Integer
        For i = 0 To eNames.Length - 1
            Debug.WriteLine("Getting Item for -- " & eNames(i))
            Try
                CSI = New CShItem(eNums(i))
                CSI.DebugDump() : CSI.DumpPidl(CSI.PIDL)
            Catch ex As Exception
                Debug.WriteLine("Error on making new CShitem")
            End Try
            Debug.WriteLine("")
        Next
    End Sub

    Private Sub mnuMakeDigest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim SR As New StreamReader("F:\DragNDropV4\LegalCSIDL.txt")
        Dim SW As New StreamWriter("F:\DragNDropV4\LegalCSIDLDigest.txt", False)
        Dim PidlOne As Boolean = False
        Do While SR.Peek <> -1
            Dim inp As String = SR.ReadLine()
            Dim tInp As String = inp.Trim
            If tInp.Length > 0 Then
                If tInp.StartsWith("Getting") OrElse _
                   tInp.StartsWith("Error") OrElse _
                   tInp.StartsWith("DisplayName") OrElse _
                   tInp.StartsWith("Path") OrElse _
                   tInp.StartsWith("IsFileSystem") OrElse _
                   tInp.StartsWith("PIDL") OrElse _
                   tInp.StartsWith("TypeName") Then
                    SW.WriteLine(inp)
                    Debug.WriteLine(inp)
                    PidlOne = False
                ElseIf tInp.StartsWith("ItemID #1") Then
                    PidlOne = True
                    SW.WriteLine(inp)
                    Debug.WriteLine(inp)
                ElseIf tInp.StartsWith("ItemID") Then
                    SW.WriteLine(inp)
                    Debug.WriteLine(inp)
                    PidlOne = False
                ElseIf PidlOne Then
                    SW.WriteLine(inp)
                    Debug.WriteLine(inp)
                End If
            End If
        Loop
        SR.Close()
        SW.Close()
    End Sub

    Private Sub mnuTestcPidl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim CSI As CShItem = CShItem.GetDeskTop
        dumpCPidl(CSI)
        CSI = New CShItem(ShellDll.CSIDL.DESKTOPDIRECTORY)
        dumpCPidl(CSI)
        CSI = New CShItem("C:\Temp")
        dumpCPidl(CSI)
        CSI = New CShItem("F:\DragNDropV4\ClipSpy\src")
        dumpCPidl(CSI)
        CSI = New CShItem(ShellDll.CSIDL.NETHOOD)
        dumpCPidl(CSI)
        Dim b(14) As Byte
        b(0) = 43
        Debug.WriteLine("An Invalid Pidl Tests" & IIf(IsValidPidl(b), " IsValid", " Is NOT Valid"))
    End Sub

    Private Sub dumpCPidl(ByVal CSI As CShItem)
        Dim cp As cPidl = CSI.clsPidl
        Dim o() As Object = cp.Decompose
        Debug.WriteLine(CSI.DisplayName)
        DumpPidl(CSI.PIDL)
        Dim b() As Byte
        Dim i As Integer = 1
        For Each b In o
            Debug.Write("cPidl Item #" & i & IIf(IsValidPidl(b), " IsValid", " Is NOT Valid"))
            DumpHex(b)
            i += 1
        Next

    End Sub

    Private Sub mnuTestFindCShItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ExpTreeLib.Tests.TestFindCShItem()
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim xxx As CShItem = GetCShItem("C:\")
        testJoinBytes(xxx)

        xxx = GetCShItem(CType(GetDeskTop.GetItems()(3), CShItem).Path)
        testJoinBytes(xxx)

        Dim yyy As CShItem = GetCShItem(ShellDll.CSIDL.APPDATA)
        xxx = CType(yyy.GetDirectories()(3), CShItem)
        testJoinBytes(xxx)
    End Sub

    Private Sub testJoinBytes(ByVal xxx As CShItem)
        Debug.WriteLine("Testing PIDL of -- " & xxx.DisplayName)
        DumpPidl(xxx.PIDL)
        Dim o() As Object = xxx.clsPidl.Decompose
        Dim b() As Byte
        Dim R() As Byte = o(0)
        Debug.WriteLine("Joining Pidls, Step 0")
        DumpHex(R)
        Dim i As Integer
        For i = 1 To o.Length - 1
            R = cPidl.JoinPidlBytes(R, o(i))
            Debug.WriteLine("Joining Pidls, Step " & i)
            DumpHex(R)
        Next
    End Sub

#End Region
End Class
