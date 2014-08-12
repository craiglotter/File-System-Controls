Imports ExpTreeLib
Imports ExpTreeLib.CShItem
Imports ExpTreeLib.SystemImageListManager
Imports System.IO
Imports System.Threading

Public Class frmExplorerLike
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
        SystemImageListManager.SetListViewImageList(lv1, True, False)
        SystemImageListManager.SetListViewImageList(lv1, False, False)

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
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents sbr1 As System.Windows.Forms.StatusBar
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewLargeIcons As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewSmallIcons As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewList As System.Windows.Forms.MenuItem
    Friend WithEvents ExpTree1 As ExpTree
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents mnuViewDetails As System.Windows.Forms.MenuItem
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents lv1 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeaderName As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderSize As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderType As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderModifyDate As System.Windows.Forms.ColumnHeader
    Friend WithEvents cb1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmdCTest As System.Windows.Forms.Button
    Friend WithEvents cmdRefresh As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.mnuExit = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.mnuViewLargeIcons = New System.Windows.Forms.MenuItem()
        Me.mnuViewSmallIcons = New System.Windows.Forms.MenuItem()
        Me.mnuViewList = New System.Windows.Forms.MenuItem()
        Me.mnuViewDetails = New System.Windows.Forms.MenuItem()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cb1 = New System.Windows.Forms.ComboBox()
        Me.lv1 = New System.Windows.Forms.ListView()
        Me.ColumnHeaderName = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeaderSize = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeaderType = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeaderModifyDate = New System.Windows.Forms.ColumnHeader()
        Me.Splitter1 = New System.Windows.Forms.Splitter()
        Me.ExpTree1 = New ExpTreeLib.ExpTree()
        Me.sbr1 = New System.Windows.Forms.StatusBar()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdCTest = New System.Windows.Forms.Button()
        Me.cmdRefresh = New System.Windows.Forms.Button()
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
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuExit})
        Me.MenuItem1.Text = "&File"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 0
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
        Me.Panel1.Size = New System.Drawing.Size(584, 328)
        Me.Panel1.TabIndex = 1
        '
        'Panel2
        '
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.cb1, Me.lv1})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(233, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(351, 328)
        Me.Panel2.TabIndex = 2
        '
        'cb1
        '
        Me.cb1.Dock = System.Windows.Forms.DockStyle.Top
        Me.cb1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb1.Name = "cb1"
        Me.cb1.Size = New System.Drawing.Size(351, 24)
        Me.cb1.TabIndex = 5
        '
        'lv1
        '
        Me.lv1.Activation = System.Windows.Forms.ItemActivation.OneClick
        Me.lv1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeaderName, Me.ColumnHeaderSize, Me.ColumnHeaderType, Me.ColumnHeaderModifyDate})
        Me.lv1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lv1.Location = New System.Drawing.Point(0, 32)
        Me.lv1.MultiSelect = False
        Me.lv1.Name = "lv1"
        Me.lv1.Size = New System.Drawing.Size(351, 296)
        Me.lv1.TabIndex = 4
        '
        'ColumnHeaderName
        '
        Me.ColumnHeaderName.Text = "Name"
        Me.ColumnHeaderName.Width = 180
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
        Me.Splitter1.Location = New System.Drawing.Point(224, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(9, 328)
        Me.Splitter1.TabIndex = 1
        Me.Splitter1.TabStop = False
        '
        'ExpTree1
        '
        Me.ExpTree1.AllowDrop = True
        Me.ExpTree1.Cursor = System.Windows.Forms.Cursors.Default
        Me.ExpTree1.Dock = System.Windows.Forms.DockStyle.Left
        Me.ExpTree1.Name = "ExpTree1"
        Me.ExpTree1.ShowRootLines = False
        Me.ExpTree1.Size = New System.Drawing.Size(224, 328)
        Me.ExpTree1.StartUpDirectory = ExpTreeLib.ExpTree.StartDir.Desktop
        Me.ExpTree1.TabIndex = 0
        '
        'sbr1
        '
        Me.sbr1.Location = New System.Drawing.Point(0, 402)
        Me.sbr1.Name = "sbr1"
        Me.sbr1.Size = New System.Drawing.Size(584, 19)
        Me.sbr1.TabIndex = 2
        Me.sbr1.Text = "Ready"
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(454, 352)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(93, 29)
        Me.cmdExit.TabIndex = 3
        Me.cmdExit.Text = "Exit"
        '
        'cmdCTest
        '
        Me.cmdCTest.Location = New System.Drawing.Point(24, 352)
        Me.cmdCTest.Name = "cmdCTest"
        Me.cmdCTest.Size = New System.Drawing.Size(93, 29)
        Me.cmdCTest.TabIndex = 4
        Me.cmdCTest.Text = "C:\ Test"
        '
        'cmdRefresh
        '
        Me.cmdRefresh.Location = New System.Drawing.Point(224, 352)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(93, 29)
        Me.cmdRefresh.TabIndex = 6
        Me.cmdRefresh.Text = "Refresh"
        '
        'frmExplorerLike
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(584, 421)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdRefresh, Me.cmdCTest, Me.cmdExit, Me.sbr1, Me.Panel1})
        Me.Menu = Me.MainMenu1
        Me.Name = "frmExplorerLike"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Explore"
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

#Region "VisibleChanged Event"
    Private Sub lv1_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lv1.VisibleChanged
        If lv1.Visible Then
            SystemImageListManager.SetListViewImageList(lv1, True, False)
            SystemImageListManager.SetListViewImageList(lv1, False, False)
        End If
    End Sub
#End Region

#Region "Form Load"
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'The next 4 lines are for testing Try & see for an ugly button
        'Dim btnIcon As Icon
        'btnIcon = SystemImageListManager.GetIcon(ExpTree1.RootItem.IconIndexNormal)
        'Dim dispImage As Image = btnIcon.ToBitmap
        'cmdCTest.BackgroundImage = dispImage
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
        SetUpComboBox(CSI)
        TotalItems = dirList.Count + fileList.Count
        Event1.WaitOne()
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
            sbr1.Text = pathName & " Has No Items"
        End If
    End Sub

#End Region

#Region "   ListView and ComboBox Event Handling"
    Private BackList As ArrayList

    Private Sub lv1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lv1.MouseUp
        Dim lvi As ListViewItem = lv1.GetItemAt(e.X, e.Y)
        If IsNothing(lvi) Then Exit Sub
        If IsNothing(lv1.SelectedItems) OrElse lv1.SelectedItems.Count < 1 Then Exit Sub
        Dim item As CShItem = lv1.SelectedItems(0).Tag
        If item.IsFolder Then
            If e.Button = MouseButtons.Right Then
                Event1.WaitOne()
                SetUpComboBox(item)
                ExpTree1.RootItem = item
            ElseIf e.Button = MouseButtons.Left Then
                ExpTree1.ExpandANode(item)
            End If
        End If
    End Sub

    Private Sub SetUpComboBox(ByVal item As CShItem)
        BackList = New ArrayList()
        With cb1
            .Items.Clear()
            .Text = ""
            Dim CSI As CShItem = item
            Do While Not IsNothing(CSI.Parent)
                CSI = CSI.Parent
                BackList.Add(CSI)
                .Items.Add(CSI.DisplayName)
            Loop
            .SelectedIndex = -1
        End With
        lv1.Focus()
    End Sub

    Private Sub cb1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cb1.SelectedIndexChanged
        With cb1
            If .SelectedIndex > -1 AndAlso _
                 .SelectedIndex < BackList.Count Then
                Dim item As CShItem = BackList(.SelectedIndex)
                BackList = New ArrayList()
                .Items.Clear()
                ExpTree1.RootItem = item
            End If
        End With
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


    Private Sub cmdCTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCTest.Click
        Dim cDir As CShItem = GetCShItem("C:\")
        If cDir.IsFolder Then
            ExpTree1.RootItem = cDir
        End If
    End Sub

    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
        If Not ExpTree1.SelectedItem Is Nothing Then
            ExpTree1.RefreshTree()
        End If
    End Sub

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

#Region "   Various testing routines. Depend on Files/Dirs on Developmental system"
    ' The Routines in this region handle buttons that have been removed from the form
    '  with the obvious names.  The routines depend on Files and Dirs found on
    '  my development system.  To see how they work, add the buttons and change 
    '  the literal references to Files & Dirs that exist on your system
    'Private Sub cmdFilterTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFilterTest.Click
    '    Dim filtAl As New ArrayList()
    '    Dim fList As New ArrayList()
    '    Dim baseItem As New CShItem("C:\Data\Checklists")
    '    fList = baseItem.GetFiles("*.doc")
    '    Dim Item As CShItem
    '    For Each Item In fList
    '        Dim xxx As New CShItem(Item.Path)
    '        Debug.WriteLine(xxx.Path)
    '        Debug.WriteLine(xxx.DisplayName)
    '        xxx.DebugDump()
    '        filtAl.Add(xxx)
    '    Next
    '    'Build the ListViewItems & add to lv1
    '    lv1.BeginUpdate()
    '    lv1.Items.Clear()
    '    For Each Item In fList
    '        Dim lvi As New ListViewItem(Item.DisplayName)
    '        With lvi
    '            If Not Item.IsDisk And Item.IsFileSystem And Not Item.IsFolder Then
    '                If Item.Length > 1024 Then
    '                    .SubItems.Add(Format(Item.Length / 1024, "#,### KB"))
    '                Else
    '                    .SubItems.Add(Format(Item.Length, "##0 Bytes"))
    '                End If
    '            Else
    '                .SubItems.Add("")
    '            End If
    '            .SubItems.Add(Item.TypeName)
    '            If Item.IsDisk Then
    '                .SubItems.Add("")
    '            Else
    '                If Item.LastWriteTime = testTime Then '"#1/1/0001 12:00:00 AM#" is empty
    '                    .SubItems.Add("")
    '                Else
    '                    .SubItems.Add(Item.LastWriteTime)
    '                End If
    '            End If
    '            .ImageIndex = SystemImageListManager.GetIconIndex(Item, False)
    '            .Tag = Item
    '        End With
    '        lv1.Items.Add(lvi)
    '    Next
    '    lv1.EndUpdate()
    'End Sub

    'Private Sub cmdExpandTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim testPath As String = "F:\Music\Clips\Brooks & Dunn\Borderline"
    '    ExpTree1.ExpandANode(testPath)
    'End Sub
#End Region

End Class
