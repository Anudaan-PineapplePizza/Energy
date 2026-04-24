Option Strict On
Option Explicit On

Imports System.IO
Imports System.Collections.Generic
Imports System.Globalization

Public Class PreferencesForm
    Inherits System.Windows.Forms.Form

    Private Const SIG_COL_COLOR As Integer = 0
    Private Const SIG_COL_NAME As Integer = 1
    Private Const SIG_COL_UNIT As Integer = 2
    Private Const SIG_COL_PANEL As Integer = 3
    Private Const SIG_COL_CSV As Integer = 4
    Private Const SIG_COL_ID As Integer = 5

    Private Const AX_COL_COLOR As Integer = 0
    Private Const AX_COL_NAME As Integer = 1
    Private Const AX_COL_UNIT As Integer = 2
    Private Const AX_COL_MIN As Integer = 3
    Private Const AX_COL_MAX As Integer = 4
    Private Const AX_COL_GROUP As Integer = 5

    Private ReadOnly CI As CultureInfo = CultureInfo.InvariantCulture
    Private ReadOnly ColorBackground As Color = Color.FromArgb(10, 18, 34)
    Private ReadOnly ColorPanel As Color = Color.FromArgb(18, 28, 50)
    Private ReadOnly ColorHeader As Color = Color.FromArgb(22, 34, 58)
    Private ReadOnly ColorText As Color = Color.FromArgb(200, 215, 240)
    Private ReadOnly ColorSubtext As Color = Color.FromArgb(90, 115, 165)
    Private ReadOnly ColorRowEven As Color = Color.FromArgb(16, 26, 46)
    Private ReadOnly ColorRowOdd As Color = Color.FromArgb(20, 32, 56)
    Private ReadOnly ColorSelected As Color = Color.FromArgb(0, 80, 160)
    Private ReadOnly ColorApply As Color = Color.FromArgb(0, 120, 80)

    Private _dgvSignals As DataGridView = Nothing
    Private _dgvAxes As DataGridView = Nothing
    Private _filterPanel As Panel = Nothing
    Private _currentGroup As String = "ALL"
    Private _chkEpack As CheckBox = Nothing
    Private _txtEpackIP As TextBox = Nothing
    Private _nudEpackPort As NumericUpDown = Nothing
    Private _btnEpackConnect As Button = Nothing
    Private _lblEpackStatus As Label = Nothing
    Private _dgvEpack As DataGridView = Nothing
    Private _lblEpackNoSignals As Label = Nothing
    Private _comboAxY1 As ComboBox = Nothing
    Private _comboAxY2 As ComboBox = Nothing
    Private _trackBar As TrackBar = Nothing
    Private _lblInterval As Label = Nothing
    Private _nudXWindow As NumericUpDown = Nothing
    Private _lblXInfo As Label = Nothing
    Private _hdGroupChecks(9) As CheckBox
    Private _hdOrderChecks(29) As CheckBox
    Private _chkHdCsv As CheckBox

    Private ReadOnly HdGroupNames() As String = {
        "HDI1 - Current L1", "HDI2 - Current L2", "HDI3 - Current L3", "HDIN - Neutral",
        "HDU12 - Voltage 1-2", "HDU23 - Voltage 2-3", "HDU31 - Voltage 3-1",
        "HDU1N - Voltage L1-N", "HDU2N - Voltage L2-N", "HDU3N - Voltage L3-N"
    }
    Private ReadOnly HdGroupColors() As Color = {
        Color.FromArgb(0, 220, 140), Color.FromArgb(0, 190, 120),
        Color.FromArgb(0, 160, 100), Color.FromArgb(0, 250, 170),
        Color.FromArgb(80, 160, 255), Color.FromArgb(50, 130, 235),
        Color.FromArgb(30, 100, 215), Color.FromArgb(110, 185, 255),
        Color.FromArgb(140, 210, 255), Color.FromArgb(170, 230, 255)
    }

    Private ReadOnly _onApply As Action

    Public Sub New(onApplyCallback As Action)
        Me.InitializeComponent()
        _onApply = onApplyCallback
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Preferences"
        Me.Size = New Size(760, 620)
        Me.MinimumSize = New Size(700, 540)
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.BackColor = ColorBackground
        Me.ForeColor = ColorText
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Font = New Font("Segoe UI", 8.5)
        BuildUI()
    End Sub

    Private Sub BuildUI()
        Dim pBtns As New Panel()
        pBtns.Dock = DockStyle.Bottom
        pBtns.Height = 50
        pBtns.BackColor = Color.FromArgb(14, 22, 40)
        pBtns.Padding = New Padding(10, 0, 10, 0)

        Dim tbl As New TableLayoutPanel()
        tbl.Dock = DockStyle.Fill
        tbl.ColumnCount = 4
        tbl.RowCount = 1
        tbl.BackColor = Color.Transparent
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25))

        Dim btnApply As New Button()
        btnApply.Text = "Apply & Close"
        btnApply.Dock = DockStyle.Fill
        btnApply.Margin = New Padding(4, 8, 4, 8)
        btnApply.BackColor = ColorApply
        btnApply.ForeColor = Color.White
        btnApply.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        btnApply.FlatStyle = FlatStyle.Flat
        btnApply.FlatAppearance.BorderSize = 0
        AddHandler btnApply.Click, AddressOf BtnApply_Click

        Dim btnReset As New Button()
        btnReset.Text = "Reset defaults"
        btnReset.Dock = DockStyle.Fill
        btnReset.Margin = New Padding(4, 8, 4, 8)
        btnReset.BackColor = Color.FromArgb(100, 40, 30)
        btnReset.ForeColor = Color.White
        btnReset.FlatStyle = FlatStyle.Flat
        btnReset.FlatAppearance.BorderSize = 0
        AddHandler btnReset.Click, AddressOf BtnReset_Click

        Dim btnSaveCfg As New Button()
        btnSaveCfg.Text = "Save config..."
        btnSaveCfg.Dock = DockStyle.Fill
        btnSaveCfg.Margin = New Padding(4, 8, 4, 8)
        btnSaveCfg.BackColor = Color.FromArgb(38, 60, 95)
        btnSaveCfg.ForeColor = ColorText
        btnSaveCfg.FlatStyle = FlatStyle.Flat
        btnSaveCfg.FlatAppearance.BorderSize = 0
        AddHandler btnSaveCfg.Click, AddressOf btnSaveCfg_Click

        Dim btnLoadCfg As New Button()
        btnLoadCfg.Text = "Load config..."
        btnLoadCfg.Dock = DockStyle.Fill
        btnLoadCfg.Margin = New Padding(4, 8, 4, 8)
        btnLoadCfg.BackColor = Color.FromArgb(38, 60, 95)
        btnLoadCfg.ForeColor = ColorText
        btnLoadCfg.FlatStyle = FlatStyle.Flat
        btnLoadCfg.FlatAppearance.BorderSize = 0
        AddHandler btnLoadCfg.Click, AddressOf btnLoadCfg_Click

        tbl.Controls.Add(btnApply, 0, 0)
        tbl.Controls.Add(btnReset, 1, 0)
        tbl.Controls.Add(btnSaveCfg, 2, 0)
        tbl.Controls.Add(btnLoadCfg, 3, 0)
        pBtns.Controls.Add(tbl)

        Dim tabs As New TabControl()
        tabs.Dock = DockStyle.Fill
        tabs.BackColor = ColorBackground
        tabs.Font = New Font("Segoe UI", 9, FontStyle.Regular)
        tabs.DrawMode = TabDrawMode.OwnerDrawFixed
        tabs.ItemSize = New Size(120, 28)
        AddHandler tabs.DrawItem, Sub(s As Object, ev As DrawItemEventArgs)
                                      Dim tp As TabPage = tabs.TabPages(ev.Index)
                                      Dim isSelected As Boolean = (tabs.SelectedIndex = ev.Index)
                                      Dim bgColor As Color = If(isSelected, Color.FromArgb(0, 90, 160), Color.FromArgb(22, 34, 58))
                                      Using br As New SolidBrush(bgColor)
                                          ev.Graphics.FillRectangle(br, ev.Bounds)
                                      End Using
                                      Dim txtColor As Color = If(isSelected, Color.White, Color.FromArgb(160, 185, 220))
                                      Dim fnt As Font = If(isSelected, New Font("Segoe UI", 9, FontStyle.Bold), New Font("Segoe UI", 9, FontStyle.Regular))
                                      Dim sf As New StringFormat()
                                      sf.Alignment = StringAlignment.Center
                                      sf.LineAlignment = StringAlignment.Center
                                      ev.Graphics.DrawString(tp.Text, fnt, New SolidBrush(txtColor), ev.Bounds, sf)
                                  End Sub

        Dim tabSignals As New TabPage("53U Signals") : tabSignals.BackColor = ColorBackground
        Dim tabAxes As New TabPage("Axes") : tabAxes.BackColor = ColorBackground
        Dim tabHD As New TabPage("Harmonics (HD)") : tabHD.BackColor = ColorBackground
        Dim tabEpack As New TabPage("ePack") : tabEpack.BackColor = ColorBackground

        BuildTabSignals(tabSignals)
        BuildTabAxes(tabAxes)
        BuildTabHD(tabHD)
        BuildTabEpack(tabEpack)

        tabs.TabPages.Add(tabSignals)
        tabs.TabPages.Add(tabAxes)
        tabs.TabPages.Add(tabHD)
        tabs.TabPages.Add(tabEpack)

        Me.Controls.Add(tabs)
        Me.Controls.Add(pBtns)

        LoadFromPrefs()
    End Sub

    Private Sub BuildTabSignals(tab As TabPage)
        Dim pSample As New Panel()
        pSample.BackColor = ColorPanel
        pSample.Dock = DockStyle.Top
        pSample.Height = 46

        Dim lblSTitle As New Label()
        lblSTitle.Text = "SAMPLE INTERVAL"
        lblSTitle.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        lblSTitle.ForeColor = ColorSubtext
        lblSTitle.AutoSize = True
        lblSTitle.Location = New Point(12, 4)
        pSample.Controls.Add(lblSTitle)

        _trackBar = New TrackBar()
        _trackBar.Minimum = 1
        _trackBar.Maximum = 50
        _trackBar.Value = 5
        _trackBar.TickStyle = TickStyle.None
        _trackBar.Location = New Point(10, 18)
        _trackBar.Size = New Size(200, 22)
        AddHandler _trackBar.Scroll, AddressOf TrackBar_Scroll
        pSample.Controls.Add(_trackBar)

        _lblInterval = New Label()
        _lblInterval.Text = "500 ms"
        _lblInterval.ForeColor = Color.FromArgb(0, 210, 180)
        _lblInterval.Font = New Font("Consolas", 9, FontStyle.Bold)
        _lblInterval.AutoSize = True
        _lblInterval.Location = New Point(216, 22)
        pSample.Controls.Add(_lblInterval)

        Dim pFilter As New Panel()
        pFilter.BackColor = ColorPanel
        pFilter.Dock = DockStyle.Top
        pFilter.Height = 36
        _filterPanel = pFilter

        Dim x As Integer = 8
        For Each g As String In New String() {"ALL", "Current", "Voltage", "Power", "THD", "Phase"}
            Dim btn As New Button()
            btn.Text = g
            btn.Size = New Size(70, 22)
            btn.Location = New Point(x, 7)
            btn.Tag = g
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 0
            btn.BackColor = If(g = "ALL", ColorApply, Color.FromArgb(38, 54, 84))
            btn.ForeColor = Color.White
            btn.Font = New Font("Segoe UI", 7.5, FontStyle.Bold)
            btn.Cursor = Cursors.Hand
            AddHandler btn.Click, AddressOf FilterBtn_Click
            pFilter.Controls.Add(btn)
            x += 76
        Next

        Dim pSelAll As New Panel()
        pSelAll.BackColor = ColorPanel
        pSelAll.Dock = DockStyle.Top
        pSelAll.Height = 30

        Dim lblSel As New Label()
        lblSel.Text = "Panel :"
        lblSel.ForeColor = ColorSubtext
        lblSel.Font = New Font("Segoe UI", 7.5, FontStyle.Bold)
        lblSel.AutoSize = True
        lblSel.Location = New Point(10, 8)
        pSelAll.Controls.Add(lblSel)

        Dim btnPanelAll As New Button()
        btnPanelAll.Text = "All"
        btnPanelAll.Size = New Size(44, 18)
        btnPanelAll.Location = New Point(55, 6)
        btnPanelAll.FlatStyle = FlatStyle.Flat
        btnPanelAll.FlatAppearance.BorderSize = 0
        btnPanelAll.BackColor = Color.FromArgb(38, 60, 95)
        btnPanelAll.ForeColor = ColorText
        btnPanelAll.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        AddHandler btnPanelAll.Click, Sub(s, ev) SetSignalsColumn(SIG_COL_PANEL, True)
        pSelAll.Controls.Add(btnPanelAll)

        Dim btnPanelNone As New Button()
        btnPanelNone.Text = "None"
        btnPanelNone.Size = New Size(44, 18)
        btnPanelNone.Location = New Point(103, 6)
        btnPanelNone.FlatStyle = FlatStyle.Flat
        btnPanelNone.FlatAppearance.BorderSize = 0
        btnPanelNone.BackColor = Color.FromArgb(38, 60, 95)
        btnPanelNone.ForeColor = ColorText
        btnPanelNone.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        AddHandler btnPanelNone.Click, Sub(s, ev) SetSignalsColumn(SIG_COL_PANEL, False)
        pSelAll.Controls.Add(btnPanelNone)

        Dim lblCsvSel As New Label()
        lblCsvSel.Text = "CSV :"
        lblCsvSel.ForeColor = ColorSubtext
        lblCsvSel.Font = New Font("Segoe UI", 7.5, FontStyle.Bold)
        lblCsvSel.AutoSize = True
        lblCsvSel.Location = New Point(168, 8)
        pSelAll.Controls.Add(lblCsvSel)

        Dim btnCsvAll As New Button()
        btnCsvAll.Text = "All"
        btnCsvAll.Size = New Size(44, 18)
        btnCsvAll.Location = New Point(207, 6)
        btnCsvAll.FlatStyle = FlatStyle.Flat
        btnCsvAll.FlatAppearance.BorderSize = 0
        btnCsvAll.BackColor = Color.FromArgb(38, 60, 95)
        btnCsvAll.ForeColor = ColorText
        btnCsvAll.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        AddHandler btnCsvAll.Click, Sub(s, ev) SetSignalsColumn(SIG_COL_CSV, True)
        pSelAll.Controls.Add(btnCsvAll)

        Dim btnCsvNone As New Button()
        btnCsvNone.Text = "None"
        btnCsvNone.Size = New Size(44, 18)
        btnCsvNone.Location = New Point(255, 6)
        btnCsvNone.FlatStyle = FlatStyle.Flat
        btnCsvNone.FlatAppearance.BorderSize = 0
        btnCsvNone.BackColor = Color.FromArgb(38, 60, 95)
        btnCsvNone.ForeColor = ColorText
        btnCsvNone.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        AddHandler btnCsvNone.Click, Sub(s, ev) SetSignalsColumn(SIG_COL_CSV, False)
        pSelAll.Controls.Add(btnCsvNone)

        Dim pGrid As New Panel()
        pGrid.BackColor = ColorBackground
        pGrid.Dock = DockStyle.Fill

        _dgvSignals = NewDGV()
        AppendColorCol(_dgvSignals)
        AppendTextCol(_dgvSignals, "SCol_Name", "Signal", 150, False)
        AppendTextCol(_dgvSignals, "SCol_Unit", "Unit", 50, True)

        Dim cPanel As New DataGridViewCheckBoxColumn()
        cPanel.HeaderText = "Panel"
        cPanel.Name = "SCol_Panel"
        cPanel.Width = 55
        cPanel.SortMode = DataGridViewColumnSortMode.NotSortable
        _dgvSignals.Columns.Add(cPanel)

        Dim cCsv As New DataGridViewCheckBoxColumn()
        cCsv.HeaderText = "CSV"
        cCsv.Name = "SCol_CSV"
        cCsv.Width = 55
        cCsv.SortMode = DataGridViewColumnSortMode.NotSortable
        _dgvSignals.Columns.Add(cCsv)

        AppendHiddenCol(_dgvSignals, "SCol_SID")
        AddHandler _dgvSignals.CellPainting, AddressOf DgvSignals_CellPainting
        _dgvSignals.Dock = DockStyle.Fill
        pGrid.Controls.Add(_dgvSignals)

        tab.Controls.Add(pGrid)
        tab.Controls.Add(pSelAll)
        tab.Controls.Add(pFilter)
        tab.Controls.Add(pSample)
    End Sub

    Private Sub BuildTabAxes(tab As TabPage)
        Dim pY As New Panel()
        pY.BackColor = ColorPanel
        pY.Dock = DockStyle.Top
        pY.Height = 52

        Dim lblYTitle As New Label()
        lblYTitle.Text = "Y AXIS ASSIGNMENT"
        lblYTitle.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        lblYTitle.ForeColor = ColorSubtext
        lblYTitle.AutoSize = True
        lblYTitle.Location = New Point(12, 4)
        pY.Controls.Add(lblYTitle)

        Dim lblY1 As New Label()
        lblY1.Text = "Y1 ="
        lblY1.Font = New Font("Segoe UI", 8.5, FontStyle.Bold)
        lblY1.ForeColor = Color.FromArgb(100, 170, 255)
        lblY1.AutoSize = True
        lblY1.Location = New Point(12, 24)
        pY.Controls.Add(lblY1)

        _comboAxY1 = New ComboBox()
        _comboAxY1.DropDownStyle = ComboBoxStyle.DropDownList
        _comboAxY1.BackColor = Color.FromArgb(16, 26, 46)
        _comboAxY1.ForeColor = Color.FromArgb(200, 220, 255)
        _comboAxY1.Font = New Font("Segoe UI", 8.5)
        _comboAxY1.Location = New Point(48, 20)
        _comboAxY1.Size = New Size(190, 24)
        _comboAxY1.FlatStyle = FlatStyle.Flat
        pY.Controls.Add(_comboAxY1)

        Dim lblY2 As New Label()
        lblY2.Text = "Y2 ="
        lblY2.Font = New Font("Segoe UI", 8.5, FontStyle.Bold)
        lblY2.ForeColor = Color.FromArgb(100, 220, 140)
        lblY2.AutoSize = True
        lblY2.Location = New Point(260, 24)
        pY.Controls.Add(lblY2)

        _comboAxY2 = New ComboBox()
        _comboAxY2.DropDownStyle = ComboBoxStyle.DropDownList
        _comboAxY2.BackColor = Color.FromArgb(16, 26, 46)
        _comboAxY2.ForeColor = Color.FromArgb(200, 220, 255)
        _comboAxY2.Font = New Font("Segoe UI", 8.5)
        _comboAxY2.Location = New Point(296, 20)
        _comboAxY2.Size = New Size(190, 24)
        _comboAxY2.FlatStyle = FlatStyle.Flat
        pY.Controls.Add(_comboAxY2)

        For Each g As UserPreferences.AxisGroupDef In UserPreferences.AxisGroups
            _comboAxY1.Items.Add(g.Name)
            _comboAxY2.Items.Add(g.Name)
        Next

        Dim pXRow As New Panel()
        pXRow.BackColor = ColorPanel
        pXRow.Dock = DockStyle.Top
        pXRow.Height = 46

        Dim lblXTitle As New Label()
        lblXTitle.Text = "X AXIS - TIME WINDOW"
        lblXTitle.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        lblXTitle.ForeColor = ColorSubtext
        lblXTitle.AutoSize = True
        lblXTitle.Location = New Point(12, 4)
        pXRow.Controls.Add(lblXTitle)

        Dim lblXLbl As New Label()
        lblXLbl.Text = "Window:"
        lblXLbl.ForeColor = ColorText
        lblXLbl.AutoSize = True
        lblXLbl.Location = New Point(12, 24)
        pXRow.Controls.Add(lblXLbl)

        _nudXWindow = New NumericUpDown()
        _nudXWindow.Minimum = 10
        _nudXWindow.Maximum = 3600
        _nudXWindow.Increment = 30
        _nudXWindow.Value = 180
        _nudXWindow.BackColor = Color.FromArgb(16, 26, 46)
        _nudXWindow.ForeColor = Color.FromArgb(0, 210, 180)
        _nudXWindow.Font = New Font("Consolas", 9, FontStyle.Bold)
        _nudXWindow.Location = New Point(80, 20)
        _nudXWindow.Size = New Size(90, 22)
        _nudXWindow.BorderStyle = BorderStyle.FixedSingle
        AddHandler _nudXWindow.ValueChanged, AddressOf NudXWindow_Changed
        pXRow.Controls.Add(_nudXWindow)

        _lblXInfo = New Label()
        _lblXInfo.ForeColor = ColorSubtext
        _lblXInfo.AutoSize = True
        _lblXInfo.Location = New Point(180, 24)
        _lblXInfo.Text = "seconds"
        pXRow.Controls.Add(_lblXInfo)

        Dim pInfoRow As New Panel()
        pInfoRow.BackColor = ColorPanel
        pInfoRow.Dock = DockStyle.Top
        pInfoRow.Height = 26

        Dim lblInfo As New Label()
        lblInfo.Text = "Y AXIS - Min / Max per measurement type  (applies to all signals in the group)"
        lblInfo.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        lblInfo.ForeColor = ColorSubtext
        lblInfo.AutoSize = True
        lblInfo.Location = New Point(12, 7)
        pInfoRow.Controls.Add(lblInfo)

        Dim pGrid As New Panel()
        pGrid.BackColor = ColorBackground
        pGrid.Dock = DockStyle.Fill

        _dgvAxes = NewDGV()
        AppendColorCol(_dgvAxes)
        AppendTextCol(_dgvAxes, "ACol_Name", "Measurement type", 195, False)
        AppendTextCol(_dgvAxes, "ACol_Unit", "Unit", 55, True)

        Dim cMin As New DataGridViewTextBoxColumn()
        cMin.HeaderText = "Min" : cMin.Name = "ACol_Min" : cMin.Width = 105
        cMin.SortMode = DataGridViewColumnSortMode.NotSortable
        Dim sMin As New DataGridViewCellStyle()
        sMin.Alignment = DataGridViewContentAlignment.MiddleRight
        sMin.ForeColor = Color.FromArgb(255, 160, 80)
        cMin.DefaultCellStyle = sMin
        _dgvAxes.Columns.Add(cMin)

        Dim cMax As New DataGridViewTextBoxColumn()
        cMax.HeaderText = "Max" : cMax.Name = "ACol_Max" : cMax.Width = 105
        cMax.SortMode = DataGridViewColumnSortMode.NotSortable
        Dim sMax As New DataGridViewCellStyle()
        sMax.Alignment = DataGridViewContentAlignment.MiddleRight
        sMax.ForeColor = Color.FromArgb(80, 200, 120)
        cMax.DefaultCellStyle = sMax
        _dgvAxes.Columns.Add(cMax)

        AppendHiddenCol(_dgvAxes, "ACol_GRP")
        AddHandler _dgvAxes.CellPainting, AddressOf DgvAxes_CellPainting
        _dgvAxes.Dock = DockStyle.Fill
        pGrid.Controls.Add(_dgvAxes)

        tab.Controls.Add(pGrid)
        tab.Controls.Add(pInfoRow)
        tab.Controls.Add(pXRow)
        tab.Controls.Add(pY)
    End Sub

    Private Sub BuildTabHD(tab As TabPage)
        Dim pOpts As New Panel()
        pOpts.BackColor = ColorPanel
        pOpts.Dock = DockStyle.Top
        pOpts.Height = 36

        _chkHdCsv = New CheckBox()
        _chkHdCsv.Text = "Export HD data in Excel recording"
        _chkHdCsv.ForeColor = ColorText
        _chkHdCsv.Font = New Font("Segoe UI", 8.5)
        _chkHdCsv.AutoSize = True
        _chkHdCsv.Location = New Point(14, 9)
        _chkHdCsv.BackColor = Color.Transparent
        pOpts.Controls.Add(_chkHdCsv)

        Dim pSplit As New Panel()
        pSplit.BackColor = ColorBackground
        pSplit.Dock = DockStyle.Fill

        Dim pGroups As New Panel()
        pGroups.BackColor = ColorPanel
        pGroups.Location = New Point(0, 0)
        pGroups.Size = New Size(230, 420)
        pGroups.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Bottom

        Dim pGHdr As New Panel()
        pGHdr.BackColor = Color.FromArgb(32, 46, 72)
        pGHdr.Dock = DockStyle.Top
        pGHdr.Height = 30

        Dim lblGT As New Label()
        lblGT.Text = "GROUPS  (read + display)"
        lblGT.Font = New Font("Segoe UI", 7.5, FontStyle.Bold)
        lblGT.ForeColor = ColorSubtext
        lblGT.AutoSize = True
        lblGT.Location = New Point(10, 8)
        pGHdr.Controls.Add(lblGT)

        Dim btnAllG As New Button()
        btnAllG.Text = "All" : btnAllG.Size = New Size(38, 18)
        btnAllG.Location = New Point(138, 6)
        btnAllG.FlatStyle = FlatStyle.Flat
        btnAllG.FlatAppearance.BorderSize = 0
        btnAllG.BackColor = Color.FromArgb(38, 60, 95)
        btnAllG.ForeColor = ColorText
        btnAllG.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        AddHandler btnAllG.Click, Sub(s, ev) SetAllGroups(True)
        pGHdr.Controls.Add(btnAllG)

        Dim btnNoneG As New Button()
        btnNoneG.Text = "None" : btnNoneG.Size = New Size(40, 18)
        btnNoneG.Location = New Point(180, 6)
        btnNoneG.FlatStyle = FlatStyle.Flat
        btnNoneG.FlatAppearance.BorderSize = 0
        btnNoneG.BackColor = Color.FromArgb(38, 60, 95)
        btnNoneG.ForeColor = ColorText
        btnNoneG.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        AddHandler btnNoneG.Click, Sub(s, ev) SetAllGroups(False)
        pGHdr.Controls.Add(btnNoneG)
        pGroups.Controls.Add(pGHdr)

        Dim yG As Integer = 36
        For i As Integer = 0 To 9
            If i = 4 Then
                Dim sep As New Panel()
                sep.BackColor = Color.FromArgb(40, 55, 80)
                sep.Location = New Point(8, yG - 4)
                sep.Size = New Size(210, 1)
                pGroups.Controls.Add(sep)
            End If
            Dim chk As New CheckBox()
            chk.Text = HdGroupNames(i)
            chk.ForeColor = HdGroupColors(i)
            chk.Font = New Font("Segoe UI", 8.5)
            chk.AutoSize = True
            chk.Location = New Point(12, yG)
            chk.BackColor = Color.Transparent
            _hdGroupChecks(i) = chk
            pGroups.Controls.Add(chk)
            yG += 28
        Next

        Dim pOrders As New Panel()
        pOrders.BackColor = ColorPanel
        pOrders.Location = New Point(238, 0)
        pOrders.Size = New Size(460, 420)
        pOrders.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom

        Dim pOHdr As New Panel()
        pOHdr.BackColor = Color.FromArgb(32, 46, 72)
        pOHdr.Dock = DockStyle.Top
        pOHdr.Height = 30

        Dim lblOT As New Label()
        lblOT.Text = "ORDERS  (columns displayed in Harmonics tab)"
        lblOT.Font = New Font("Segoe UI", 7.5, FontStyle.Bold)
        lblOT.ForeColor = ColorSubtext
        lblOT.AutoSize = True
        lblOT.Location = New Point(10, 8)
        pOHdr.Controls.Add(lblOT)

        Dim btnAllO As New Button()
        btnAllO.Text = "All" : btnAllO.Size = New Size(38, 18)
        btnAllO.Location = New Point(368, 6)
        btnAllO.FlatStyle = FlatStyle.Flat
        btnAllO.FlatAppearance.BorderSize = 0
        btnAllO.BackColor = Color.FromArgb(38, 60, 95)
        btnAllO.ForeColor = ColorText
        btnAllO.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        AddHandler btnAllO.Click, Sub(s, ev) SetAllOrders(True)
        pOHdr.Controls.Add(btnAllO)

        Dim btnNoneO As New Button()
        btnNoneO.Text = "None" : btnNoneO.Size = New Size(40, 18)
        btnNoneO.Location = New Point(410, 6)
        btnNoneO.FlatStyle = FlatStyle.Flat
        btnNoneO.FlatAppearance.BorderSize = 0
        btnNoneO.BackColor = Color.FromArgb(38, 60, 95)
        btnNoneO.ForeColor = ColorText
        btnNoneO.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        AddHandler btnNoneO.Click, Sub(s, ev) SetAllOrders(False)
        pOHdr.Controls.Add(btnNoneO)
        pOrders.Controls.Add(pOHdr)

        Dim lblNote As New Label()
        lblNote.Text = "All orders are always displayed in the Harmonics tab. " &
                            "These checkboxes only control CSV export."
        lblNote.ForeColor = Color.FromArgb(80, 105, 140)
        lblNote.Font = New Font("Segoe UI", 7.5, FontStyle.Italic)
        lblNote.Location = New Point(10, 35)
        lblNote.Size = New Size(440, 28)
        pOrders.Controls.Add(lblNote)

        Dim colW As Integer = 72 : Dim rowH As Integer = 24
        For idx As Integer = 0 To 29
            Dim order As Integer = idx + 2
            Dim chk As New CheckBox()
            chk.Text = "H" & order.ToString()
            chk.ForeColor = If(order <= ModbusRsMaster.HD_HARMONICS_COUNT + 1, ColorText, Color.FromArgb(80, 95, 120))
            chk.Font = New Font("Segoe UI", 8)
            chk.Size = New Size(colW, rowH)
            chk.Location = New Point(10 + (idx Mod 6) * colW, 68 + (idx \ 6) * rowH)
            chk.BackColor = Color.Transparent
            _hdOrderChecks(idx) = chk
            pOrders.Controls.Add(chk)
        Next

        pSplit.Controls.Add(pOrders)
        pSplit.Controls.Add(pGroups)

        tab.Controls.Add(pSplit)
        tab.Controls.Add(pOpts)
    End Sub

    Private Sub BuildTabEpack(tab As TabPage)
        Dim pConn As New Panel()
        pConn.BackColor = ColorPanel
        pConn.Dock = DockStyle.Top
        pConn.Height = 80

        Dim lblTitle As New Label()
        lblTitle.Text = "ePACK TCP CONNECTION"
        lblTitle.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        lblTitle.ForeColor = ColorSubtext
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(12, 6)
        pConn.Controls.Add(lblTitle)

        _chkEpack = New CheckBox()
        _chkEpack.Text = "Enable ePack TCP"
        _chkEpack.ForeColor = ColorText
        _chkEpack.Font = New Font("Segoe UI", 8.5)
        _chkEpack.AutoSize = True
        _chkEpack.Location = New Point(12, 24)
        _chkEpack.BackColor = Color.Transparent
        pConn.Controls.Add(_chkEpack)

        Dim lblIP As New Label()
        lblIP.Text = "IP :"
        lblIP.ForeColor = ColorSubtext
        lblIP.AutoSize = True
        lblIP.Location = New Point(160, 26)
        pConn.Controls.Add(lblIP)

        _txtEpackIP = New TextBox()
        _txtEpackIP.BackColor = Color.FromArgb(16, 26, 46)
        _txtEpackIP.ForeColor = Color.FromArgb(0, 210, 180)
        _txtEpackIP.Font = New Font("Consolas", 8.5)
        _txtEpackIP.Location = New Point(185, 22)
        _txtEpackIP.Size = New Size(130, 22)
        _txtEpackIP.BorderStyle = BorderStyle.FixedSingle
        pConn.Controls.Add(_txtEpackIP)

        Dim lblPort As New Label()
        lblPort.Text = "Port :"
        lblPort.ForeColor = ColorSubtext
        lblPort.AutoSize = True
        lblPort.Location = New Point(326, 26)
        pConn.Controls.Add(lblPort)

        _nudEpackPort = New NumericUpDown()
        _nudEpackPort.Minimum = 1
        _nudEpackPort.Maximum = 65535
        _nudEpackPort.Value = 502
        _nudEpackPort.BackColor = Color.FromArgb(16, 26, 46)
        _nudEpackPort.ForeColor = Color.FromArgb(0, 210, 180)
        _nudEpackPort.Font = New Font("Consolas", 8.5)
        _nudEpackPort.Location = New Point(362, 22)
        _nudEpackPort.Size = New Size(72, 22)
        _nudEpackPort.BorderStyle = BorderStyle.FixedSingle
        pConn.Controls.Add(_nudEpackPort)

        _btnEpackConnect = New Button()
        _btnEpackConnect.Text = "Connect"
        _btnEpackConnect.Size = New Size(88, 22)
        _btnEpackConnect.Location = New Point(444, 22)
        _btnEpackConnect.FlatStyle = FlatStyle.Flat
        _btnEpackConnect.FlatAppearance.BorderSize = 0
        _btnEpackConnect.BackColor = Color.FromArgb(38, 60, 95)
        _btnEpackConnect.ForeColor = ColorText
        _btnEpackConnect.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        AddHandler _btnEpackConnect.Click, AddressOf BtnEpackConnect_Click
        pConn.Controls.Add(_btnEpackConnect)

        _lblEpackStatus = New Label()
        _lblEpackStatus.AutoSize = True
        _lblEpackStatus.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        _lblEpackStatus.Location = New Point(542, 26)
        pConn.Controls.Add(_lblEpackStatus)

        Dim lblInfo As New Label()
        lblInfo.Text = "Signal checkboxes control visibility in the Electrical Values panel and CSV export."
        lblInfo.Font = New Font("Segoe UI", 7.5, FontStyle.Italic)
        lblInfo.ForeColor = Color.FromArgb(80, 105, 140)
        lblInfo.AutoSize = True
        lblInfo.Location = New Point(12, 56)
        pConn.Controls.Add(lblInfo)

        Dim pGrid As New Panel()
        pGrid.BackColor = ColorBackground
        pGrid.Dock = DockStyle.Fill

        _lblEpackNoSignals = New Label()
        _lblEpackNoSignals.Text = "ePack not connected — connect above to configure signals."
        _lblEpackNoSignals.Font = New Font("Segoe UI", 9, FontStyle.Italic)
        _lblEpackNoSignals.ForeColor = Color.FromArgb(90, 115, 165)
        _lblEpackNoSignals.AutoSize = True
        _lblEpackNoSignals.Location = New Point(20, 20)
        pGrid.Controls.Add(_lblEpackNoSignals)

        _dgvEpack = NewDGV()
        AppendColorCol(_dgvEpack)
        AppendTextCol(_dgvEpack, "ECol_Name", "Signal", 180, False)
        AppendTextCol(_dgvEpack, "ECol_Unit", "Unit", 55, True)

        Dim ePanel As New DataGridViewCheckBoxColumn()
        ePanel.HeaderText = "Panel"
        ePanel.Name = "ECol_Panel"
        ePanel.Width = 55
        ePanel.SortMode = DataGridViewColumnSortMode.NotSortable
        _dgvEpack.Columns.Add(ePanel)

        Dim eCsv As New DataGridViewCheckBoxColumn()
        eCsv.HeaderText = "CSV"
        eCsv.Name = "ECol_CSV"
        eCsv.Width = 55
        eCsv.SortMode = DataGridViewColumnSortMode.NotSortable
        _dgvEpack.Columns.Add(eCsv)

        AppendHiddenCol(_dgvEpack, "ECol_SID")
        AddHandler _dgvEpack.CellPainting, AddressOf DgvEpack_CellPainting
        _dgvEpack.Dock = DockStyle.Fill
        pGrid.Controls.Add(_dgvEpack)

        Dim pESelAll As New Panel()
        pESelAll.BackColor = ColorPanel
        pESelAll.Dock = DockStyle.Bottom
        pESelAll.Height = 30

        Dim lblEP As New Label()
        lblEP.Text = "Panel :"
        lblEP.ForeColor = ColorSubtext
        lblEP.Font = New Font("Segoe UI", 7.5, FontStyle.Bold)
        lblEP.AutoSize = True
        lblEP.Location = New Point(10, 8)
        pESelAll.Controls.Add(lblEP)

        Dim btnEPanelAll As New Button()
        btnEPanelAll.Text = "All"
        btnEPanelAll.Size = New Size(44, 18)
        btnEPanelAll.Location = New Point(55, 6)
        btnEPanelAll.FlatStyle = FlatStyle.Flat
        btnEPanelAll.FlatAppearance.BorderSize = 0
        btnEPanelAll.BackColor = Color.FromArgb(38, 60, 95)
        btnEPanelAll.ForeColor = ColorText
        btnEPanelAll.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        AddHandler btnEPanelAll.Click, Sub(s, ev) SetEpackColumn(3, True)
        pESelAll.Controls.Add(btnEPanelAll)

        Dim btnEPanelNone As New Button()
        btnEPanelNone.Text = "None"
        btnEPanelNone.Size = New Size(44, 18)
        btnEPanelNone.Location = New Point(103, 6)
        btnEPanelNone.FlatStyle = FlatStyle.Flat
        btnEPanelNone.FlatAppearance.BorderSize = 0
        btnEPanelNone.BackColor = Color.FromArgb(38, 60, 95)
        btnEPanelNone.ForeColor = ColorText
        btnEPanelNone.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        AddHandler btnEPanelNone.Click, Sub(s, ev) SetEpackColumn(3, False)
        pESelAll.Controls.Add(btnEPanelNone)

        Dim lblEC As New Label()
        lblEC.Text = "CSV :"
        lblEC.ForeColor = ColorSubtext
        lblEC.Font = New Font("Segoe UI", 7.5, FontStyle.Bold)
        lblEC.AutoSize = True
        lblEC.Location = New Point(168, 8)
        pESelAll.Controls.Add(lblEC)

        Dim btnECsvAll As New Button()
        btnECsvAll.Text = "All"
        btnECsvAll.Size = New Size(44, 18)
        btnECsvAll.Location = New Point(207, 6)
        btnECsvAll.FlatStyle = FlatStyle.Flat
        btnECsvAll.FlatAppearance.BorderSize = 0
        btnECsvAll.BackColor = Color.FromArgb(38, 60, 95)
        btnECsvAll.ForeColor = ColorText
        btnECsvAll.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        AddHandler btnECsvAll.Click, Sub(s, ev) SetEpackColumn(4, True)
        pESelAll.Controls.Add(btnECsvAll)

        Dim btnECsvNone As New Button()
        btnECsvNone.Text = "None"
        btnECsvNone.Size = New Size(44, 18)
        btnECsvNone.Location = New Point(255, 6)
        btnECsvNone.FlatStyle = FlatStyle.Flat
        btnECsvNone.FlatAppearance.BorderSize = 0
        btnECsvNone.BackColor = Color.FromArgb(38, 60, 95)
        btnECsvNone.ForeColor = ColorText
        btnECsvNone.Font = New Font("Segoe UI", 7, FontStyle.Bold)
        AddHandler btnECsvNone.Click, Sub(s, ev) SetEpackColumn(4, False)
        pESelAll.Controls.Add(btnECsvNone)

        pGrid.Controls.Add(pESelAll)

        tab.Controls.Add(pGrid)
        tab.Controls.Add(pConn)
    End Sub

    Private Sub LoadFromPrefs()
        Dim p As UserPreferences = UserPreferences.Instance

        _trackBar.Value = Math.Min(50, Math.Max(1, p.SampleInterval \ 100))
        UpdateIntervalLabel()
        PopulateSignalsGrid(_currentGroup)

        _nudXWindow.Value = CDec(Math.Min(3600, Math.Max(10, p.ChartTimeWindow)))
        UpdateXLabel()
        PopulateAxesGrid()

        If _comboAxY1 IsNot Nothing Then
            _comboAxY1.SelectedIndex = Math.Max(0, Math.Min(p.Y1GroupIndex, _comboAxY1.Items.Count - 1))
        End If
        If _comboAxY2 IsNot Nothing Then
            _comboAxY2.SelectedIndex = Math.Max(0, Math.Min(p.Y2GroupIndex, _comboAxY2.Items.Count - 1))
        End If

        For i As Integer = 0 To 9
            If _hdGroupChecks(i) IsNot Nothing Then _hdGroupChecks(i).Checked = p.HdGroups.Contains(i)
        Next
        For idx As Integer = 0 To 29
            If _hdOrderChecks(idx) IsNot Nothing Then _hdOrderChecks(idx).Checked = p.HdOrders.Contains(idx + 2)
        Next
        If _chkHdCsv IsNot Nothing Then _chkHdCsv.Checked = p.HdExportEnabled

        If _chkEpack IsNot Nothing Then _chkEpack.Checked = p.EpackEnabled
        If _txtEpackIP IsNot Nothing Then _txtEpackIP.Text = p.EpackIP
        If _nudEpackPort IsNot Nothing Then _nudEpackPort.Value = CDec(p.EpackPort)
        UpdateEpackStatusLabel()
        PopulateEpackGrid()
    End Sub

    Private Sub PopulateSignalsGrid(groupFilter As String)
        _dgvSignals.Rows.Clear()
        Dim p As UserPreferences = UserPreferences.Instance
        Dim i As Integer = 0
        For Each r As RegisterDef In RegisterMap.GetRealTimeSignals()
            If Not MatchesGroupFilter(r, groupFilter) Then Continue For
            Dim ri As Integer = _dgvSignals.Rows.Add()
            Dim row As DataGridViewRow = _dgvSignals.Rows(ri)
            row.Cells(SIG_COL_COLOR).Value = ""
            row.Cells(SIG_COL_NAME).Value = r.Name
            row.Cells(SIG_COL_UNIT).Value = r.Unit
            row.Cells(SIG_COL_PANEL).Value = p.IsPanelVisible(r.ID)
            row.Cells(SIG_COL_CSV).Value = p.IsCsvEnabled(r.ID)
            row.Cells(SIG_COL_ID).Value = CInt(r.ID)
            row.DefaultCellStyle.BackColor = If(i Mod 2 = 0, ColorRowEven, ColorRowOdd)
            i += 1
        Next
    End Sub

    Private Sub PopulateAxesGrid()
        _dgvAxes.Rows.Clear()
        Dim p As UserPreferences = UserPreferences.Instance
        For i As Integer = 0 To UserPreferences.AxisGroups.Length - 1
            Dim g As UserPreferences.AxisGroupDef = UserPreferences.AxisGroups(i)
            Dim ri As Integer = _dgvAxes.Rows.Add()
            Dim row As DataGridViewRow = _dgvAxes.Rows(ri)
            row.Cells(AX_COL_COLOR).Value = ""
            row.Cells(AX_COL_NAME).Value = g.Name
            row.Cells(AX_COL_UNIT).Value = g.Unit
            Dim repId As SignalID = g.Members(0)
            row.Cells(AX_COL_MIN).Value = p.GetAxisMin(repId).ToString("G6", CI)
            row.Cells(AX_COL_MAX).Value = p.GetAxisMax(repId).ToString("G6", CI)
            row.Cells(AX_COL_GROUP).Value = i
            row.DefaultCellStyle.BackColor = If(i Mod 2 = 0, ColorRowEven, ColorRowOdd)
        Next
    End Sub

    Private Function MatchesGroupFilter(r As RegisterDef, filter As String) As Boolean
        If r.Group = SignalGroup.ePack Then Return False
        If filter = "ALL" Then Return True
        Select Case filter
            Case "Current" : Return r.Group = SignalGroup.Current
            Case "Voltage" : Return r.Group = SignalGroup.Voltage
            Case "Power" : Return r.Group = SignalGroup.Power
            Case "THD" : Return r.Group = SignalGroup.Harmonic
            Case "Phase" : Return r.Group = SignalGroup.Phase
        End Select
        Return True
    End Function

    Private Sub DgvSignals_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs)
        If e.ColumnIndex <> SIG_COL_COLOR OrElse e.RowIndex < 0 Then Return
        e.PaintBackground(e.CellBounds, True)
        Dim idVal As Object = _dgvSignals.Rows(e.RowIndex).Cells(SIG_COL_ID).Value
        If idVal Is Nothing Then e.Handled = True : Return
        Using br As New SolidBrush(RegisterMap.GetDef(CType(CInt(idVal), SignalID)).PlotColor)
            e.Graphics.FillRectangle(br, e.CellBounds)
        End Using
        e.Handled = True
    End Sub

    Private Sub DgvAxes_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs)
        If e.ColumnIndex <> AX_COL_COLOR OrElse e.RowIndex < 0 Then Return
        e.PaintBackground(e.CellBounds, True)
        Dim grpVal As Object = _dgvAxes.Rows(e.RowIndex).Cells(AX_COL_GROUP).Value
        If grpVal Is Nothing Then e.Handled = True : Return
        Dim gi As Integer = CInt(grpVal)
        If gi >= 0 AndAlso gi < UserPreferences.AxisGroups.Length Then
            Using br As New SolidBrush(UserPreferences.AxisGroups(gi).Clr)
                e.Graphics.FillRectangle(br, e.CellBounds)
            End Using
        End If
        e.Handled = True
    End Sub

    Private Sub FilterBtn_Click(sender As Object, e As EventArgs)
        FlushSignalsGrid()
        _currentGroup = CStr(DirectCast(sender, Button).Tag)
        For Each ctrl As Control In _filterPanel.Controls
            Dim b As Button = TryCast(ctrl, Button)
            If b IsNot Nothing Then
                b.BackColor = If(b.Tag.ToString() = _currentGroup, ColorApply, Color.FromArgb(38, 54, 84))
            End If
        Next
        PopulateSignalsGrid(_currentGroup)
    End Sub

    Private Sub TrackBar_Scroll(sender As Object, e As EventArgs)
        UpdateIntervalLabel()
    End Sub
    Private Sub UpdateIntervalLabel()
        _lblInterval.Text = (_trackBar.Value * 100).ToString() & " ms"
        _lblInterval.Refresh()
    End Sub
    Private Sub NudXWindow_Changed(sender As Object, e As EventArgs)
        UpdateXLabel()
    End Sub
    Private Sub UpdateXLabel()
        Dim s As Double = CDbl(_nudXWindow.Value)
        If s < 60 Then
            _lblXInfo.Text = s.ToString("0") & " s"
        Else
            _lblXInfo.Text = (s / 60.0).ToString("0.#") & " min"
        End If
    End Sub

    Private Sub SetAllGroups(checked As Boolean)
        For Each chk As CheckBox In _hdGroupChecks
            If chk IsNot Nothing Then chk.Checked = checked
        Next
    End Sub
    Private Sub SetAllOrders(checked As Boolean)
        For Each chk As CheckBox In _hdOrderChecks
            If chk IsNot Nothing Then chk.Checked = checked
        Next
    End Sub

    Private Sub FlushSignalsGrid()
        Dim p As UserPreferences = UserPreferences.Instance
        For Each row As DataGridViewRow In _dgvSignals.Rows
            If row.Cells(SIG_COL_ID).Value Is Nothing Then Continue For
            Dim id As SignalID = CType(CInt(row.Cells(SIG_COL_ID).Value), SignalID)
            Dim panelVisible As Boolean = CBool(If(row.Cells(SIG_COL_PANEL).Value, False))
            p.SetPanelVisible(id, panelVisible)
            p.SetCsvEnabled(id, CBool(If(row.Cells(SIG_COL_CSV).Value, False)))
            If Not panelVisible Then p.SetEnabled(id, False)
        Next
    End Sub

    Private Sub FlushAxesGrid()
        Dim p As UserPreferences = UserPreferences.Instance
        For Each row As DataGridViewRow In _dgvAxes.Rows
            If row.Cells(AX_COL_GROUP).Value Is Nothing Then Continue For
            Dim gi As Integer = CInt(row.Cells(AX_COL_GROUP).Value)
            If gi < 0 OrElse gi >= UserPreferences.AxisGroups.Length Then Continue For
            Dim minTxt As String = If(TryCast(row.Cells(AX_COL_MIN).Value, String), "")
            Dim maxTxt As String = If(TryCast(row.Cells(AX_COL_MAX).Value, String), "")
            Dim minV As Double
            Dim maxV As Double
            For Each id As SignalID In UserPreferences.AxisGroups(gi).Members
                If Not RegisterMap.HasDef(id) Then Continue For
                If Not Double.TryParse(minTxt, NumberStyles.Any, CI, minV) Then minV = RegisterMap.GetDef(id).MinVal
                If Not Double.TryParse(maxTxt, NumberStyles.Any, CI, maxV) Then maxV = RegisterMap.GetDef(id).MaxVal
                If minV >= maxV Then maxV = minV + 1.0
                p.SetAxisOverride(id, minV, maxV)
            Next
        Next
        p.ChartTimeWindow = CDbl(_nudXWindow.Value)
        If _comboAxY1 IsNot Nothing Then p.Y1GroupIndex = _comboAxY1.SelectedIndex
        If _comboAxY2 IsNot Nothing Then p.Y2GroupIndex = _comboAxY2.SelectedIndex
    End Sub

    Private Sub FlushHdTab()
        Dim p As UserPreferences = UserPreferences.Instance
        p.HdGroups.Clear()
        For i As Integer = 0 To 9
            If _hdGroupChecks(i) IsNot Nothing AndAlso _hdGroupChecks(i).Checked Then p.HdGroups.Add(i)
        Next
        p.HdOrders.Clear()
        For idx As Integer = 0 To 29
            If _hdOrderChecks(idx) IsNot Nothing AndAlso _hdOrderChecks(idx).Checked Then p.HdOrders.Add(idx + 2)
        Next
        If _chkHdCsv IsNot Nothing Then p.HdExportEnabled = _chkHdCsv.Checked
    End Sub

    Private Sub BtnApply_Click(sender As Object, e As EventArgs)
        FlushSignalsGrid()
        FlushAxesGrid()
        FlushHdTab()
        FlushEpackPanel()
        FlushEpackGrid()
        UserPreferences.Instance.SampleInterval = _trackBar.Value * 100
        UserPreferences.Instance.Save()
        _onApply?.Invoke()
        Me.Close()
    End Sub

    Private Sub BtnReset_Click(sender As Object, e As EventArgs)
        UserPreferences.Instance.ClearAxisOverrides()
        UserPreferences.Instance.ResetHdToDefaults()
        UserPreferences.Instance.Y1GroupIndex = 0
        UserPreferences.Instance.Y2GroupIndex = 3
        UserPreferences.Instance.Save()
        _onApply?.Invoke()
        LoadFromPrefs()
    End Sub

    Private Sub btnSaveCfg_Click(sender As Object, e As EventArgs)
        FlushSignalsGrid()
        FlushAxesGrid()
        Dim p As UserPreferences = UserPreferences.Instance
        p.SampleInterval = _trackBar.Value * 100

        Dim jsonFile As New SaveFileDialog()
        jsonFile.Title = "Save preferences as JSON"
        jsonFile.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        jsonFile.DefaultExt = "json"
        jsonFile.FileName = "electric_meter_config.json"
        jsonFile.OverwritePrompt = True
        jsonFile.InitialDirectory = AppFolder("config")
        If jsonFile.ShowDialog() <> DialogResult.OK Then Return

        Try
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("{")
            sb.AppendLine("  ""sample_ms"": " & p.SampleInterval.ToString() & ",")
            sb.AppendLine("  ""x_window_s"": " & p.ChartTimeWindow.ToString(CI) & ",")
            sb.AppendLine("  ""y1_group"": " & p.Y1GroupIndex.ToString() & ",")
            sb.AppendLine("  ""y2_group"": " & p.Y2GroupIndex.ToString() & ",")
            For Each pair As String() In New String()() {
                    New String() {"chart", "ChartSignals"},
                    New String() {"panel", "PanelSignals"},
                    New String() {"csv", "CsvSignals"}}
                Dim ids As List(Of SignalID)
                Select Case pair(0)
                    Case "chart" : ids = New List(Of SignalID)(p.ChartSignals)
                    Case "panel" : ids = New List(Of SignalID)(p.PanelSignals)
                    Case Else : ids = New List(Of SignalID)(p.CsvSignals)
                End Select
                sb.AppendLine("  """ & pair(0) & """: [")
                For i As Integer = 0 To ids.Count - 1
                    sb.AppendLine("    """ & ids(i).ToString() & """" & If(i < ids.Count - 1, ",", ""))
                Next
                sb.AppendLine("  ],")
            Next
            sb.AppendLine("  ""axis_min"": {")
            Dim minList As New List(Of KeyValuePair(Of SignalID, Double))(p.AxisMinimums)
            For i As Integer = 0 To minList.Count - 1
                sb.AppendLine("    """ & minList(i).Key.ToString() & """: " & minList(i).Value.ToString(CI) & If(i < minList.Count - 1, ",", ""))
            Next
            sb.AppendLine("  },")
            sb.AppendLine("  ""axis_max"": {")
            Dim maxList As New List(Of KeyValuePair(Of SignalID, Double))(p.AxisMaximums)
            For i As Integer = 0 To maxList.Count - 1
                sb.AppendLine("    """ & maxList(i).Key.ToString() & """: " & maxList(i).Value.ToString(CI) & If(i < maxList.Count - 1, ",", ""))
            Next
            sb.AppendLine("  }")
            sb.AppendLine("}")
            File.WriteAllText(jsonFile.FileName, sb.ToString(), System.Text.Encoding.UTF8)
            MessageBox.Show("Configuration saved to:" & Environment.NewLine & jsonFile.FileName, "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error saving config:" & Environment.NewLine & ex.Message, "Save error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnLoadCfg_Click(sender As Object, e As EventArgs)
        Dim jsonFile As New OpenFileDialog()
        jsonFile.Title = "Load preferences from JSON"
        jsonFile.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        jsonFile.InitialDirectory = AppFolder("config")
        If jsonFile.ShowDialog() <> DialogResult.OK Then Return
        Try
            UserPreferences.Instance.LoadFromFile(jsonFile.FileName)
            UserPreferences.Instance.Save()
            LoadFromPrefs()
            MessageBox.Show("Configuration loaded.", "Loaded", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error loading config:" & Environment.NewLine & ex.Message, "Load error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub PopulateEpackGrid()
        If _dgvEpack Is Nothing Then Return
        _dgvEpack.Rows.Clear()
        If Not ModbusTcpMaster.modbustcpisconnected Then
            If _lblEpackNoSignals IsNot Nothing Then _lblEpackNoSignals.Visible = True
            _dgvEpack.Visible = False
            Return
        End If
        If _lblEpackNoSignals IsNot Nothing Then _lblEpackNoSignals.Visible = False
        _dgvEpack.Visible = True
        Dim p As UserPreferences = UserPreferences.Instance
        Dim i As Integer = 0
        For Each r As RegisterDef In RegisterMap.GetRealTimeSignals()
            If r.Group <> SignalGroup.ePack Then Continue For
            Dim ri As Integer = _dgvEpack.Rows.Add()
            Dim row As DataGridViewRow = _dgvEpack.Rows(ri)
            row.Cells(0).Value = ""
            row.Cells(1).Value = r.Name
            row.Cells(2).Value = r.Unit
            row.Cells(3).Value = p.IsPanelVisible(r.ID)
            row.Cells(4).Value = p.IsCsvEnabled(r.ID)
            row.Cells(5).Value = CInt(r.ID)
            row.DefaultCellStyle.BackColor = If(i Mod 2 = 0, ColorRowEven, ColorRowOdd)
            i += 1
        Next
    End Sub

    Private Sub FlushEpackGrid()
        If _dgvEpack Is Nothing OrElse Not _dgvEpack.Visible Then Return
        Dim p As UserPreferences = UserPreferences.Instance
        For Each row As DataGridViewRow In _dgvEpack.Rows
            If row.Cells(5).Value Is Nothing Then Continue For
            Dim id As SignalID = CType(CInt(row.Cells(5).Value), SignalID)
            Dim panelVisible As Boolean = CBool(If(row.Cells(3).Value, False))
            p.SetPanelVisible(id, panelVisible)
            p.SetCsvEnabled(id, CBool(If(row.Cells(4).Value, False)))
            If Not panelVisible Then p.SetEnabled(id, False)
        Next
    End Sub

    Private Sub DgvEpack_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs)
        If e.ColumnIndex <> 0 OrElse e.RowIndex < 0 Then Return
        e.PaintBackground(e.CellBounds, True)
        Dim idVal As Object = _dgvEpack.Rows(e.RowIndex).Cells(5).Value
        If idVal Is Nothing Then e.Handled = True : Return
        Dim sid As SignalID = CType(CInt(idVal), SignalID)
        If RegisterMap.HasDef(sid) Then
            Using br As New SolidBrush(RegisterMap.GetDef(sid).PlotColor)
                e.Graphics.FillRectangle(br, e.CellBounds)
            End Using
        End If
        e.Handled = True
    End Sub

    Private Sub FlushEpackPanel()
        If _chkEpack Is Nothing Then Return
        Dim p As UserPreferences = UserPreferences.Instance
        Dim wasEnabled As Boolean = p.EpackEnabled
        p.EpackEnabled = _chkEpack.Checked
        If _txtEpackIP IsNot Nothing Then p.EpackIP = _txtEpackIP.Text.Trim()
        If _nudEpackPort IsNot Nothing Then p.EpackPort = CInt(_nudEpackPort.Value)
        If p.EpackEnabled AndAlso Not ModbusTcpMaster.modbustcpisconnected Then
            ModbusTcpMaster.OpenTCP(p.EpackIP, p.EpackPort)
            If ModbusTcpMaster.modbustcpisconnected Then RegisterMap.Refresh()
        ElseIf Not p.EpackEnabled AndAlso ModbusTcpMaster.modbustcpisconnected Then
            ModbusTcpMaster.CloseTCP()
            RegisterMap.Refresh()
        End If
        UpdateEpackStatusLabel()
    End Sub

    Private Sub BtnEpackConnect_Click(sender As Object, e As EventArgs)
        Dim p As UserPreferences = UserPreferences.Instance
        If _txtEpackIP IsNot Nothing Then p.EpackIP = _txtEpackIP.Text.Trim()
        If _nudEpackPort IsNot Nothing Then p.EpackPort = CInt(_nudEpackPort.Value)
        If ModbusTcpMaster.modbustcpisconnected Then
            ModbusTcpMaster.CloseTCP()
            RegisterMap.Refresh()
            p.EpackEnabled = False
        Else
            ModbusTcpMaster.OpenTCP(p.EpackIP, p.EpackPort)
            If ModbusTcpMaster.modbustcpisconnected Then
                RegisterMap.Refresh()
                If _chkEpack IsNot Nothing Then _chkEpack.Checked = True
                p.EpackEnabled = True
                For Each r As RegisterDef In RegisterMap.GetRealTimeSignals()
                    If r.Group = SignalGroup.ePack Then
                        p.SetPanelVisible(r.ID, True)
                        p.SetCsvEnabled(r.ID, True)
                        p.SetEnabled(r.ID, True)
                    End If
                Next
            End If
        End If
        p.Save()
        ' Notifier MeterWindow pour regenerer le panel sans fermer le form
        _onApply?.Invoke()
        UpdateEpackStatusLabel()
        PopulateEpackGrid()
    End Sub

    Private Sub UpdateEpackStatusLabel()
        If _lblEpackStatus Is Nothing Then Return
        If ModbusTcpMaster.modbustcpisconnected Then
            _lblEpackStatus.Text = "Connected"
            _lblEpackStatus.ForeColor = Color.FromArgb(0, 210, 140)
            If _btnEpackConnect IsNot Nothing Then
                _btnEpackConnect.Text = "Disconnect"
                _btnEpackConnect.BackColor = Color.FromArgb(100, 40, 30)
            End If
        Else
            _lblEpackStatus.Text = "Disconnected"
            _lblEpackStatus.ForeColor = Color.FromArgb(180, 80, 80)
            If _btnEpackConnect IsNot Nothing Then
                _btnEpackConnect.Text = "Connect"
                _btnEpackConnect.BackColor = Color.FromArgb(38, 60, 95)
            End If
        End If
    End Sub

    Private Sub SetSignalsColumn(colIndex As Integer, checked As Boolean)
        For Each row As DataGridViewRow In _dgvSignals.Rows
            row.Cells(colIndex).Value = checked
        Next
    End Sub

    Private Sub SetEpackColumn(colIndex As Integer, checked As Boolean)
        If _dgvEpack Is Nothing Then Return
        For Each row As DataGridViewRow In _dgvEpack.Rows
            row.Cells(colIndex).Value = checked
        Next
    End Sub

    Private Function NewDGV() As DataGridView
        Dim dgv As New DataGridView()
        dgv.BackgroundColor = ColorBackground
        dgv.BorderStyle = BorderStyle.None
        dgv.GridColor = Color.FromArgb(36, 50, 76)
        dgv.RowHeadersVisible = False
        dgv.AllowUserToAddRows = False
        dgv.AllowUserToDeleteRows = False
        dgv.AllowUserToResizeRows = False
        dgv.MultiSelect = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dgv.ColumnHeadersHeight = 28
        dgv.RowTemplate.Height = 26
        dgv.ScrollBars = ScrollBars.Vertical
        Dim hStyle As New DataGridViewCellStyle()
        hStyle.BackColor = ColorHeader : hStyle.ForeColor = ColorSubtext
        hStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        hStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.ColumnHeadersDefaultCellStyle = hStyle
        Dim cStyle As New DataGridViewCellStyle()
        cStyle.BackColor = ColorRowEven : cStyle.ForeColor = ColorText
        cStyle.SelectionBackColor = ColorSelected : cStyle.SelectionForeColor = Color.White
        dgv.DefaultCellStyle = cStyle
        Return dgv
    End Function

    Private Sub AppendColorCol(dgv As DataGridView)
        Dim c As New DataGridViewTextBoxColumn()
        c.HeaderText = "" : c.Name = "ColorBand" : c.Width = 8
        c.ReadOnly = True : c.Resizable = DataGridViewTriState.False
        c.SortMode = DataGridViewColumnSortMode.NotSortable
        dgv.Columns.Add(c)
    End Sub

    Private Sub AppendTextCol(dgv As DataGridView, name As String, header As String, width As Integer, center As Boolean)
        Dim c As New DataGridViewTextBoxColumn()
        c.HeaderText = header : c.Name = name : c.Width = width
        c.ReadOnly = True : c.SortMode = DataGridViewColumnSortMode.NotSortable
        If center Then
            Dim s As New DataGridViewCellStyle()
            s.ForeColor = ColorSubtext
            s.Alignment = DataGridViewContentAlignment.MiddleCenter
            c.DefaultCellStyle = s
        End If
        dgv.Columns.Add(c)
    End Sub

    Private Sub AppendHiddenCol(dgv As DataGridView, name As String)
        Dim c As New DataGridViewTextBoxColumn()
        c.Name = name : c.Visible = False
        dgv.Columns.Add(c)
    End Sub

    'Private Shared Function AppFolder(subFolder As String) As String
    '    Dim exeDir As String = IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    '    Dim projDir As String = IO.Path.GetDirectoryName(IO.Path.GetDirectoryName(exeDir))
    '    Dim dir As String = IO.Path.Combine(projDir, subFolder)
    '    If Not IO.Directory.Exists(dir) Then IO.Directory.CreateDirectory(dir)
    '    Return dir
    'End Function

    Private Shared Function AppFolder(subFolder As String) As String
        Dim baseDir As String =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Electric_Meter_53U"
            )

        Dim dir As String = Path.Combine(baseDir, subFolder)

        If Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If

        Return dir
    End Function

End Class