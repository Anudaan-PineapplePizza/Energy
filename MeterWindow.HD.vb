Option Strict On
Option Explicit On
Imports System.Drawing.Drawing2D

' =============================================================================
' MeterWindow.HD.vb  - Partial Class (4/4)
' Responsibility : Harmonic Detail (HD) panel.
'
' Layout :
'   - One ROW  per enabled HD group  (HDI1, HDI2 ... HDU3N)
'   - One COL  per enabled harmonic order  (H2, H3 ... H31)
'   - Extra col "Fundamental" shows the corresponding I or U live value
'   - Each harmonic cell contains a vertical gauge bar + numeric value
'   - Status bar docked to Top; DGV fills remaining space (no overlap)
'
' Panel_HD is created at runtime (InitHdPanel). HD readings run on Timer_Sample.
' when Visual Studio regenerates MeterWindow.Designer.vb.
' =============================================================================
Partial Class MeterWindow

    ' RUNTIME CONTROL DECLARATIONS
    Private Panel_HD As System.Windows.Forms.Panel

    ' HD STATE
    ' _harmonicNumber
    Private _harmonicNumber As Integer = 29

    ' _harmonicValues(groupIdx)(orderIdx)  orderIdx 0 = H2 .. 29 = H31
    Private _harmonicValues(_harmonicNumber)() As Double
    Private _isHdPanelBuilt As Boolean = False
    Private _harmonicsGrid As DataGridView = Nothing
    Private _statusLabel As Label = Nothing      ' direct ref for fast tick update

    ' Mappings rebuilt on each InitHdPanel call
    Private _rowToGroupMap() As Integer   ' row  index -> group index (0-9)
    Private _colToOrderMap() As Integer   ' (col - HD_COLUMN_FIRST) -> order index (0-based, 0=H2)
    Friend _isHdFrozen As Boolean = False  ' True = don't update HD grid from Sample tick

    Private Const HD_COLUMN_GROUP As Integer = 0
    Private Const HD_COLUMN_FUND As Integer = 1
    Private Const HD_COLUMN_THD As Integer = 2
    Private Const HD_COLUMN_FIRST As Integer = 4     ' first harmonic order column (col 3 = separator)

    ' Group metadata (index matches ModbusMaster.HDGroup enum)
    Private ReadOnly HD_GroupNames() As String = {
        "HDI1", "HDI2", "HDI3", "HDIN",
        "HDU12", "HDU23", "HDU31",
        "HDU1N", "HDU2N", "HDU3N"
    }
    Private ReadOnly HD_GroupColors() As Color = {
        Color.FromArgb(0, 220, 140),
        Color.FromArgb(0, 190, 120),
        Color.FromArgb(0, 160, 100),
        Color.FromArgb(0, 250, 170),
        Color.FromArgb(80, 160, 255),
        Color.FromArgb(50, 130, 235),
        Color.FromArgb(30, 100, 215),
        Color.FromArgb(110, 185, 255),
        Color.FromArgb(140, 210, 255),
        Color.FromArgb(170, 230, 255)
    }
    ' Fundamental signal (I or U) displayed alongside each HD group
    Private ReadOnly HD_FundamentalSignal() As SignalID = {
        SignalID.I1,    ' HDI1  -> Current L1
        SignalID.I2,    ' HDI2  -> Current L2
        SignalID.I3,    ' HDI3  -> Current L3
        SignalID.I_Neutral,  ' HDIN  -> Current N
        SignalID.U12,   ' HDU12 -> Voltage 1-2
        SignalID.U23,   ' HDU23 -> Voltage 2-3
        SignalID.U31,   ' HDU31 -> Voltage 3-1
        SignalID.U1N,   ' HDU1N -> Voltage 1N
        SignalID.U2N,   ' HDU2N -> Voltage 2N
        SignalID.U3N    ' HDU3N -> Voltage 3N
    }
    ' Corresponding THD signal for each HD group
    Private ReadOnly HD_ThdSignal() As SignalID = {
        SignalID.THDi1,  ' HDI1  -> THDi L1
        SignalID.THDi2,  ' HDI2  -> THDi L2
        SignalID.THDi3,  ' HDI3  -> THDi L3
        SignalID.THDiN,  ' HDIN  -> THDi N
        SignalID.THDu12, ' HDU12 -> THDu 1-2
        SignalID.THDu23, ' HDU23 -> THDu 2-3
        SignalID.THDu31, ' HDU31 -> THDu 3-1
        SignalID.THDu1N, ' HDU1N -> THDu 1N
        SignalID.THDu2N, ' HDU2N -> THDu 2N
        SignalID.THDu3N  ' HDU3N -> THDu 3N
    }

    ' INIT RUNTIME CONTROLS  (called from MeterWindow_Load)
    Friend Sub InitRuntimeControls()

        ' Panel_HD
        Panel_HD = New Panel()
        Panel_HD.BackColor = Color.FromArgb(18, 24, 38)
        Panel_HD.Location = New Point(14, 14)
        Panel_HD.Size = New Size(1244, 670)
        Panel_HD.Visible = False
        Panel_HD.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or
                             AnchorStyles.Right Or AnchorStyles.Bottom
        Panel_Main.Controls.Add(Panel_HD)

        ' Init data buffers
        For i As Integer = 0 To _harmonicNumber
            _harmonicValues(i) = Nothing
        Next
        _rowToGroupMap = New Integer() {}
        _colToOrderMap = New Integer() {}

        ' Y1/Y2 axis group selector panel (above GroupBox_ElecValues)
        Dim pAxis As New Panel()
        pAxis.BackColor = Color.FromArgb(12, 20, 36)
        pAxis.Height = 54
        pAxis.Dock = DockStyle.Top
        GroupBox_ElecValues.Controls.Add(pAxis)
        InitY1Y2Combos(pAxis)
    End Sub

    '=========================================================================
    ' INIT HD PANEL
    ' Builds the DataGridView with rows = enabled groups,
    '                              cols = enabled harmonic orders.
    ' Called once on first ShowView(HD) and again after prefs change.
    '=========================================================================
    Friend Sub InitHdPanel()
        Panel_HD.SuspendLayout()
        Panel_HD.Controls.Clear()

        If _harmonicsGrid IsNot Nothing Then
            RemoveHandler _harmonicsGrid.CellPainting, AddressOf DgvHD_CellPainting
        End If
        _harmonicsGrid = Nothing
        _statusLabel = Nothing

        ' ── Color palette ────────────────────────────────────────────────
        Dim clrBg As Color = Color.FromArgb(18, 24, 38)
        Dim clrHdr As Color = Color.FromArgb(28, 38, 58)
        Dim clrPanel As Color = Color.FromArgb(24, 34, 54)
        Dim clrSub As Color = Color.FromArgb(120, 145, 180)
        Dim clrText As Color = Color.FromArgb(210, 225, 245)
        Dim clrGreen As Color = Color.FromArgb(0, 180, 140)

        ' ── Build row -> group mapping ────────────────────────────────────
        Dim rowList As New List(Of Integer)()
        For i As Integer = 0 To _harmonicNumber
            If _prefs.HdGroups.Contains(i) Then rowList.Add(i)
        Next
        _rowToGroupMap = rowList.ToArray()

        ' -- Build col -> order-index mapping ─────────────────────────────
        ' Columns shown = orders checked in Preferences > Harmonics > Orders
        Dim colList As New List(Of Integer)()
        For idx As Integer = 0 To ModbusRsMaster.HD_HARMONICS_COUNT - 1
            Dim order As Integer = idx + 2
            If _prefs.HdOrders.Contains(order) Then colList.Add(idx)
        Next
        _colToOrderMap = colList.ToArray()

        ' ── Status bar (DockStyle.Top, 32px) ─────────────────────────────
        Dim pStatus As New Panel()
        pStatus.BackColor = clrPanel
        pStatus.Height = 32
        pStatus.Dock = DockStyle.Top

        _statusLabel = New Label()
        _statusLabel.Name = "Lbl_HD_Status"
        _statusLabel.Text = "Waiting for first read..."
        _statusLabel.ForeColor = clrSub
        _statusLabel.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        _statusLabel.AutoSize = True
        _statusLabel.Location = New Point(12, 8)
        pStatus.Controls.Add(_statusLabel)

        Dim lblInfo As New Label()
        lblInfo.Text = _prefs.HdGroups.Count.ToString() & " group(s)  " &
                            Chr(183) & "  " & _colToOrderMap.Length.ToString() & " orders"
        lblInfo.ForeColor = clrGreen
        lblInfo.Font = New Font("Consolas", 7.5, FontStyle.Bold)
        lblInfo.AutoSize = True
        lblInfo.Location = New Point(280, 9)
        pStatus.Controls.Add(lblInfo)

        ' ── DataGridView ─────────────────────────────────────────────────
        _harmonicsGrid = New DataGridView()
        _harmonicsGrid.Dock = DockStyle.None
        _harmonicsGrid.BackgroundColor = clrBg
        _harmonicsGrid.BorderStyle = BorderStyle.None
        _harmonicsGrid.GridColor = Color.FromArgb(36, 50, 76)
        _harmonicsGrid.RowHeadersVisible = False
        _harmonicsGrid.AllowUserToAddRows = False
        _harmonicsGrid.AllowUserToDeleteRows = False
        _harmonicsGrid.AllowUserToResizeRows = False
        _harmonicsGrid.AllowUserToResizeColumns = False
        _harmonicsGrid.MultiSelect = False
        _harmonicsGrid.ReadOnly = True
        _harmonicsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        _harmonicsGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        _harmonicsGrid.ColumnHeadersHeight = 22
        _harmonicsGrid.RowTemplate.Height = 72
        _harmonicsGrid.ScrollBars = ScrollBars.Horizontal
        _harmonicsGrid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        _harmonicsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        ' Header style
        Dim hStyle As New DataGridViewCellStyle()
        hStyle.BackColor = clrHdr
        hStyle.ForeColor = clrSub
        hStyle.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        hStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        _harmonicsGrid.ColumnHeadersDefaultCellStyle = hStyle

        ' Default cell style
        Dim cStyle As New DataGridViewCellStyle()
        cStyle.BackColor = clrBg
        cStyle.ForeColor = clrText
        cStyle.SelectionBackColor = Color.FromArgb(40, 60, 100)
        cStyle.SelectionForeColor = Color.White
        cStyle.Font = New Font("Consolas", 8, FontStyle.Bold)
        _harmonicsGrid.DefaultCellStyle = cStyle

        ' ── Column 0: Group name ─────────────────────────────────────────
        Dim cGroup As New DataGridViewTextBoxColumn()
        cGroup.HeaderText = "Group"
        cGroup.Name = "HD_Group"
        cGroup.Width = 60
        cGroup.ReadOnly = True
        Dim sGrp As New DataGridViewCellStyle()
        sGrp.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        sGrp.Alignment = DataGridViewContentAlignment.MiddleCenter
        sGrp.BackColor = Color.FromArgb(8, 15, 20)
        cGroup.DefaultCellStyle = sGrp
        _harmonicsGrid.Columns.Add(cGroup)

        ' -- Column 1: Fundamental (I or U) ──────────────────────────────
        Dim cFund As New DataGridViewTextBoxColumn()
        cFund.HeaderText = "Fundam."
        cFund.Name = "HD_Fund"
        cFund.Width = 67
        cFund.ReadOnly = True
        Dim sFund As New DataGridViewCellStyle()
        sFund.Font = New Font("Consolas", 7.5, FontStyle.Bold)
        sFund.Alignment = DataGridViewContentAlignment.MiddleLeft
        sFund.BackColor = Color.FromArgb(20, 30, 50)
        sFund.Padding = New Padding(4, 0, 2, 0)
        cFund.DefaultCellStyle = sFund
        _harmonicsGrid.Columns.Add(cFund)

        ' -- Column 2: THD (total harmonic distortion for this group) ─────
        Dim cThd As New DataGridViewTextBoxColumn()
        cThd.HeaderText = "THD"
        cThd.Name = "HD_Thd"
        cThd.Width = 64
        cThd.ReadOnly = True
        Dim sThd As New DataGridViewCellStyle()
        sThd.Font = New Font("Consolas", 7.5, FontStyle.Bold)
        sThd.Alignment = DataGridViewContentAlignment.MiddleLeft
        sThd.BackColor = Color.FromArgb(25, 35, 55)
        sThd.Padding = New Padding(4, 0, 2, 0)
        cThd.DefaultCellStyle = sThd
        _harmonicsGrid.Columns.Add(cThd)

        ' -- Separator column: thin visual divider between fixed cols and harmonic cols
        Dim cSep As New DataGridViewTextBoxColumn()
        cSep.HeaderText = ""
        cSep.Name = "HD_Sep"
        cSep.Width = 4
        cSep.ReadOnly = True
        cSep.Resizable = DataGridViewTriState.False
        cSep.SortMode = DataGridViewColumnSortMode.NotSortable
        Dim sSep As New DataGridViewCellStyle()
        sSep.BackColor = Color.FromArgb(201, 201, 201)
        cSep.DefaultCellStyle = sSep
        _harmonicsGrid.Columns.Add(cSep)

        ' -- Columns 3+: One per enabled harmonic order
        For i As Integer = 0 To _colToOrderMap.Length - 1
            Dim order As Integer = _colToOrderMap(i) + 2
            Dim col As New DataGridViewTextBoxColumn()
            col.HeaderText = order.ToString()    ' "n order of the harmonic"
            col.Name = "HD_H" & order.ToString()
            col.Width = 27
            col.ReadOnly = True
            Dim cs As New DataGridViewCellStyle()
            cs.Alignment = DataGridViewContentAlignment.BottomCenter
            cs.BackColor = clrBg
            col.DefaultCellStyle = cs
            _harmonicsGrid.Columns.Add(col)
        Next



        ' ── Rows: one per enabled group ───────────────────────────────────
        For rowIdx As Integer = 0 To _rowToGroupMap.Length - 1
            Dim g As Integer = _rowToGroupMap(rowIdx)
            Dim ri As Integer = _harmonicsGrid.Rows.Add()
            Dim fundDef As RegisterDef = RegisterMap.GetDef(HD_FundamentalSignal(g))

            ' Group name cell (colored)
            _harmonicsGrid.Rows(ri).Cells("HD_Group").Value = HD_GroupNames(g)
            Dim gsCell As New DataGridViewCellStyle()
            gsCell.ForeColor = HD_GroupColors(g)
            gsCell.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            gsCell.Alignment = DataGridViewContentAlignment.MiddleCenter
            gsCell.BackColor = Color.FromArgb(22, 32, 52)
            _harmonicsGrid.Rows(ri).Cells("HD_Group").Style = gsCell

            ' Fundamental cell placeholder
            _harmonicsGrid.Rows(ri).Cells("HD_Fund").Value =
                fundDef.Name & Environment.NewLine & "-- " & fundDef.Unit

            ' THD cell placeholder
            Dim thdDef As RegisterDef = RegisterMap.GetDef(HD_ThdSignal(g))
            _harmonicsGrid.Rows(ri).Cells("HD_Thd").Value =
                thdDef.Name & Environment.NewLine & "--  %"

            ' Harmonic cells init
            For colHarmIdx As Integer = 0 To _colToOrderMap.Length - 1
                Dim order As Integer = _colToOrderMap(colHarmIdx) + 2
                _harmonicsGrid.Rows(ri).Cells("HD_H" & order.ToString()).Value = "0"
            Next

            ' Alternating row background
            _harmonicsGrid.Rows(ri).DefaultCellStyle.BackColor =
                If(rowIdx Mod 2 = 0, Color.FromArgb(18, 28, 46), Color.FromArgb(22, 34, 56))
        Next

        AddHandler _harmonicsGrid.CellPainting, AddressOf DgvHD_CellPainting

        ' Panel scrollable with HDs values
        Dim pScroll As New Panel()
        pScroll.Dock = DockStyle.Fill
        pScroll.AutoScroll = True
        pScroll.BackColor = clrBg

        Dim dgvH As Integer = _rowToGroupMap.Length * 72 + _harmonicsGrid.ColumnHeadersHeight + 4
        _harmonicsGrid.Dock = DockStyle.None
        _harmonicsGrid.Width = Panel_HD.ClientSize.Width - SystemInformation.VerticalScrollBarWidth
        _harmonicsGrid.Height = dgvH
        _harmonicsGrid.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        pScroll.Controls.Add(_harmonicsGrid)

        ' Status bar en haut, pScroll remplit le reste
        Panel_HD.Controls.Add(pScroll)
        Panel_HD.Controls.Add(pStatus)
        Panel_HD.ResumeLayout(True)
    End Sub

    ' CELL PAINTING  - vertical gauge bars for harmonic columns
    Private Sub DgvHD_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs)
        If e.RowIndex < 0 Then Return

        ' ── Fundamental column (col 1) — custom paint ────────────────────
        If e.ColumnIndex = HD_COLUMN_FUND Then
            If e.RowIndex >= _rowToGroupMap.Length Then Return
            e.PaintBackground(e.CellBounds, True)
            Dim g As Integer = _rowToGroupMap(e.RowIndex)
            Dim fid As SignalID = HD_FundamentalSignal(g)
            Dim def As RegisterDef = RegisterMap.GetDef(fid)

            Dim nameFont As New Font("Segoe UI", 7.5, FontStyle.Bold)
            Dim valFont As New Font("Consolas", 9.5, FontStyle.Bold)
            Dim b As Rectangle = e.CellBounds
            Dim nameRect As New Rectangle(b.X + 4, b.Y + 16, b.Width - 6, 15)
            Dim valRect As New Rectangle(b.X + 4, b.Y + 34, b.Width - 6, 22)

            TextRenderer.DrawText(e.Graphics, def.Name, nameFont, nameRect,
                Color.FromArgb(120, 145, 180),
                TextFormatFlags.Left Or TextFormatFlags.Top)

            Dim valStr As String = "--"
            Dim valClr As Color = Color.FromArgb(70, 90, 115)
            If _latestReadings.ContainsKey(fid) Then
                valStr = FormatFundamentalValue(fid, _latestReadings(fid)) & " " & def.Unit
                valClr = HD_GroupColors(g)
            End If
            TextRenderer.DrawText(e.Graphics, valStr, valFont, valRect,
                valClr, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)

            nameFont.Dispose() : valFont.Dispose()
            e.Handled = True
            Return
        End If

        ' ── THD column (col 2) — custom paint ────────────────────────────
        If e.ColumnIndex = HD_COLUMN_THD Then
            If e.RowIndex >= _rowToGroupMap.Length Then Return
            e.PaintBackground(e.CellBounds, True)
            Dim g As Integer = _rowToGroupMap(e.RowIndex)
            Dim tid As SignalID = HD_ThdSignal(g)
            Dim def As RegisterDef = RegisterMap.GetDef(tid)

            Dim nameFont As New Font("Segoe UI", 7.5, FontStyle.Bold)
            Dim valFont As New Font("Consolas", 9.5, FontStyle.Bold)
            Dim b As Rectangle = e.CellBounds
            Dim nameRect As New Rectangle(b.X + 4, b.Y + 16, b.Width - 6, 15)
            Dim valRect As New Rectangle(b.X + 4, b.Y + 34, b.Width - 6, 22)

            TextRenderer.DrawText(e.Graphics, def.Name, nameFont, nameRect,
                Color.FromArgb(120, 145, 180),
                TextFormatFlags.Left Or TextFormatFlags.Top)

            Dim valStr As String = "--"
            Dim valClr As Color = Color.FromArgb(70, 90, 115)
            If _latestReadings.ContainsKey(tid) Then
                valStr = _latestReadings(tid).ToString("0.0") & " %"
                ' Color intensity: green < 5%, yellow 5-10%, orange/red > 10%
                Dim thdVal As Double = _latestReadings(tid)
                valClr = If(thdVal < 5.0, HD_GroupColors(g),
                            If(thdVal < 10.0, Color.FromArgb(255, 200, 60),
                                              Color.FromArgb(255, 90, 60)))
            End If
            TextRenderer.DrawText(e.Graphics, valStr, valFont, valRect,
                valClr, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)

            nameFont.Dispose() : valFont.Dispose()
            e.Handled = True
            Return
        End If

        ' ── Harmonic columns (col >= HD_COLUMN_FIRST) — gauge bar ───────────
        If e.ColumnIndex < HD_COLUMN_FIRST Then Return
        Dim colHarmIdx As Integer = e.ColumnIndex - HD_COLUMN_FIRST
        If colHarmIdx >= _colToOrderMap.Length Then Return
        If e.RowIndex >= _rowToGroupMap.Length Then Return

        e.PaintBackground(e.CellBounds, True)

        Dim grp As Integer = _rowToGroupMap(e.RowIndex)
        Dim idx As Integer = _colToOrderMap(colHarmIdx)
        Dim pct As Double = 0.0
        If _harmonicValues(grp) IsNot Nothing AndAlso idx < _harmonicValues(grp).Length Then
            pct = Math.Max(0, Math.Min(100, _harmonicValues(grp)(idx)))
        End If

        Dim bnd As Rectangle = e.CellBounds
        Dim padX As Integer = 3
        Dim textH As Integer = 14
        Dim barX As Integer = bnd.X + padX
        Dim barW As Integer = bnd.Width - padX * 2
        Dim barTop As Integer = bnd.Y + 3
        Dim barBot As Integer = bnd.Bottom - textH - 3
        Dim barH As Integer = barBot - barTop

        If barH > 2 AndAlso barW > 2 Then
            Using brBg As New SolidBrush(Color.FromArgb(30, 42, 64))
                e.Graphics.FillRectangle(brBg, barX, barTop, barW, barH)
            End Using

            Dim fillH As Integer = CInt((pct / 100.0) * barH)
            If fillH > 0 Then
                Dim barColor As Color = HD_GroupColors(grp)
                Dim darkColor As Color = Color.FromArgb(
                    CInt(barColor.R * 0.35), CInt(barColor.G * 0.35), CInt(barColor.B * 0.35))
                Dim gradRect As New Rectangle(barX, barBot - fillH, barW, fillH + 1)
                Try
                    Using br As New LinearGradientBrush(
                            gradRect, darkColor, barColor, LinearGradientMode.Vertical)
                        e.Graphics.FillRectangle(br, barX, barBot - fillH, barW, fillH)
                    End Using
                Catch
                    Using br As New SolidBrush(barColor)
                        e.Graphics.FillRectangle(br, barX, barBot - fillH, barW, fillH)
                    End Using
                End Try
            End If

            Using pen As New Pen(Color.FromArgb(45, 65, 95))
                e.Graphics.DrawRectangle(pen, barX, barTop, barW - 1, barH - 1)
            End Using
        End If

        ' Tiny value text (1 decimal, only if meaningful)
        If pct >= 0.1 Then
            Dim valText As String = pct.ToString("0.0")
            Dim textRect As Rectangle = New Rectangle(bnd.X, bnd.Bottom - textH - 1, bnd.Width, textH)
            TextRenderer.DrawText(
                e.Graphics, valText,
                New Font("Segoe UI", 6.5, FontStyle.Regular),
                textRect, HD_GroupColors(grp),
                TextFormatFlags.HorizontalCenter Or TextFormatFlags.Bottom)
        End If

        e.Handled = True
    End Sub

    ' ─────────────────────────────────────────────────────────────────────────────
    ' TIMER HD TICK  - slow independent Modbus read
    ' Called every Timer_Sample tick when HD panel is visible and not frozen.
    ' ─────────────────────────────────────────────────────────────────────────────
    Friend Sub ReadAndRefreshHD()
        If Not IsConnected() AndAlso Not IsDemoMode Then Return
        If _isHdFrozen Then Return
        If Not _isHdPanelBuilt Then Return

        ' Read each enabled group
        For i As Integer = 0 To 9
            If Not _prefs.HdGroups.Contains(i) Then Continue For
            Dim data() As Double = ModbusRsMaster.ReadHarmonics(CType(i, ModbusRsMaster.HDGroup))
            If data IsNot Nothing AndAlso data.Length = ModbusRsMaster.HD_HARMONICS_COUNT Then
                _harmonicValues(i) = data
            End If
        Next

        ' Update status label
        If _statusLabel IsNot Nothing Then
            _statusLabel.Text = "Last read: " & DateTime.Now.ToString("HH:mm:ss")
            _statusLabel.ForeColor = Color.FromArgb(0, 210, 160)
        End If

        RefreshHdGrid()

        '' CSV HD export
        'If _isRecording AndAlso Not _csvPaused AndAlso _prefs.HdExportEnabled Then
        '    WriteHdCsvLines()
        'End If

    End Sub

    ' ─────────────────────────────────────────────────────────────────────────────
    ' REFRESH HD GRID
    ' rows = groups, cols = harmonic orders + fundamental
    ' ─────────────────────────────────────────────────────────────────────────────
    Private Sub RefreshHdGrid()
        If _harmonicsGrid Is Nothing Then Return
        If _rowToGroupMap Is Nothing OrElse _rowToGroupMap.Length = 0 Then Return

        For rowIdx As Integer = 0 To _rowToGroupMap.Length - 1
            Dim g As Integer = _rowToGroupMap(rowIdx)
            If rowIdx >= _harmonicsGrid.Rows.Count Then Exit For

            ' Last Fundamental value read — stored as tag, painted via CellPainting
            _harmonicsGrid.Rows(rowIdx).Cells("HD_Fund").Value = ""

            ' Harmonic order cells (value stored for painting)
            For colHarmIdx As Integer = 0 To _colToOrderMap.Length - 1
                Dim order As Integer = _colToOrderMap(colHarmIdx) + 2
                Dim colName As String = "HD_H" & order.ToString()
                If Not _harmonicsGrid.Columns.Contains(colName) Then Continue For
                Dim val As Double = 0
                If _harmonicValues(g) IsNot Nothing AndAlso _colToOrderMap(colHarmIdx) < _harmonicValues(g).Length Then
                    val = _harmonicValues(g)(_colToOrderMap(colHarmIdx))
                End If
                _harmonicsGrid.Rows(rowIdx).Cells(colName).Value = val.ToString("G5", System.Globalization.CultureInfo.InvariantCulture)
            Next
        Next

        _harmonicsGrid.Invalidate()
    End Sub

    Private Function FormatFundamentalValue(id As SignalID, v As Double) As String
        Select Case id
            Case SignalID.I1, SignalID.I2, SignalID.I3, SignalID.I_Neutral
                Return v.ToString("0.000")
            Case Else
                Return v.ToString("0.0")
        End Select
    End Function

End Class