Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.IO
Imports ClosedXML.Excel

' =============================================================================
' MeterWindow.CsvRecording.vb  -- Partial Class (3/4)
' =============================================================================
Partial Class MeterWindow

#Region "1 - Champs d etat"

    Friend _isRecording As Boolean = False
    Friend _csvPaused As Boolean = False

    Private Enum ExportFormat
        XlsxOnly
        CsvOnly
        Both
        Cancelled
    End Enum
    Private _exportFormat As ExportFormat = ExportFormat.Both

    Private _xlsxPath As String = ""
    Private _csvPath As String = ""
    Private _recStartDateTime As DateTime = DateTime.Now

    Private _sigHeaders As List(Of String) = Nothing
    Private _sigRows As List(Of Double()) = Nothing
    Private _sigMeta As List(Of (Name As String, Unit As String, ColorHex As String, Group As String)) = Nothing

    Private _hdBuffer As SortedDictionary(Of Double, Dictionary(Of Integer, Double())) = Nothing
    Private _hdGroups As Integer() = New Integer() {}
    Private _hdOrders As Integer() = New Integer() {}

    Private _backupSigWriter As IO.StreamWriter = Nothing
    Private _backupHdWriter As IO.StreamWriter = Nothing
    Private _backupSigPath As String = ""
    Private _backupHdPath As String = ""

#End Region

#Region "2 - Dialog de choix de format"

    Private Function ShowExportFormatDialog() As ExportFormat
        Dim chosen As ExportFormat = ExportFormat.Cancelled

        Using frm As New Form()
            frm.Text = "Export format"
            frm.Size = New Size(340, 160)
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.StartPosition = FormStartPosition.CenterParent
            frm.MaximizeBox = False : frm.MinimizeBox = False
            frm.BackColor = Drawing.Color.FromArgb(18, 28, 50)
            frm.ForeColor = Drawing.Color.FromArgb(200, 215, 240)

            Dim lbl As New Label()
            lbl.Text = "Choose output format for this recording :"
            lbl.Font = New Font("Segoe UI", 9)
            lbl.AutoSize = True
            lbl.Location = New Point(14, 14)
            frm.Controls.Add(lbl)

            Dim clrBtn As Drawing.Color = Drawing.Color.FromArgb(38, 60, 95)
            Dim clrFg As Drawing.Color = Drawing.Color.FromArgb(200, 215, 240)
            Dim btnFont As New Font("Segoe UI", 8.5, FontStyle.Bold)

            Dim btnBoth As New Button()
            btnBoth.Text = "Excel + CSV"
            btnBoth.Size = New Size(92, 32)
            btnBoth.Location = New Point(14, 56)
            btnBoth.FlatStyle = FlatStyle.Flat
            btnBoth.FlatAppearance.BorderSize = 0
            btnBoth.BackColor = Drawing.Color.FromArgb(0, 100, 60)
            btnBoth.ForeColor = Drawing.Color.White
            btnBoth.Font = btnFont
            AddHandler btnBoth.Click, Sub(s2 As Object, ev2 As EventArgs)
                                          chosen = ExportFormat.Both
                                          frm.Close()
                                      End Sub
            frm.Controls.Add(btnBoth)

            Dim btnXlsx As New Button()
            btnXlsx.Text = "Excel only"
            btnXlsx.Size = New Size(92, 32)
            btnXlsx.Location = New Point(116, 56)
            btnXlsx.FlatStyle = FlatStyle.Flat
            btnXlsx.FlatAppearance.BorderSize = 0
            btnXlsx.BackColor = clrBtn
            btnXlsx.ForeColor = clrFg
            btnXlsx.Font = btnFont
            AddHandler btnXlsx.Click, Sub(s2 As Object, ev2 As EventArgs)
                                          chosen = ExportFormat.XlsxOnly
                                          frm.Close()
                                      End Sub
            frm.Controls.Add(btnXlsx)

            Dim btnCsv As New Button()
            btnCsv.Text = "CSV only"
            btnCsv.Size = New Size(92, 32)
            btnCsv.Location = New Point(218, 56)
            btnCsv.FlatStyle = FlatStyle.Flat
            btnCsv.FlatAppearance.BorderSize = 0
            btnCsv.BackColor = clrBtn
            btnCsv.ForeColor = clrFg
            btnCsv.Font = btnFont
            AddHandler btnCsv.Click, Sub(s2 As Object, ev2 As EventArgs)
                                         chosen = ExportFormat.CsvOnly
                                         frm.Close()
                                     End Sub
            frm.Controls.Add(btnCsv)

            frm.ShowDialog(Me)
        End Using
        Return chosen
    End Function

#End Region

#Region "3 - Cycle d enregistrement"

    Private Sub Button_Record_Click(sender As Object, e As EventArgs) Handles Button_Record.Click
        If _isRecording Then
            ShowStopRecordingDialog()
        Else
            StartRecording()
        End If
    End Sub

    ' Dialog de fin d enregistrement
    '   "Save & Close"       -> sauvegarde le fichier et arrete le record
    '   "Discard recording"  -> arrete sans sauvegarder, supprime le backup
    '   "Continue"           -> ferme le dialog, le record continue
    Private Sub ShowStopRecordingDialog()
        Dim result As DialogResult = DialogResult.Cancel

        Using frm As New Form()
            frm.Text = "Stop recording ?"
            frm.Size = New Size(310, 130)
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.StartPosition = FormStartPosition.CenterParent
            frm.MaximizeBox = False : frm.MinimizeBox = False
            frm.BackColor = Drawing.Color.FromArgb(18, 28, 50)
            frm.ForeColor = Drawing.Color.FromArgb(200, 215, 240)

            Dim lbl As New Label()
            lbl.Text = "What do you want to do with this recording ?"
            lbl.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            lbl.AutoSize = True
            lbl.Location = New Point(16, 16)
            frm.Controls.Add(lbl)

            Dim btnFont As New Font("Segoe UI", 8.5, FontStyle.Bold)
            Dim y As Integer = 45

            ' Bouton 1 : Sauvegarder
            Dim btnSave As New Button()
            btnSave.Text = "Save Recording"
            btnSave.Size = New Size(115, 34)
            btnSave.Location = New Point(16, y)
            btnSave.FlatStyle = FlatStyle.Flat
            btnSave.FlatAppearance.BorderSize = 0
            btnSave.BackColor = Drawing.Color.FromArgb(0, 110, 65)
            btnSave.ForeColor = Drawing.Color.White
            btnSave.Font = btnFont
            btnSave.Cursor = Cursors.Hand
            AddHandler btnSave.Click, Sub(s As Object, ev As EventArgs)
                                          result = DialogResult.Yes
                                          frm.Close()
                                      End Sub
            frm.Controls.Add(btnSave)

            ' Bouton 2 : Supprimer sans sauvegarder
            Dim btnDiscard As New Button()
            btnDiscard.Text = "Discard Recording"
            btnDiscard.Size = New Size(138, 34)
            btnDiscard.Location = New Point(142, y)
            btnDiscard.FlatStyle = FlatStyle.Flat
            btnDiscard.FlatAppearance.BorderSize = 0
            btnDiscard.BackColor = Drawing.Color.FromArgb(100, 28, 28)
            btnDiscard.ForeColor = Drawing.Color.FromArgb(255, 200, 200)
            btnDiscard.Font = btnFont
            btnDiscard.Cursor = Cursors.Hand
            AddHandler btnDiscard.Click, Sub(s As Object, ev As EventArgs)
                                             result = DialogResult.No
                                             frm.Close()
                                         End Sub
            frm.Controls.Add(btnDiscard)

            frm.ShowDialog(Me)
        End Using

        Select Case result
            Case DialogResult.Yes
                StopRecording(saveFiles:=True)
            Case DialogResult.No
                StopRecording(saveFiles:=False)
            Case Else
                ' Continue : ne rien faire, l'enregistrement continue.
        End Select
    End Sub

    Private Sub Button_RecordPause_Click(sender As Object, e As EventArgs) Handles Button_RecordPause.Click
        If Not _isRecording Then Return
        _csvPaused = Not _csvPaused
        UpdateRecordButton()
    End Sub

    Private Sub StartRecording()
        _exportFormat = ShowExportFormatDialog()
        If _exportFormat = ExportFormat.Cancelled Then Return

        Dim baseName As String = "record_" & Format(Now, "yyyyMMdd_HHmmss")
        Dim dlg As New SaveFileDialog()
        dlg.OverwritePrompt = True

        If _exportFormat = ExportFormat.CsvOnly Then
            dlg.Title = "Save recording as CSV"
            dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
            dlg.DefaultExt = "csv"
            dlg.FileName = baseName & ".csv"
            dlg.InitialDirectory = AppFolder("csv")
        Else
            dlg.Title = "Save recording as Excel"
            dlg.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            dlg.DefaultExt = "xlsx"
            dlg.FileName = baseName & ".xlsx"
            dlg.InitialDirectory = AppFolder("xlsx")
        End If

        If dlg.ShowDialog() <> DialogResult.OK Then Return

        Select Case _exportFormat
            Case ExportFormat.CsvOnly
                _csvPath = dlg.FileName
                _xlsxPath = ""
            Case ExportFormat.XlsxOnly
                _xlsxPath = dlg.FileName
                _csvPath = ""
            Case ExportFormat.Both
                _xlsxPath = dlg.FileName
                _csvPath = Path.Combine(AppFolder("csv"),
                    Path.GetFileNameWithoutExtension(dlg.FileName) & ".csv")
        End Select
        _recStartDateTime = DateTime.Now

        _sigHeaders = New List(Of String)()
        _sigMeta = New List(Of (Name As String, Unit As String, ColorHex As String, Group As String))()
        _sigHeaders.Add("Time (s)")
        For Each r As RegisterDef In RegisterMap.Registers
            If _prefs.IsCsvEnabled(r.ID) Then
                _sigHeaders.Add(r.Name & If(r.Unit <> "", " (" & r.Unit & ")", ""))
                _sigMeta.Add((r.Name, r.Unit, String.Format("{0:X8}", r.PlotColor.ToArgb()),
                             CInt(r.Group).ToString()))
            End If
        Next
        _sigRows = New List(Of Double())()

        ' HD : activer des que des groupes sont configures (independamment de HdExportEnabled)
        Dim gList As New List(Of Integer)()
        For g As Integer = 0 To 29
            If _prefs.HdGroups.Contains(g) Then gList.Add(g)
        Next
        _hdGroups = gList.ToArray()

        Dim oList As New List(Of Integer)()
        For idx As Integer = 0 To ModbusRsMaster.HD_HARMONICS_COUNT - 1
            Dim order As Integer = idx + 2
            If _prefs.HdOrders.Contains(order) Then oList.Add(order)
        Next
        _hdOrders = oList.ToArray()

        _hdBuffer = If(_hdGroups.Length > 0 AndAlso _hdOrders.Length > 0,
            New SortedDictionary(Of Double, Dictionary(Of Integer, Double()))(),
            Nothing)

        _isRecording = True : _csvPaused = False

        Dim backupBase As String = Path.GetFileNameWithoutExtension(
            If(_xlsxPath <> "", _xlsxPath, _csvPath))
        OpenBackupWriters(backupBase)

        ChartReset()
        UpdateRecordButton()
    End Sub

    ' saveFiles=True  : sauvegarde normale
    ' saveFiles=False : supprime tout sans sauvegarder
    Friend Sub StopRecording(Optional saveFiles As Boolean = True)
        _isRecording = False : _csvPaused = False
        UpdateRecordButton()

        If Not saveFiles Then
            ' Supprimer backup et ne rien ecrire
            CloseBackupWriters(deleteFiles:=True)
            ClearBuffers()
            Return
        End If

        If _sigRows Is Nothing OrElse _sigRows.Count = 0 Then
            ClearBuffers() : Return
        End If

        Dim xlsxOk As Boolean = False
        Dim csvOk As Boolean = False
        Dim errMsg As String = ""

        If _exportFormat = ExportFormat.XlsxOnly OrElse _exportFormat = ExportFormat.Both Then
            Try
                WriteXlsx(_xlsxPath)
                xlsxOk = True
            Catch ex As Exception
                errMsg &= "Excel error: " & ex.Message & Environment.NewLine
            End Try
        End If

        If _exportFormat = ExportFormat.CsvOnly OrElse _exportFormat = ExportFormat.Both Then
            Try
                WriteFinalCsv(_csvPath)
                csvOk = True
            Catch ex As Exception
                errMsg &= "CSV error: " & ex.Message & Environment.NewLine
            End Try
        End If

        CloseBackupWriters(deleteFiles:=(xlsxOk OrElse csvOk))

        If xlsxOk OrElse csvOk Then
            Dim info As New System.Text.StringBuilder()
            info.AppendLine("Recording saved.")
            If xlsxOk Then info.AppendLine("  Excel : " & _xlsxPath)
            If csvOk Then info.AppendLine("  CSV   : " & _csvPath)
            info.AppendLine()
            info.Append("Signals : " & _sigRows.Count.ToString() & " rows")
            If _hdBuffer IsNot Nothing Then
                info.Append("   |   HD : " & _hdBuffer.Count.ToString() & " rows")
            End If
            MessageBox.Show(info.ToString(), "Recording saved",
                MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        If errMsg <> "" Then
            MessageBox.Show(errMsg, "Save error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        ClearBuffers()
    End Sub

#End Region

#Region "4 - Ecriture par tick"

    Friend Sub WriteCsvLine()
        If Not _isRecording OrElse _sigRows Is Nothing Then Return
        Dim ci As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
        Dim arr As New List(Of Double)()
        arr.Add(Main_Chart.ElapsedSeconds)
        For Each r As RegisterDef In RegisterMap.Registers
            If _prefs.IsCsvEnabled(r.ID) Then
                arr.Add(If(_latestReadings.ContainsKey(r.ID), _latestReadings(r.ID), 0.0))
            End If
        Next
        Dim row As Double() = arr.ToArray()
        _sigRows.Add(row)

        If _backupSigWriter IsNot Nothing Then
            Try
                _backupSigWriter.WriteLine(String.Join(",",
                    Array.ConvertAll(row, Function(d As Double) d.ToString("G10", ci))))
            Catch
            End Try
        End If
    End Sub

    ' Appele a chaque tick — le buffer HD est actif des que des groupes sont configures,
    ' independamment de _prefs.HdExportEnabled (qui ne concerne que l export xlsx)
    Friend Sub WriteHdCsvLines()
        If Not _isRecording OrElse _hdBuffer Is Nothing Then Return
        If _hdGroups.Length = 0 OrElse _hdOrders.Length = 0 Then Return

        Dim ts As Double = Main_Chart.ElapsedSeconds
        If Not _hdBuffer.ContainsKey(ts) Then
            _hdBuffer(ts) = New Dictionary(Of Integer, Double())()
        End If

        Dim ci As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture

        For Each g As Integer In _hdGroups
            If _harmonicValues(g) Is Nothing Then Continue For

            Dim vals(_hdOrders.Length - 1) As Double
            For i As Integer = 0 To _hdOrders.Length - 1
                Dim idx As Integer = _hdOrders(i) - 2
                vals(i) = If(idx >= 0 AndAlso idx < _harmonicValues(g).Length,
                    _harmonicValues(g)(idx), 0.0)
            Next
            _hdBuffer(ts)(g) = vals

            If _backupHdWriter IsNot Nothing Then
                Try
                    Dim sb As New System.Text.StringBuilder()
                    sb.Append(ts.ToString("G10", ci))
                    sb.Append(",")
                    sb.Append(HD_GroupNames(g))
                    For Each v As Double In vals
                        sb.Append(",")
                        sb.Append(v.ToString("0.00", ci))
                    Next
                    _backupHdWriter.WriteLine(sb.ToString())
                Catch
                End Try
            End If
        Next
    End Sub

#End Region

#Region "5 - Sortie Excel (ClosedXML)"

    Private Sub WriteXlsx(path As String)
        Using wb As New XLWorkbook()
            Dim ws1 As IXLWorksheet = wb.Worksheets.Add("Signals")
            ws1.TabColor = XLColor.FromColor(Drawing.Color.FromArgb(68, 114, 196))
            For col As Integer = 1 To _sigHeaders.Count
                ws1.Cell(1, col).Value = _sigHeaders(col - 1)
            Next
            StyleXlsxHeader(ws1, _sigHeaders.Count)
            For rowIdx As Integer = 0 To _sigRows.Count - 1
                Dim xlRow As Integer = rowIdx + 2
                Dim arr As Double() = _sigRows(rowIdx)
                For col As Integer = 1 To Math.Min(arr.Length, _sigHeaders.Count)
                    ws1.Cell(xlRow, col).Value = arr(col - 1)
                Next
                If rowIdx Mod 2 = 0 Then
                    ws1.Row(xlRow).Style.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242)
                End If
            Next
            ws1.SheetView.FreezeRows(1)
            ws1.Columns().AdjustToContents()

            Dim ws2 As IXLWorksheet = wb.Worksheets.Add("Harmonics")
            ws2.TabColor = XLColor.FromColor(Drawing.Color.FromArgb(112, 173, 71))

            If _hdBuffer IsNot Nothing AndAlso _hdBuffer.Count > 0 Then
                Dim hdCols As New List(Of String)()
                hdCols.Add("Time (s)")
                For Each g As Integer In _hdGroups
                    For Each order As Integer In _hdOrders
                        hdCols.Add(HD_GroupNames(g) & "_H" & order.ToString() & " (%)")
                    Next
                Next
                For col As Integer = 1 To hdCols.Count
                    ws2.Cell(1, col).Value = hdCols(col - 1)
                Next
                StyleXlsxHeader(ws2, hdCols.Count)

                Dim xlRow As Integer = 2
                For Each kvp As KeyValuePair(Of Double, Dictionary(Of Integer, Double())) In _hdBuffer
                    ws2.Cell(xlRow, 1).Value = kvp.Key
                    Dim col As Integer = 2
                    For Each g As Integer In _hdGroups
                        For Each order As Integer In _hdOrders
                            Dim idx As Integer = order - 2
                            Dim val As Double = 0.0
                            If kvp.Value.ContainsKey(g) Then
                                Dim arr As Double() = kvp.Value(g)
                                If idx >= 0 AndAlso idx < arr.Length Then val = arr(idx)
                            End If
                            ws2.Cell(xlRow, col).Value = val
                            col += 1
                        Next
                    Next
                    If (xlRow - 2) Mod 2 = 0 Then
                        ws2.Row(xlRow).Style.Fill.BackgroundColor = XLColor.FromArgb(242, 242, 242)
                    End If
                    xlRow += 1
                Next
                ws2.SheetView.FreezeRows(1)
                ws2.SheetView.FreezeColumns(1)
                ws2.Columns().AdjustToContents()
            Else
                ws2.Cell(1, 1).Value = "No HD data recorded."
                ws2.Cell(2, 1).Value = "Enable harmonic groups in Preferences > Harmonics."
            End If

            wb.SaveAs(path)
        End Using
    End Sub

    Private Shared Sub StyleXlsxHeader(ws As IXLWorksheet, colCount As Integer)
        Dim hdr As IXLRange = ws.Range(ws.Cell(1, 1), ws.Cell(1, colCount))
        hdr.Style.Font.Bold = True
        hdr.Style.Fill.BackgroundColor = XLColor.FromArgb(217, 217, 217)
        hdr.Style.Border.BottomBorder = XLBorderStyleValues.Double
        ws.Row(1).Height = 22
    End Sub

#End Region

#Region "6 - Sortie CSV"

    Private Sub WriteFinalCsv(path As String)
        Dim ci As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
        Using sw As New IO.StreamWriter(path, False, System.Text.Encoding.UTF8)
            sw.WriteLine("# Electric Meter Monitor")
            sw.WriteLine("# Date," & _recStartDateTime.ToString("yyyy-MM-dd HH:mm:ss"))
            sw.WriteLine("# SampleInterval," & _prefs.SampleInterval.ToString())
            sw.WriteLine("# Rows," & _sigRows.Count.ToString())

            If _sigMeta IsNot Nothing Then
                For i As Integer = 0 To _sigMeta.Count - 1
                    sw.WriteLine("# Signal," & i.ToString() & "," &
                        _sigMeta(i).Name & "," &
                        _sigMeta(i).Unit & "," &
                        _sigMeta(i).ColorHex & "," &
                        _sigMeta(i).Group)
                Next
            End If

            sw.WriteLine(String.Join(",", _sigHeaders))
            For Each row As Double() In _sigRows
                sw.WriteLine(String.Join(",",
                    Array.ConvertAll(row, Function(d As Double) d.ToString("G10", ci))))
            Next

            If _hdBuffer IsNot Nothing AndAlso _hdBuffer.Count > 0 Then
                sw.WriteLine("# HD_Rows," & _hdBuffer.Count.ToString())
                For Each g As Integer In _hdGroups
                    sw.WriteLine("# HD_Group," & g.ToString() & "," & HD_GroupNames(g))
                Next
                sw.WriteLine("# HD_Orders," & String.Join(",",
                    Array.ConvertAll(_hdOrders, Function(o As Integer) o.ToString())))
                sw.WriteLine("# HD_SECTION")

                Dim hdHdr As New System.Text.StringBuilder("Time (s)")
                For Each g As Integer In _hdGroups
                    For Each order As Integer In _hdOrders
                        hdHdr.Append(",")
                        hdHdr.Append(HD_GroupNames(g) & "_H" & order.ToString() & " (%)")
                    Next
                Next
                sw.WriteLine(hdHdr.ToString())

                For Each kvp As KeyValuePair(Of Double, Dictionary(Of Integer, Double())) In _hdBuffer
                    Dim sb As New System.Text.StringBuilder()
                    sb.Append(kvp.Key.ToString("G10", ci))
                    For Each g As Integer In _hdGroups
                        For Each order As Integer In _hdOrders
                            sb.Append(",")
                            Dim idx As Integer = order - 2
                            Dim val As Double = 0.0
                            If kvp.Value.ContainsKey(g) Then
                                Dim arr As Double() = kvp.Value(g)
                                If idx >= 0 AndAlso idx < arr.Length Then val = arr(idx)
                            End If
                            sb.Append(val.ToString("G10", ci))
                        Next
                    Next
                    sw.WriteLine(sb.ToString())
                Next
            End If
        End Using
    End Sub

#End Region

#Region "7 - Backup crash-safe"

    Private Sub OpenBackupWriters(baseName As String)
        _backupSigPath = Path.Combine(AppFolder("csv"), baseName & "_backup.csv")
        Try
            _backupSigWriter = New IO.StreamWriter(_backupSigPath, False, System.Text.Encoding.UTF8)
            _backupSigWriter.WriteLine(String.Join(",", _sigHeaders))
            _backupSigWriter.AutoFlush = True
        Catch
            _backupSigWriter = Nothing
        End Try

        If _hdBuffer IsNot Nothing Then
            _backupHdPath = Path.Combine(AppFolder("csv"), baseName & "_backup_hd.csv")
            Try
                _backupHdWriter = New IO.StreamWriter(_backupHdPath, False, System.Text.Encoding.UTF8)
                Dim hdHead As New System.Text.StringBuilder("Time (s),Group")
                For Each order As Integer In _hdOrders
                    hdHead.Append(",H" & order.ToString() & " (%)")
                Next
                _backupHdWriter.WriteLine(hdHead.ToString())
                _backupHdWriter.AutoFlush = True
            Catch
                _backupHdWriter = Nothing
            End Try
        End If
    End Sub

    Private Sub CloseBackupWriters(deleteFiles As Boolean)
        For Each pair As Object() In New Object()() {
                New Object() {_backupSigWriter, _backupSigPath},
                New Object() {_backupHdWriter, _backupHdPath}}
            Dim w As IO.StreamWriter = TryCast(pair(0), IO.StreamWriter)
            Dim p As String = TryCast(pair(1), String)
            If w IsNot Nothing Then
                Try : w.Flush() : w.Close() : Catch : End Try
            End If
            If deleteFiles AndAlso p IsNot Nothing AndAlso File.Exists(p) Then
                Try : File.Delete(p) : Catch : End Try
            End If
        Next
        _backupSigWriter = Nothing
        _backupHdWriter = Nothing
    End Sub

#End Region

#Region "8 - Utilitaires"

    Private Sub ClearBuffers()
        _sigHeaders = Nothing
        _sigRows = Nothing
        _sigMeta = Nothing
        _hdBuffer = Nothing
        _hdGroups = New Integer() {}
        _hdOrders = New Integer() {}
        _xlsxPath = ""
        _csvPath = ""
    End Sub

    Private Shared Function AppFolder(subFolder As String) As String
        Dim exeDir As String = Path.GetDirectoryName(
            System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim projDir As String = Path.GetDirectoryName(Path.GetDirectoryName(exeDir))
        Dim dir As String = Path.Combine(projDir, subFolder)
        If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
        Return dir
    End Function

    Friend Sub UpdateRecordButton()
        Button_RecordPause.Enabled = _isRecording
        If _isRecording Then
            Button_Record.BackColor = Drawing.Color.FromArgb(140, 20, 20)
            Button_Record.ForeColor = Drawing.Color.FromArgb(255, 200, 200)
            Button_Record.Text = "■ STOP REC"
            Panel_FooterRecordGroup.BackColor = Drawing.Color.FromArgb(60, 18, 18)
            Label_FooterRecordTitle.ForeColor = Drawing.Color.FromArgb(220, 100, 100)
            If _csvPaused Then
                Button_RecordPause.BackColor = Drawing.Color.FromArgb(120, 90, 0)
                Button_RecordPause.ForeColor = Drawing.Color.FromArgb(255, 220, 100)
                Button_RecordPause.Text = "▶ RESUME REC"
            Else
                Button_RecordPause.BackColor = Drawing.Color.FromArgb(28, 44, 70)
                Button_RecordPause.ForeColor = Drawing.Color.FromArgb(140, 165, 200)
                Button_RecordPause.Text = "⏸ PAUSE REC"
            End If
        Else
            Button_Record.BackColor = Drawing.Color.FromArgb(30, 80, 30)
            Button_Record.ForeColor = Drawing.Color.FromArgb(160, 240, 160)
            Button_Record.Text = "⏺ REC"
            Panel_FooterRecordGroup.BackColor = Drawing.Color.FromArgb(28, 44, 70)
            Label_FooterRecordTitle.ForeColor = Drawing.Color.FromArgb(90, 130, 190)
            Button_RecordPause.BackColor = Drawing.Color.FromArgb(28, 44, 70)
            Button_RecordPause.ForeColor = Drawing.Color.FromArgb(90, 108, 136)
        End If
    End Sub

#End Region

End Class