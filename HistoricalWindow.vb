Option Strict On
Option Explicit On

Imports System.IO

' =============================================================================
' HistoricalWindow.vb  -- Partial Class (1/6)
' =============================================================================
Public Class HistoricalWindow
    Inherits System.Windows.Forms.Form

    Friend _csvData As CsvData = Nothing

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub HistoricalWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetChartModeZoom()

    End Sub

    Private Sub Button_Open_Click(sender As Object, e As EventArgs) Handles Button_Open.Click
        Dim dlg As New OpenFileDialog()
        dlg.Title = "Open CSV recording"
        dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        dlg.InitialDirectory = AppFolder("csv")
        If dlg.ShowDialog() <> DialogResult.OK Then Return

        Dim data As CsvData = LoadCsv(dlg.FileName)
        If data Is Nothing Then Return
        _csvData = data

        Me.Text = "Historical Viewer  --  " & Path.GetFileName(dlg.FileName)
        If data.RecordDate <> DateTime.MinValue Then
            Me.Text &= "   (" & data.RecordDate.ToString("yyyy-MM-dd  HH:mm") & ")"
        End If

        BuildChart(data)
        BuildValuePanel(data)
        BuildHdPanel(data)

        Label_CursorTime.Text = "t = --"
        SetChartModeZoom()
    End Sub

    Private Sub Button_Zoom_Click(sender As Object, e As EventArgs) Handles Button_Zoom.Click
        SetChartModeZoom()
    End Sub

    Private Sub Button_Cursor_Click(sender As Object, e As EventArgs) Handles Button_Cursor.Click
        SetChartModeCursor()
    End Sub

    'Private Shared Function AppFolder(subFolder As String) As String
    '    Dim exeDir As String = IO.Path.GetDirectoryName(
    '        System.Reflection.Assembly.GetExecutingAssembly().Location)
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