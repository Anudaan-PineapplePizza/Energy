Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Windows.Forms
Imports System.IO.Ports
Imports System.Management
Imports System.Text.RegularExpressions
Imports EasyModbus

Public Class ConnectionWindow
    ' =============================================================================
    ' CONFIG JSON — Fill Config files in the configuration combobox
    ' =============================================================================
    Private Class ConfigItem
        Public Property DisplayName As String
        Public Property FilePath As String
        Public Overrides Function ToString() As String
            Return DisplayName
        End Function
    End Class

    Private Const CONFIG_BROWSE As String = "Browse for file..."

    ' =============================================================================
    ' LOAD
    ' =============================================================================
    Private Sub ConnectionWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Baudrate
        With ComboBox_Baudrate.Items
            .Clear()
            .Add("9600") : .Add("19200") : .Add("38400")
        End With
        ComboBox_Baudrate.SelectedIndex = 2   ' 38400 default

        ' Parity
        ComboBox_Parity.Items.Clear()
        ComboBox_Parity.Items.Add(New ComboItem(Of IO.Ports.Parity) With {.Text = "None", .Value = IO.Ports.Parity.None})
        ComboBox_Parity.Items.Add(New ComboItem(Of IO.Ports.Parity) With {.Text = "Odd", .Value = IO.Ports.Parity.Odd})
        ComboBox_Parity.Items.Add(New ComboItem(Of IO.Ports.Parity) With {.Text = "Even", .Value = IO.Ports.Parity.Even})
        ComboBox_Parity.SelectedIndex = 1   ' Odd default

        ' Stop bits
        ComboBox_StopBit.Items.Clear()
        ComboBox_StopBit.Items.Add(New ComboItem(Of IO.Ports.StopBits) With {.Text = "1bit", .Value = IO.Ports.StopBits.One})
        ComboBox_StopBit.Items.Add(New ComboItem(Of IO.Ports.StopBits) With {.Text = "2bit", .Value = IO.Ports.StopBits.Two})
        ComboBox_StopBit.SelectedIndex = 0

        ' COM ports
        If ComPort.IsOpen Then ComPort.Close()
        ComPortMaster.RefreshComPorts()

        ' Config JSON combobox
        RefreshConfigList()
        CheckBox_KeepConfig.Checked = True
        ComboBox_Configuration.Enabled = False

        DisconnectUI()
        Timer_Connection.Interval = 500
        Timer_Connection.Stop()
    End Sub

    ' ─────────────────────────────────────────────────────────────────────────────
    ' BUTTON CONNECT / DISCONNECT
    ' ─────────────────────────────────────────────────────────────────────────────
    Private Sub Button_Connect_Click(sender As Object, e As EventArgs) Handles Button_Connect.Click
        If ModbusRsMaster.IsConnected() Then
            DisconnectUI()
            ModbusRsMaster.Disconnect()
            Return
        End If

        ' COM port
        Dim selected As ComPortInfo = TryCast(ComboBox_COMPort.SelectedItem, ComPortInfo)
        Dim COMport As String = If(selected IsNot Nothing, selected.Port, Nothing)
        If String.IsNullOrWhiteSpace(COMport) Then
            MessageBox.Show("No COM port selected.")
            Return
        End If

        ' Baudrate
        Dim baud As Integer
        If Not Integer.TryParse(ComboBox_Baudrate.Text, baud) Then
            MessageBox.Show("Invalid baudrate.")
            Return
        End If

        ' Device address
        Dim deviceAddressInt As Integer
        If String.IsNullOrWhiteSpace(TextBox_DeviceAddress.Text) OrElse
           Not Integer.TryParse(TextBox_DeviceAddress.Text, deviceAddressInt) OrElse
           deviceAddressInt < 1 OrElse deviceAddressInt > 247 Then
            MessageBox.Show("Invalid Modbus address (must be 1-247).")
            Return
        End If
        Dim deviceAddress As Byte = CByte(deviceAddressInt)

        ' Parity / StopBits
        Dim parityEnum As IO.Ports.Parity =
            CType(ComboBox_Parity.SelectedItem, ComboItem(Of IO.Ports.Parity)).Value
        Dim stopBitsEnum As IO.Ports.StopBits =
            CType(ComboBox_StopBit.SelectedItem, ComboItem(Of IO.Ports.StopBits)).Value

        ' Save config
        ModbusRsMaster.Config.Baudrate = baud
        ModbusRsMaster.Config.Parity = parityEnum
        ModbusRsMaster.Config.StopBit = stopBitsEnum
        ModbusRsMaster.Config.DeviceAddress = deviceAddress
        ModbusRsMaster.Config.ConnectString =
            $"{COMport},{baud},{parityEnum},{stopBitsEnum},ID={deviceAddress}"

        ' Connect
        If ModbusRsMaster.Connect(COMport, baud, parityEnum, stopBitsEnum, deviceAddress) Then
            ConnectUI()
        Else
            MessageBox.Show("Cannot open COM port.")
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────────────────────
    ' UI STATE (CONNECT / DISCONNECT)
    ' ─────────────────────────────────────────────────────────────────────────────
    Public Sub ConnectUI()
        Timer_Connection.Start()
        ModbusRsMaster.Connected = True
        Label_ConnectionStatus.BackColor = Color.FromArgb(0, 200, 100)
        Button_Start.Enabled = True
        Button_Connect.Text = "Disconnect Modbus"
        Button_Connect.BackColor = Color.FromArgb(100, 28, 28)
        MeterWindow.Label_StatusText.Text = "ONLINE"
        TextBox_DeviceAddress.Enabled = False
        ComboBox_Baudrate.Enabled = False
        ComboBox_Parity.Enabled = False
        ComboBox_StopBit.Enabled = False
        ComboBox_COMPort.Enabled = False
    End Sub

    Public Sub DisconnectUI()
        Timer_Connection.Stop()
        ModbusRsMaster.Connected = False
        Label_ConnectionStatus.BackColor = Color.FromArgb(140, 30, 30)
        MeterWindow.Label_ConnectionStatus2.BackColor = Color.FromArgb(140, 30, 30)
        Button_Start.Enabled = False
        Button_Connect.Text = "Connect Modbus"
        Button_Connect.BackColor = Color.FromArgb(28, 44, 70)
        MeterWindow.Label_StatusText.Text = "OFFLINE"
        TextBox_DeviceAddress.Enabled = True
        ComboBox_Baudrate.Enabled = True
        ComboBox_Parity.Enabled = True
        ComboBox_StopBit.Enabled = True
        ComboBox_COMPort.Enabled = True
    End Sub

    ' ─────────────────────────────────────────────────────────────────────────────
    ' BUTTON START — open MeterWindow
    ' ─────────────────────────────────────────────────────────────────────────────
    Private Sub Button_Start_Click(sender As Object, e As EventArgs) Handles Button_Start.Click
        ' Mode demo : device address = 67
        If TextBox_DeviceAddress.Text.Trim() = "67" Then
            MeterWindow.IsDemoMode = True
            ApplySelectedConfig()
            Me.Hide()
            MeterWindow.Show()
            Return
        End If

        ' Mode normal
        MeterWindow.IsDemoMode = False
        If Not ModbusRsMaster.IsConnected() Then
            MessageBox.Show("Not connected.")
            Return
        End If
        ApplySelectedConfig()
        Me.Hide()
        MeterWindow.Show()
    End Sub

    ' ─────────────────────────────────────────────────────────────────────────────
    ' TIMER — connection status blink
    ' ─────────────────────────────────────────────────────────────────────────────
    Private Sub Timer_Connection_Tick(sender As Object, e As EventArgs) Handles Timer_Connection.Tick
        If Not ModbusRsMaster.IsConnected() Then
            DisconnectUI()
            Return
        End If
        Select Case ModbusRsMaster._blinkState
            Case 0
                Label_ConnectionStatus.BackColor = Color.FromArgb(0, 200, 100)
                MeterWindow.Label_ConnectionStatus2.BackColor = Color.FromArgb(0, 200, 100)
                ModbusRsMaster._blinkState = 1
            Case 1
                Label_ConnectionStatus.BackColor = Color.FromArgb(40, 80, 40)
                MeterWindow.Label_ConnectionStatus2.BackColor = Color.FromArgb(40, 80, 40)
                ModbusRsMaster._blinkState = 0
        End Select
        ' Synchroniser l etat du bouton TCP avec la connexion reelle
        If ModbusTcpMaster.modbustcpisconnected Then
            Button_Connect_TCP.BackColor = Color.FromArgb(0, 110, 70)
            Button_Connect_TCP.Text = "Disconn. TCP"
        Else
            Button_Connect_TCP.BackColor = Color.FromArgb(70, 28, 28)
            Button_Connect_TCP.Text = "Conn. TCP"
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────────────────────
    ' COM PORT REFRESH
    ' ─────────────────────────────────────────────────────────────────────────────
    Private Sub ComboBox_COMPort_MouseClick(sender As Object, e As MouseEventArgs) _
            Handles ComboBox_COMPort.MouseClick
        RefreshComPorts()
    End Sub

    ' DEVICE ADDRESS — digits only + mode demo si valeur = 67
    Private Sub TextBox_DeviceAddress_KeyPress(sender As Object, e As KeyPressEventArgs) _
            Handles TextBox_DeviceAddress.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox_DeviceAddress_TextChanged(sender As Object, e As EventArgs) _
            Handles TextBox_DeviceAddress.TextChanged
        If TextBox_DeviceAddress.Text.Trim() = "67" Then
            ' Mode demo : deverrouille Start sans connexion Modbus
            Button_Start.Enabled = True
            Button_Start.BackColor = Color.FromArgb(130, 85, 0)
            Button_Start.Text = "Start (DEMO)"
        Else
            ' Comportement normal : Start uniquement si connecte
            Button_Start.Enabled = ModbusRsMaster.IsConnected()
            Button_Start.BackColor = If(ModbusRsMaster.IsConnected(),
                Color.FromArgb(0, 130, 80), Color.FromArgb(28, 44, 70))
            Button_Start.Text = "Start"
        End If
    End Sub

    ' =============================================================================
    ' JSON CONFIG SUBs / FUNCTION
    ' =============================================================================

    ' ─────────────────────────────────────────────────────────────────────────────
    ' CONFIG LIST — scan exe directory for *.json files
    ' ─────────────────────────────────────────────────────────────────────────────
    Private Sub RefreshConfigList()
        ComboBox_Configuration.Items.Clear()

        Dim configDir As String = AppFolder("config")
        If Directory.Exists(configDir) Then
            For Each f As String In Directory.GetFiles(configDir, "*.json")
                ComboBox_Configuration.Items.Add(New ConfigItem With {
                    .DisplayName = Path.GetFileNameWithoutExtension(f),
                    .FilePath = f
                })
            Next
        End If

        ' Always add "Browse..." at the end
        ComboBox_Configuration.Items.Add(New ConfigItem With {
            .DisplayName = CONFIG_BROWSE,
            .FilePath = ""
        })

        If ComboBox_Configuration.Items.Count > 1 Then
            ComboBox_Configuration.SelectedIndex = 0
        End If
    End Sub

    Private Sub CheckBox_KeepConfig_CheckedChanged(sender As Object, e As EventArgs) _
            Handles CheckBox_KeepConfig.CheckedChanged
        ComboBox_Configuration.Enabled = Not CheckBox_KeepConfig.Checked
        If Not CheckBox_KeepConfig.Checked Then RefreshConfigList()
    End Sub

    Private Sub ComboBox_Configuration_SelectedIndexChanged(sender As Object, e As EventArgs) _
            Handles ComboBox_Configuration.SelectedIndexChanged
        Dim item As ConfigItem = TryCast(ComboBox_Configuration.SelectedItem, ConfigItem)
        If item Is Nothing Then Return
        If item.DisplayName = CONFIG_BROWSE Then
            ' Open file picker
            Dim jsonFile As New OpenFileDialog()
            jsonFile.Title = "Select a JSON configuration file"
            jsonFile.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            jsonFile.InitialDirectory = AppFolder("config")
            If jsonFile.ShowDialog() = DialogResult.OK Then
                ' Add file to list and select it
                Dim newItem As New ConfigItem With {
                    .DisplayName = Path.GetFileNameWithoutExtension(jsonFile.FileName),
                    .FilePath = jsonFile.FileName
                }
                ' Insert before Browse
                ComboBox_Configuration.Items.Insert(
                    ComboBox_Configuration.Items.Count - 1, newItem)
                ComboBox_Configuration.SelectedItem = newItem
            Else
                ' Revert to first item if available
                If ComboBox_Configuration.Items.Count > 1 Then
                    ComboBox_Configuration.SelectedIndex = 0
                End If
            End If
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────────────────────
    ' Load the selected JSON config into UserPreferences
    ' ─────────────────────────────────────────────────────────────────────────────
    Private Sub ApplySelectedConfig()
        If CheckBox_KeepConfig.Checked Then Return   ' keep current prefs
        Dim item As ConfigItem = TryCast(ComboBox_Configuration.SelectedItem, ConfigItem)
        If item Is Nothing OrElse item.DisplayName = CONFIG_BROWSE Then Return
        If Not File.Exists(item.FilePath) Then
            MessageBox.Show("Config file not found:" & Environment.NewLine & item.FilePath,
                "Config error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Try
            UserPreferences.Instance.LoadFromFile(item.FilePath)
        Catch ex As Exception
            MessageBox.Show("Error loading config:" & Environment.NewLine & ex.Message,
                "Config error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    ' ─────────────────────────────────────────────────────────────────────────────
    ' Config scan Directory re localisation for JSON :
    '
    ' (goes up 2 levels from bin\Debug\ to project root)
    ' ─────────────────────────────────────────────────────────────────────────────
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

    Private Sub Button_ConnectTCP(sender As Object, e As EventArgs) Handles Button_Connect_TCP.Click
        Dim p As UserPreferences = UserPreferences.Instance
        If ModbusTcpMaster.modbustcpisconnected Then
            ' Deja connecte -> deconnecter
            ModbusTcpMaster.CloseTCP()
            RegisterMap.Refresh()
            p.EpackEnabled = False
            p.Save()
            Button_Connect_TCP.BackColor = Color.FromArgb(70, 28, 28)
            Button_Connect_TCP.Text = "Conn. TCP"
        Else
            ' Non connecte -> connecter
            ModbusTcpMaster.OpenTCP(p.EpackIP, p.EpackPort)
            If ModbusTcpMaster.modbustcpisconnected Then
                RegisterMap.Refresh()
                p.EpackEnabled = True
                ' Activer tous les signaux ePack dans panel et CSV
                For Each r As RegisterDef In RegisterMap.GetRealTimeSignals()
                    If r.Group = SignalGroup.ePack Then
                        p.SetPanelVisible(r.ID, True)
                        p.SetCsvEnabled(r.ID, True)
                        p.SetEnabled(r.ID, True)
                    End If
                Next
                p.Save()
                Button_Connect_TCP.BackColor = Color.FromArgb(0, 110, 70)
                Button_Connect_TCP.Text = "Disconn. TCP"
            Else
                Button_Connect_TCP.BackColor = Color.FromArgb(70, 28, 28)
                Button_Connect_TCP.Text = "Conn. TCP"
            End If
        End If
    End Sub

    Private Sub Button_Read_epack(sender As Object, e As EventArgs) Handles Button_ReadTCP.Click
        Read_epackTCP(True)
    End Sub

End Class