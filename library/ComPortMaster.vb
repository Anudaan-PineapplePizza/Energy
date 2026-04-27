Option Strict Off
Option Explicit On

Imports System.IO.Ports
Imports System.Management
Imports System.Text.RegularExpressions
Imports Microsoft.Win32

'=========================================================================
' COM PORT INFO CLASS
'=========================================================================
Public Class ComPortInfo
    Public Property Port As String
    Public Property Caption As String

    Public Sub New(port As String, caption As String)
        Me.Port = port
        Me.Caption = caption
    End Sub

    Public Overrides Function ToString() As String
        Dim friendly As String = Caption
        If Not String.IsNullOrEmpty(friendly) AndAlso friendly <> Port Then
            friendly = friendly.Replace("(" & Port & ")", "").Trim()
            Return Port & "  -  " & friendly
        End If
        Return Port
    End Function
End Class

'=========================================================================
' COM PORT MASTER MODULE
'=========================================================================
Public Module ComPortMaster

    ' -----------------------------------------------------------------------
    ' COUCHE 1 : WMI (Win32_PnPEntity)
    ' Meilleure source - noms complets avec fabricant
    ' Peut echouer si System.Management.dll absent ou GPO restrictive
    ' -----------------------------------------------------------------------
    Private Function TryGetNamesViaWmi() As Dictionary(Of String, String)
        Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Try
            Using searcher As New ManagementObjectSearcher(
                    "SELECT Name FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'")
                For Each obj As ManagementObject In searcher.Get()
                    Dim name As String = TryCast(obj("Name"), String)
                    If String.IsNullOrEmpty(name) Then Continue For
                    Dim m As Match = Regex.Match(name, "\((COM\d+)\)", RegexOptions.IgnoreCase)
                    If m.Success Then
                        map(m.Groups(1).Value.ToUpperInvariant()) = name
                    End If
                Next
            End Using
        Catch
            ' WMI indisponible ou GPO restrictive -> retourner dict vide
        End Try
        Return map
    End Function

    ' -----------------------------------------------------------------------
    ' COUCHE 2 : Registry (HKLM\HARDWARE\DEVICEMAP\SERIALCOMM)
    ' Tres fiable, presente sur tous les Windows depuis XP
    ' Donne les noms de port mais pas toujours
    ' -----------------------------------------------------------------------
    Private Function TryGetPortsViaRegistry() As List(Of String)
        Dim ports As New List(Of String)()
        Try
            Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(
                    "HARDWARE\DEVICEMAP\SERIALCOMM", False)
                If key Is Nothing Then Return ports
                For Each valueName As String In key.GetValueNames()
                    Dim port As String = TryCast(key.GetValue(valueName), String)
                    If Not String.IsNullOrEmpty(port) Then
                        ports.Add(port.ToUpperInvariant())
                    End If
                Next
            End Using
        Catch
        End Try
        Return ports
    End Function

    ' -----------------------------------------------------------------------
    ' COUCHE 3 : SerialPort.GetPortNames()
    ' Universel mais peut rater les ports virtuels sur certains Windows
    ' -----------------------------------------------------------------------
    Private Function TryGetPortsViaSerialPort() As List(Of String)
        Dim ports As New List(Of String)()
        Try
            For Each p As String In SerialPort.GetPortNames()
                ports.Add(p.ToUpperInvariant())
            Next
        Catch
        End Try
        Return ports
    End Function

    ' -----------------------------------------------------------------------
    ' FUSION : union des 3 couches + noms conviviaux depuis WMI
    ' -----------------------------------------------------------------------
    Public Function GetComPortInfoList() As List(Of ComPortInfo)
        ' 1. Noms conviviaux via WMI
        Dim wmiNames As Dictionary(Of String, String) = TryGetNamesViaWmi()

        ' 2. Liste des ports : union Registry + SerialPort (evite les doublons)
        Dim allPorts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each p As String In TryGetPortsViaRegistry()
            allPorts.Add(p)
        Next
        For Each p As String In TryGetPortsViaSerialPort()
            allPorts.Add(p)
        Next
        ' Ajouter aussi les ports trouves par WMI (au cas ou les 2 autres ont rate)
        For Each p As String In wmiNames.Keys
            allPorts.Add(p)
        Next

        ' 3. Construire la liste finale
        Dim list As New List(Of ComPortInfo)()
        For Each port As String In allPorts
            Dim caption As String = Nothing
            If Not wmiNames.TryGetValue(port, caption) Then
                caption = port  ' pas de nom convivial disponible
            End If
            list.Add(New ComPortInfo(port, caption))
        Next

        ' 4. Trier par numero de port (COM1, COM2, COM10...)
        list.Sort(Function(a, b)
                      Dim na As Integer = 0
                      Dim nb As Integer = 0
                      Integer.TryParse(Regex.Replace(a.Port, "[^\d]", ""), na)
                      Integer.TryParse(Regex.Replace(b.Port, "[^\d]", ""), nb)
                      Return na.CompareTo(nb)
                  End Function)
        Return list
    End Function

    ' -----------------------------------------------------------------------
    ' REFRESH : alimente le ComboBox de ConnectionWindow
    ' -----------------------------------------------------------------------
    Public Sub RefreshComPorts()
        ConnectionWindow.ComboBox_COMPort.Items.Clear()

        Dim items As List(Of ComPortInfo) = GetComPortInfoList()
        For Each info As ComPortInfo In items
            ConnectionWindow.ComboBox_COMPort.Items.Add(info)
        Next

        If ConnectionWindow.ComboBox_COMPort.Items.Count > 0 Then
            ConnectionWindow.ComboBox_COMPort.SelectedIndex = 0
        End If
    End Sub

End Module