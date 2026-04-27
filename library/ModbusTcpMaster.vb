Module ModbusTcpMaster
    Public modbustcp As New Class_ModbusTCP
    Public modbustcpisconnected As Boolean

    Public Structure TCP_Variables
        Public vline As Single      '256 Tension de ligne
        Public irms As Single       '257 Courant efficace de la charge
        Public isqburst As Single   '258 Valeur moyenne du carré du courant de charge en train d'onde
        Public isq As Single        '259 Valeur du carré du courant de charge
        Public vrms As Single       '260 Tension efficace de la charge
        Public vsq As Single        '261 Valeur du carré de la tension de charge
        Public pburst As Single     '262 Mesure puissance réelle en train d'onde
        Public P As Single          '263 Mesure de la puissance réelle
        Public S As Single          '264 Mesure de la puissance apparente
        Public PF As Single         '265 Mesure du facteur de puissance
        Public z As Single          '266 Impédance de charge
        Public frequency As Single  '267 Fréquence
        Public vsqburst As Single   '268 Valeur moyenne du carré de la tension de charge en train d'onde

        Public Ramping As Integer   '1440 Ramping = '0' // else '1'

        'Public V_Nominal As Single
        'Public I_Nominal As Single
        'Public Firing As Single
        'Public Control As Single
        'Public I_Limit As Single
        'Public I2_Transfer As Single
        'Public Xfmr As Single
        'Public Heater As Single
        'Public AI_Fct As Single
        'Public AI_Type As Single
        'Public DI1_Fct As Single
        'Public DI2_Fct As Single


    End Structure

    Public TCP_Var As TCP_Variables

    '=========================================================================
    ' OPEN TCP PORT
    ' ip   : adresse IP du thyristor ePack
    ' port : port Modbus TCP (generalement 502)
    '=========================================================================
    Public Sub OpenTCP(ip As String, port As Integer)
        modbustcp.ConnectionType = Class_ModbusTCP.E_MODBUS_TYPE.TCP_Modbus
        modbustcpisconnected = modbustcp.Open(ip, port)
    End Sub

    '=========================================================================
    ' CLOSE TCP PORT
    '=========================================================================
    Public Sub CloseTCP()
        modbustcp.Close()
        modbustcpisconnected = False
    End Sub

    '=========================================================================
    ' GET EPACK VALUES AS DICTIONARY
    ' Retourne les valeurs lues dans TCP_Var sous forme de dictionnaire
    ' SignalID -> valeur physique (apres division par ScaleFactor).
    ' A appeler apres Read_epackTCP(False).
    '=========================================================================
    Public Function GetEpackValues() As Dictionary(Of SignalID, Double)
        Dim d As New Dictionary(Of SignalID, Double)
        d(SignalID.ePack_Vline) = TCP_Var.vline
        d(SignalID.ePack_Irms) = TCP_Var.irms
        d(SignalID.ePack_Isqburst) = TCP_Var.isqburst
        d(SignalID.ePack_Isq) = TCP_Var.isq
        d(SignalID.ePack_Vrms) = TCP_Var.vrms / 10
        d(SignalID.ePack_Vsq) = TCP_Var.vsq / 10
        d(SignalID.ePack_Pburst) = TCP_Var.pburst / 10
        d(SignalID.ePack_P) = TCP_Var.P / 10
        d(SignalID.ePack_S) = TCP_Var.S / 10
        d(SignalID.ePack_PF) = TCP_Var.PF / 10
        d(SignalID.ePack_Z) = TCP_Var.z / 10
        d(SignalID.ePack_Frequency) = TCP_Var.frequency
        d(SignalID.ePack_Vsqburst) = TCP_Var.vsqburst

        d(SignalID.ePack_Ramping) = TCP_Var.Ramping
        Return d
    End Function

    '=========================================================================
    ' READ EPACK 
    '
    ' Argument : true to display Read Values on a TextBox
    '=========================================================================
    Public Sub Read_epackTCP(TextBox As Boolean)
        Dim byteresult(61) As Byte
        Dim val As Integer = 256
        Dim mybytelength As Integer = 61

        If modbustcp.Read_NChar(1, val, byteresult, mybytelength) Then

            ' Lecture des registres en une boucle
            Dim reg(12) As Integer
            For i As Integer = 0 To 12
                reg(i) = makeuint(byteresult(i * 2), byteresult(i * 2 + 1))
            Next

            TCP_Var.vline = reg(0)
            TCP_Var.irms = reg(1)
            TCP_Var.isqburst = reg(2)
            TCP_Var.isq = reg(3)
            TCP_Var.vrms = reg(4)
            TCP_Var.vsq = reg(5)
            TCP_Var.pburst = reg(6)
            TCP_Var.P = reg(7)
            TCP_Var.S = reg(8)
            TCP_Var.PF = reg(9)
            TCP_Var.z = reg(10)
            TCP_Var.frequency = reg(11)
            TCP_Var.vsqburst = reg(12)

            modbustcp.ReadInteger(1, 1440, TCP_Var.Ramping)

            If TextBox Then
                Dim result As String = ""
                For i As Integer = 0 To 12
                    result &= (val + i).ToString & "=" & reg(i).ToString & vbCrLf
                Next
                MsgBox(result)
            End If

        Else
            modbustcpisconnected = False
        End If
    End Sub


    Public Sub getbytesfrominteger(ByVal s As Integer, ByRef hh As Byte, ByRef h As Byte, ByRef l As Byte, ByRef ll As Byte)
        Dim b(3) As Byte

        b = BitConverter.GetBytes(s)
        ll = b(0)
        l = b(1)
        h = b(2)
        hh = b(3)
    End Sub

    Public Sub expandsingle(ByVal s As Single, ByRef hh As Byte, ByRef h As Byte, ByRef l As Byte, ByRef ll As Byte)
        Dim b(3) As Byte

        b = BitConverter.GetBytes(s)
        ll = b(0)
        l = b(1)
        h = b(2)
        hh = b(3)
    End Sub

    Public Function makesingle(ByVal hh As Byte, ByVal h As Byte, ByVal l As Byte, ByVal ll As Byte) As Single
        Dim sing As Single
        Dim b(3) As Byte

        b(0) = ll
        b(1) = l
        b(2) = h
        b(3) = hh
        sing = BitConverter.ToSingle(b, 0)
        Return sing
    End Function
    Public Function makesingle(ByVal hh As Integer, ByVal ll As Integer) As Single
        Dim sing As Single
        Dim b(3) As Byte
        Dim bh(1) As Byte
        Dim bl(1) As Byte
        bh = BitConverter.GetBytes(hh)
        bl = BitConverter.GetBytes(ll)

        b(0) = bl(0)
        b(1) = bl(1)
        b(2) = bh(0)
        b(3) = bh(1)
        sing = BitConverter.ToSingle(b, 0)
        Return sing
    End Function

    Public Function makeint32(ByVal hh As Byte, ByVal h As Byte, ByVal l As Byte, ByVal ll As Byte) As Integer
        Dim sing As Integer
        Dim b(3) As Byte

        b(0) = ll
        b(1) = l
        b(2) = h
        b(3) = hh
        sing = BitConverter.ToInt32(b, 0)
        Return sing
    End Function

    Public Function makelong(ByVal hh As Byte, ByVal h As Byte, ByVal l As Byte, ByVal ll As Byte) As Long
        Dim sing As Long
        Dim b(3) As Byte

        b(0) = ll
        b(1) = l
        b(2) = h
        b(3) = hh
        sing = BitConverter.ToUInt32(b, 0)
        Return sing
    End Function

    Public Function makelong(ByVal hh As Integer, ByVal ll As Integer) As Long
        Dim sing As Long
        Dim b(3) As Byte
        Dim bh(1) As Byte
        Dim bl(1) As Byte
        bh = BitConverter.GetBytes(hh)
        bl = BitConverter.GetBytes(ll)

        b(0) = bl(0)
        b(1) = bl(1)
        b(2) = bh(0)
        b(3) = bh(1)
        sing = BitConverter.ToUInt32(b, 0)
        Return sing
    End Function

    Public Function makeuint(ByVal hi As Byte, ByVal lo As Byte) As Integer
        Return CInt((hi * 256) + lo)
    End Function

    Public Function wordtobits(ByVal word As Integer) As Boolean()
        Dim tblbits(15) As Boolean
        Dim mask As Integer = 1

        For i As Integer = 0 To 15
            tblbits(i) = ((word And mask) = mask)
            mask <<= 1
        Next

        Return tblbits
    End Function

End Module