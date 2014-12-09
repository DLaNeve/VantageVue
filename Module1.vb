﻿Imports System.Collections.Specialized
Module VantageVue
    Private Property GetBarometerData_V As Single

    Sub Main()
        Dim msgs As StringCollection = New StringCollection
        Open_USB_Port()
        Do Until True = False
            LoadCurrentVantageData_V()
            msgs.Add(Create_Wind_msg())
            msgs.Add(Create_Temp_msg())
            msgs.Add(Create_Bar_msg())
            msgs.Add(Create_Humidity_msg())
            Broadcast_MSG(msgs)
            msgs.Clear()
            System.Threading.Thread.Sleep(3000)
        Loop

    End Sub

    Sub Open_USB_Port()
        Dim usbSerialNumber As UInteger
        Dim usbserialnumbertextbox As String
        Dim portWorking As Boolean
        CloseUSBPort_V()
        usbSerialNumber = GetUSBDevSerialNumber_V()
        usbserialnumbertextbox = usbSerialNumber.ToString
        If (OpenUSBPort_V(usbSerialNumber) = 0) Then
            If (InitStation_V() = COM_ERROR) Then
                MsgBox("Cannot open the USB port")
                portWorking = False
            Else
                SetCommTimeoutVal_V(100, 100)
                portWorking = True
            End If
        Else
            MsgBox("Cannot open the USB port")
            portWorking = False
        End If
    End Sub

    Function Create_Wind_msg() As String
        Dim windspeed As Single
        Dim windDir As Short
        Dim windmsg As String

        windspeed = GetWindSpeed_V()
        windDir = GetWindDir_V()
        windmsg = "$WIVWR," & windDir & ",R," & windspeed & ",N,,,," & vbCrLf
        Console.WriteLine(windmsg)
        Create_Wind_msg = windmsg
    End Function

    Function Create_Temp_msg() As String
        Dim temp As Single
        Dim TempMsg As String

        '$IIXDR,P,1.02481,B,Barometer*0D
        '$IIXDR,C,19.52,C,TempAir*3D
        '$IIMTA,19.52,C*1E


        temp = GetInsideTemp_V()
        TempMsg = "$WIMTA," & temp & ",C" & vbCrLf
        'TempMsg = "$WIXDR" & ",c," & temp & ",C1C" & vbCrLf
        Console.WriteLine(TempMsg)
        Create_Temp_msg = TempMsg
    End Function
    Function Create_Bar_msg() As String
        Dim bar As Short
        Dim barMsg As String

        bar = GetBarometer_V()
        barMsg = "$WIMMB," & bar & ",,,B" & vbCrLf
        Console.WriteLine(barMsg)
        Create_Bar_msg = barMsg
    End Function
    Function Create_Humidity_msg() As String
        Dim Humidity As Single
        Dim Humidity_msg As String

        Humidity = GetInsideHumidity_V()
        Humidity_msg = "$WIMHU," & Humidity & vbCrLf
        Console.WriteLine(Humidity_msg)
        Create_Humidity_msg = Humidity_msg
    End Function
    Sub Broadcast_MSG(msgs As StringCollection)
        ' Dim broadcast As IPAddress = IPAddress.Parse("192.168.0.37")
        ' Dim ep As New IPEndPoint(broadcast, 10110)
        Dim ep As New IPEndPoint(IPAddress.Loopback, 10110)
        Dim socket As New Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp)
        socket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, True)
        ' socket.Bind(ep)
        Dim sendbuf As Byte() = Encoding.ASCII.GetBytes(msgs(0))
        For Each msg As String In msgs
            sendbuf = Encoding.ASCII.GetBytes(msg)
            socket.SendTo(sendbuf, ep)
        Next
    End Sub

End Module