Option Explicit On
Option Strict On

Imports System.Net
Imports System.Net.Sockets
Imports System.Threading


Public Class clsSocketData
    Protected buffer As Byte()
    Protected sock As Socket
    Protected endPoint As IPEndPoint

    Public Sub New(ByVal socket As Socket)
        sock = socket
        buffer = New Byte(2048 - 1) {}
        endPoint = New IPEndPoint(IPAddress.Any, 0)
    End Sub
    Public Sub New(ByVal socket As Socket, ByVal point As IPEndPoint)
        sock = socket
        buffer = New Byte(2048 - 1) {}
        endPoint = point
    End Sub
    Public Sub New(ByVal socket As Socket, ByVal data As Byte(), ByVal point As IPEndPoint)
        sock = socket
        buffer = data
        endPoint = point
    End Sub

    Public Function GetBuffer() As Byte()
        Return buffer
    End Function
    Public Function GetSocket() As Socket
        Return sock
    End Function
    Public Function GetEndPoint() As IPEndPoint
        Return endPoint
    End Function
End Class

Public MustInherit Class clsSocket
    Public Enum SocketEvent
        ConnectionEstablished
        ConnectionAccepted
        ConnectionFailed
        ConnectionClosedByPeer
        SocketDispose
        DataArrival
    End Enum

    Private socketDisposed As Boolean

    Public Event eventMessage(ByVal socketEvent As SocketEvent, ByVal socket As clsSocket)
    Public Event eventError(ByVal errorFunction As String, ByVal errorString As String, ByVal socket As clsSocket)

    Protected Sub New()
        socketDisposed = False
    End Sub

    Public Function IsDisposed() As Boolean
        Return socketDisposed
    End Function

    Public Shared Function ConvertIP(ByVal ip As String) As Byte()
        Dim octets As String()
        Try
            octets = ip.Split(Convert.ToChar("."))
            If octets.Length = 4 AndAlso IsNumeric(octets(0)) AndAlso IsNumeric(octets(1)) AndAlso IsNumeric(octets(2)) AndAlso IsNumeric(octets(3)) Then
                Return New Byte() {CType(octets(0), Byte), CType(octets(1), Byte), CType(octets(2), Byte), CType(octets(3), Byte)}
            End If
            Return New Byte() {}
        Catch ex As Exception
            Debug.WriteLine(ex)
            Return New Byte() {}
        End Try
    End Function
    Public Shared Function GetFirstIP(ByVal name As String) As String
        Dim IPs As String()
        Try
            IPs = GetIP(name)
            If IPs.Length > 0 Then
                Return IPs(0)
            End If
            Return ""
        Catch ex As Exception
            Debug.WriteLine(ex)
            Return ""
        End Try
    End Function

    Public Shared Function GetIP() As String()
        Return GetIP(Environment.MachineName)
    End Function
    Public Shared Function GetIP(ByVal name As String) As String()
        Dim list As ArrayList
        Dim IP As IPAddress
        Try
            list = New ArrayList
            For Each IP In Dns.GetHostAddresses(name)
                list.Add(IP.ToString)
            Next
            Return CType(list.ToArray(GetType(String)), String())
        Catch ex As Exception
            Debug.WriteLine(ex)
            Return New String() {}
        End Try
    End Function


    Protected Sub RaiseEventMessage(ByVal socketEvent As SocketEvent, ByVal socket As clsSocket)
        If IsDisposed() = False Then
            RaiseEvent eventMessage(socketEvent, socket)
        End If

        If socketEvent = clsSocket.SocketEvent.SocketDispose Then
            socketDisposed = True
        End If
    End Sub
    Protected Sub RaiseEventError(ByVal errorFunction As String, ByVal errorString As String, ByVal socket As clsSocket)
        If IsDisposed() = False Then
            RaiseEvent eventError(errorFunction, errorString, socket)
        End If

    End Sub
End Class

#Region "UDP/IP"
Public Class clsSocketUDP
    Inherits clsSocket

    Private sock As Socket

    Private bufferReceiveQueue As Queue
    Private bufferSendQueue As Queue
    Private receiveSockData As clsSocketData

    Public Sub New()

        sock = New Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp)
        sock.Bind(New IPEndPoint(IPAddress.Any, 0))
        sock.Blocking = True
        sock.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.Broadcast, True)

        bufferReceiveQueue = Queue.Synchronized(New Queue)
        bufferSendQueue = Queue.Synchronized(New Queue)

        receiveSockData = New clsSocketData(sock)
    End Sub
    Public Sub New(ByVal port As Integer)

        sock = New Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp)
        sock.Bind(New IPEndPoint(IPAddress.Any, port))
        sock.Blocking = True
        sock.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.Broadcast, True)

        bufferReceiveQueue = Queue.Synchronized(New Queue)
        bufferSendQueue = Queue.Synchronized(New Queue)

        receiveSockData = New clsSocketData(sock, New IPEndPoint(IPAddress.Any, port))
        Receive()
    End Sub
    Public Function Dispose() As Boolean
        Try
            sock.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function Send(ByVal IP As String, ByVal port As Integer, ByVal data As Byte()) As Boolean
        Dim sockData As clsSocketData
        Dim thread As Thread
        Try
            sockData = New clsSocketData(sock, data, New IPEndPoint(IPAddress.Parse(IP), port))
            bufferSendQueue.Enqueue(sockData)
            thread = New Thread(AddressOf SendThread)
            thread.Start()

            Return True
        Catch ex As Exception
            RaiseEventError("Send", ex.ToString, Me)
            Return False
        End Try
    End Function
    Private Sub SendThread()
        Dim sockData As clsSocketData
        Dim mutexSend As Mutex

        mutexSend = New Mutex(False, String.Format("mutexsend-{0}", sock.GetHashCode)) 'lock when sock is sending
        mutexSend.WaitOne()
        Try
            If bufferSendQueue.Count > 0 Then
                sockData = CType(bufferSendQueue.Dequeue, clsSocketData)
                sock.SendTo(sockData.GetBuffer, 0, sockData.GetBuffer().Length, SocketFlags.None, sockData.GetEndPoint)
            End If
        Catch ex As Exception
            RaiseEventError("SendThread", ex.ToString, Me)
        Finally
            mutexSend.ReleaseMutex()
        End Try
    End Sub

    Private Function Receive() As Boolean
        Dim thread As Thread
        Try
            thread = New Thread(AddressOf ReceiveThread)
            thread.Start()

            Return True
        Catch ex As Exception
            RaiseEventError("Receive", ex.ToString, Me)
            Return False
        End Try
    End Function
    Private Sub ReceiveThread()
        Dim sockData As clsSocketData
        Dim receiveSize As Integer
        Dim i As Integer
        Try
            While (Not IsDisposed())
                sockData = New clsSocketData(sock)
                receiveSize = sock.ReceiveFrom(sockData.GetBuffer(), sockData.GetBuffer().Length, SocketFlags.None, receiveSockData.GetEndPoint)

                If receiveSize > 0 Then
                    SyncLock bufferReceiveQueue.SyncRoot()
                        For i = 0 To receiveSize - 1
                            bufferReceiveQueue.Enqueue(sockData.GetBuffer()(i))
                        Next
                    End SyncLock
                    RaiseEventMessage(SocketEvent.DataArrival, Me)
                Else
                    RaiseEventMessage(SocketEvent.ConnectionClosedByPeer, Me)
                    Exit While
                End If
            End While
        Catch ex As Exception
            RaiseEventError("ReceiveThread", ex.ToString, Me)
        End Try
    End Sub


End Class
#End Region

#Region "TCP/IP"
Public Class clsSocketTCPClient
    Inherits clsSocket

    Private client As Socket
    Private bufferReceiveQueue As Queue
    Private bufferSendQueue As Queue
    Private connectSockData As clsSocketData


    Public Sub New()
        client = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        client.Blocking = True

        bufferReceiveQueue = Queue.Synchronized(New Queue)
        bufferSendQueue = Queue.Synchronized(New Queue)

        connectSockData = New clsSocketData(client)
    End Sub

    Public Function GetLocalIP() As Byte()
        Try
            If IsConnected() Then
                Return CType(client.LocalEndPoint, IPEndPoint).Address.GetAddressBytes
            End If
            Return New Byte() {}
        Catch ex As Exception
            Return New Byte() {}
        End Try
    End Function
    Public Function GetRemoteIP() As Byte()
        Try
            If IsConnected() Then
                Return CType(client.RemoteEndPoint, IPEndPoint).Address.GetAddressBytes
            End If
            Return New Byte() {}
        Catch ex As Exception
            Return New Byte() {}
        End Try

    End Function
    Public Function IsConnected() As Boolean
        Try
            Return client.Connected
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function GetReceiveQueue() As Queue
        Return bufferReceiveQueue
    End Function
    Public Function GetSendQueue() As Queue
        Return bufferSendQueue
    End Function

    Public Function Dispose() As Boolean
        Try
            client.Close()

            Return True
        Catch ex As Exception
            Return False
        Finally
            RaiseEventMessage(SocketEvent.SocketDispose, Me)
        End Try
    End Function
    Public Function Accept(ByVal sock As Socket) As clsSocketTCPClient
        Try
            client = sock
            If client.Connected = True Then
                RaiseEventMessage(SocketEvent.ConnectionEstablished, Me)
                Receive()
            Else
                RaiseEventMessage(SocketEvent.ConnectionFailed, Me)
            End If
            Return Me
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function Receive() As Boolean
        Dim thread As Thread
        Try
            thread = New Thread(AddressOf ReceiveThread)
            thread.Start()

            Return True
        Catch ex As Exception
            RaiseEventError("Receive", ex.ToString, Me)
            Return False
        End Try
    End Function
    Private Sub ReceiveThread()
        Dim sockData As clsSocketData
        Dim receiveSize As Integer
        Dim i As Integer
        Try

            While (Not IsDisposed())
                sockData = New clsSocketData(client)
                receiveSize = client.Receive(sockData.GetBuffer(), sockData.GetBuffer().Length, SocketFlags.None)

                If receiveSize > 0 Then
                    SyncLock bufferReceiveQueue.SyncRoot()
                        For i = 0 To receiveSize - 1
                            bufferReceiveQueue.Enqueue(sockData.GetBuffer()(i))
                        Next
                    End SyncLock
                    RaiseEventMessage(SocketEvent.DataArrival, Me)
                Else
                    RaiseEventMessage(SocketEvent.ConnectionClosedByPeer, Me)
                    Exit While
                End If
            End While
        Catch ex As Exception
            RaiseEventError("ReceiveThread", ex.ToString, Me)
        End Try
    End Sub

    Public Function Connect(ByVal IP As String, ByVal port As Integer) As Boolean
        Dim thread As Thread
        Try
            If IP.Length > 0 Then
                connectSockData = New clsSocketData(client, New IPEndPoint(IPAddress.Parse(IP), port))

                thread = New Thread(AddressOf ConnectThread)
                thread.Start()
                Return True
            End If
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Sub ConnectThread()
        Try
            client = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            client.Blocking = True
            client.Connect(connectSockData.GetEndPoint)

            If client.Connected = True Then
                RaiseEventMessage(SocketEvent.ConnectionEstablished, Me)
                Receive()
            Else
                RaiseEventMessage(SocketEvent.ConnectionFailed, Me)
            End If
        Catch ex As Exception
            RaiseEventError("ConnectThread", ex.ToString, Me)
        End Try
    End Sub

    Public Function Send(ByVal data As String) As Boolean
        Return Send(New System.Text.ASCIIEncoding().GetBytes(data))
    End Function
    Public Function Send(ByVal data As Byte()) As Boolean
        Dim thread As Thread
        Try
            bufferSendQueue.Enqueue(data)
            thread = New Thread(AddressOf SendThread)
            thread.Start()

            Return True
        Catch ex As Exception
            RaiseEventError("Send", ex.ToString, Me)
            Return False
        End Try
    End Function
    Private Sub SendThread()
        Dim data As Byte()
        Dim mutexSend As Mutex

        mutexSend = New Mutex(False, String.Format("mutexsend-{0}", client.GetHashCode)) 'lock when sock is sending
        mutexSend.WaitOne()
        Try
            If client.Connected = True AndAlso bufferSendQueue.Count > 0 Then
                data = CType(bufferSendQueue.Dequeue, Byte())
                client.Send(data, data.Length, SocketFlags.None)
            End If
        Catch ex As Exception
            RaiseEventError("SendThread", ex.ToString, Me)
        Finally
            mutexSend.ReleaseMutex()
        End Try
    End Sub


End Class

Public Class clsSocketTCPServer
    Inherits clsSocket

    Private server As Socket


    Public Sub New()
        server = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        server.Blocking = True
    End Sub

    Public Function GetListeningPort() As Integer
        Try
            If IsListening() Then
                Return CType(server.LocalEndPoint, IPEndPoint).Port
            End If
            Return 0
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Public Function IsListening() As Boolean
        Return server.IsBound
    End Function

    Public Function Dispose() As Boolean
        Try
            Debug.WriteLine("before server.close")
            server.Close()
            Debug.WriteLine("after server.close")
            Return True
        Catch ex As Exception
            Return False
        Finally
            RaiseEventMessage(SocketEvent.SocketDispose, Me)
        End Try
    End Function

    Public Function Listen(ByVal port As Integer) As Boolean
        Return Listen("", port)
    End Function
    Public Function Listen(ByVal ip As String, ByVal port As Integer) As Boolean
        Dim thread As Thread
        Try
            server = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            server.Blocking = True
            If ip.Length = 0 Then
                server.Bind(New IPEndPoint(IPAddress.Any, port))
            Else
                server.Bind(New IPEndPoint(IPAddress.Parse(ip), port))
            End If
            server.Listen(5)

            thread = New Thread(AddressOf AcceptThread)
            thread.Start()

            Return True
        Catch ex As Exception
            RaiseEventError("Listen", ex.ToString, Me)
            Return False
        End Try
    End Function

    Private Sub AcceptThread()
        Dim client As clsSocketTCPClient
        Try
            While (Not IsDisposed())
                client = New clsSocketTCPClient
                client.Accept(server.Accept())

                If client.IsConnected Then
                    RaiseEventMessage(clsSocket.SocketEvent.ConnectionAccepted, client)
                End If
            End While
        Catch ex As Exception
            RaiseEventError("AcceptThread", ex.ToString, Me)
        End Try
    End Sub

End Class
#End Region











































'