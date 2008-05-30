Option Explicit On 
Option Strict On

Imports System.IO
Imports LainBnetCore
Imports LainHelper

Public Class frmLainEthLite
    Inherits System.Windows.Forms.Form

    Friend WithEvents menuLENP As System.Windows.Forms.MenuItem
    Friend WithEvents menuLENPServer As System.Windows.Forms.MenuItem
    Friend WithEvents menuLENPClient As System.Windows.Forms.MenuItem
    Friend WithEvents menuConvertor As System.Windows.Forms.MenuItem

    Private Const GAME_PRIVATE As Integer = 17
    Private Const GAME_PUBLIC As Integer = 16

    Public Const ProjectLainVersion As String = "Battle.net Game Host v0.11 beta"
    Public Const ProjectLainName As String = "LainEthLite"
    Public Const ProjectLainConfig As String = "LainEthLiteConfiguration"
    Public Const ProjectLainMap As String = "LainEthMap"
    Public Const ProjectLainUser As String = "LainEthLiteUser"
    Public Const ProjectLainCustomData As String = "LainEthLiteCustomData"

    Private WithEvents bnet As clsBNET
    Private WithEvents channel As clsBNETChannel
    Private WithEvents bot As clsBotCommandHostChannel

    Private WithEvents aliveTimer As Timers.Timer
    Private aliveCounter As Integer
    Private channelJoinCounter As Integer

    Private listHost As ArrayList
    Private currentHost As clsGameHost
    Private map As clsGameHostMap
    Private customData As clsBNETCustomData

    Private formLENPServer As frmLENPServer
    Private lenp As clsLENPServer


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents menuMain As System.Windows.Forms.MainMenu
    Friend WithEvents contextMain As System.Windows.Forms.ContextMenu
    Friend WithEvents trayIcon As System.Windows.Forms.NotifyIcon
    Friend WithEvents menuFile As System.Windows.Forms.MenuItem
    Friend WithEvents menuExit As System.Windows.Forms.MenuItem
    Friend WithEvents labelStatus As System.Windows.Forms.Label
    Friend WithEvents contextShow As System.Windows.Forms.MenuItem
    Friend WithEvents contextExit As System.Windows.Forms.MenuItem
    Public WithEvents txtChatLog As System.Windows.Forms.TextBox
    Friend WithEvents txtChat As System.Windows.Forms.TextBox
    Friend WithEvents buttonChat As System.Windows.Forms.Button
    Friend WithEvents buttonGo As System.Windows.Forms.Button
    Friend WithEvents labelChannel As System.Windows.Forms.Label
    Friend WithEvents comboRealm As System.Windows.Forms.ComboBox
    Friend WithEvents txtChannel As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtAccount As System.Windows.Forms.TextBox
    Friend WithEvents txtTFT As System.Windows.Forms.TextBox
    Friend WithEvents txtROC As System.Windows.Forms.TextBox
    Friend WithEvents groupParam As System.Windows.Forms.GroupBox
    Friend WithEvents labelSummary As System.Windows.Forms.Label
    Friend WithEvents labelBotServer As System.Windows.Forms.Label
    Friend WithEvents labelFirstChannel As System.Windows.Forms.Label
    Friend WithEvents labelPassword As System.Windows.Forms.Label
    Friend WithEvents labelUserName As System.Windows.Forms.Label
    Friend WithEvents labelTFTKey As System.Windows.Forms.Label
    Friend WithEvents labelROCKey As System.Windows.Forms.Label
    Friend WithEvents labelRealm As System.Windows.Forms.Label
    Friend WithEvents txtWar3 As System.Windows.Forms.TextBox
    Friend WithEvents listChannelPeople As System.Windows.Forms.ListBox
    Friend WithEvents menuAutoReconnect As System.Windows.Forms.MenuItem
    Friend WithEvents labelHostPort As System.Windows.Forms.Label
    Friend WithEvents txtHostPort As System.Windows.Forms.TextBox
    Friend WithEvents groupUser As System.Windows.Forms.GroupBox
    Friend WithEvents listUser As System.Windows.Forms.ListBox
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents buttonUserAdd As System.Windows.Forms.Button
    Friend WithEvents buttonUserRemove As System.Windows.Forms.Button
    Friend WithEvents buttonGameStop As System.Windows.Forms.Button
    Friend WithEvents listGame As System.Windows.Forms.ListBox
    Friend WithEvents buttonGameInfo As System.Windows.Forms.Button
    Friend WithEvents groupGame As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLainEthLite))
        Me.menuMain = New System.Windows.Forms.MainMenu(Me.components)
        Me.menuFile = New System.Windows.Forms.MenuItem
        Me.menuAutoReconnect = New System.Windows.Forms.MenuItem
        Me.menuConvertor = New System.Windows.Forms.MenuItem
        Me.menuExit = New System.Windows.Forms.MenuItem
        Me.menuLENP = New System.Windows.Forms.MenuItem
        Me.menuLENPServer = New System.Windows.Forms.MenuItem
        Me.menuLENPClient = New System.Windows.Forms.MenuItem
        Me.contextMain = New System.Windows.Forms.ContextMenu
        Me.contextShow = New System.Windows.Forms.MenuItem
        Me.contextExit = New System.Windows.Forms.MenuItem
        Me.trayIcon = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.labelStatus = New System.Windows.Forms.Label
        Me.txtChatLog = New System.Windows.Forms.TextBox
        Me.txtChat = New System.Windows.Forms.TextBox
        Me.buttonChat = New System.Windows.Forms.Button
        Me.buttonGo = New System.Windows.Forms.Button
        Me.labelChannel = New System.Windows.Forms.Label
        Me.comboRealm = New System.Windows.Forms.ComboBox
        Me.labelBotServer = New System.Windows.Forms.Label
        Me.labelFirstChannel = New System.Windows.Forms.Label
        Me.labelPassword = New System.Windows.Forms.Label
        Me.labelUserName = New System.Windows.Forms.Label
        Me.labelTFTKey = New System.Windows.Forms.Label
        Me.labelROCKey = New System.Windows.Forms.Label
        Me.labelRealm = New System.Windows.Forms.Label
        Me.txtChannel = New System.Windows.Forms.TextBox
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.txtAccount = New System.Windows.Forms.TextBox
        Me.txtTFT = New System.Windows.Forms.TextBox
        Me.txtROC = New System.Windows.Forms.TextBox
        Me.groupParam = New System.Windows.Forms.GroupBox
        Me.txtHostPort = New System.Windows.Forms.TextBox
        Me.labelHostPort = New System.Windows.Forms.Label
        Me.txtWar3 = New System.Windows.Forms.TextBox
        Me.labelSummary = New System.Windows.Forms.Label
        Me.listChannelPeople = New System.Windows.Forms.ListBox
        Me.groupUser = New System.Windows.Forms.GroupBox
        Me.buttonUserRemove = New System.Windows.Forms.Button
        Me.buttonUserAdd = New System.Windows.Forms.Button
        Me.txtUserName = New System.Windows.Forms.TextBox
        Me.listUser = New System.Windows.Forms.ListBox
        Me.groupGame = New System.Windows.Forms.GroupBox
        Me.buttonGameInfo = New System.Windows.Forms.Button
        Me.buttonGameStop = New System.Windows.Forms.Button
        Me.listGame = New System.Windows.Forms.ListBox
        Me.groupParam.SuspendLayout()
        Me.groupUser.SuspendLayout()
        Me.groupGame.SuspendLayout()
        Me.SuspendLayout()
        '
        'menuMain
        '
        Me.menuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.menuFile, Me.menuLENP})
        '
        'menuFile
        '
        Me.menuFile.Index = 0
        Me.menuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.menuAutoReconnect, Me.menuConvertor, Me.menuExit})
        Me.menuFile.Text = "File"
        '
        'menuAutoReconnect
        '
        Me.menuAutoReconnect.Index = 0
        Me.menuAutoReconnect.Text = "Auto Reconnect"
        '
        'menuConvertor
        '
        Me.menuConvertor.Index = 1
        Me.menuConvertor.Text = "Hexadecimal Convertor"
        '
        'menuExit
        '
        Me.menuExit.Index = 2
        Me.menuExit.Text = "Exit"
        '
        'menuLENP
        '
        Me.menuLENP.Index = 1
        Me.menuLENP.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.menuLENPServer, Me.menuLENPClient})
        Me.menuLENP.Text = "LENP"
        '
        'menuLENPServer
        '
        Me.menuLENPServer.Index = 0
        Me.menuLENPServer.Text = "LENP Server"
        '
        'menuLENPClient
        '
        Me.menuLENPClient.Index = 1
        Me.menuLENPClient.Text = "LENP Client"
        '
        'contextMain
        '
        Me.contextMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.contextShow, Me.contextExit})
        '
        'contextShow
        '
        Me.contextShow.Index = 0
        Me.contextShow.Text = "Show Game Host"
        '
        'contextExit
        '
        Me.contextExit.Index = 1
        Me.contextExit.Text = "Exit"
        '
        'trayIcon
        '
        Me.trayIcon.Icon = CType(resources.GetObject("trayIcon.Icon"), System.Drawing.Icon)
        Me.trayIcon.Text = "Lain Host"
        Me.trayIcon.Visible = True
        '
        'labelStatus
        '
        Me.labelStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.labelStatus.BackColor = System.Drawing.Color.DarkGray
        Me.labelStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.labelStatus.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelStatus.Location = New System.Drawing.Point(0, 587)
        Me.labelStatus.Name = "labelStatus"
        Me.labelStatus.Size = New System.Drawing.Size(564, 21)
        Me.labelStatus.TabIndex = 7
        '
        'txtChatLog
        '
        Me.txtChatLog.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtChatLog.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChatLog.Location = New System.Drawing.Point(7, 284)
        Me.txtChatLog.Multiline = True
        Me.txtChatLog.Name = "txtChatLog"
        Me.txtChatLog.ReadOnly = True
        Me.txtChatLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtChatLog.Size = New System.Drawing.Size(547, 255)
        Me.txtChatLog.TabIndex = 1
        Me.txtChatLog.Text = "======================" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Chat Log" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "======================"
        '
        'txtChat
        '
        Me.txtChat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtChat.Location = New System.Drawing.Point(7, 546)
        Me.txtChat.Multiline = True
        Me.txtChat.Name = "txtChat"
        Me.txtChat.Size = New System.Drawing.Size(494, 34)
        Me.txtChat.TabIndex = 3
        '
        'buttonChat
        '
        Me.buttonChat.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.buttonChat.Location = New System.Drawing.Point(508, 546)
        Me.buttonChat.Name = "buttonChat"
        Me.buttonChat.Size = New System.Drawing.Size(46, 34)
        Me.buttonChat.TabIndex = 4
        Me.buttonChat.Text = "Send"
        '
        'buttonGo
        '
        Me.buttonGo.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.buttonGo.Location = New System.Drawing.Point(7, 7)
        Me.buttonGo.Name = "buttonGo"
        Me.buttonGo.Size = New System.Drawing.Size(40, 243)
        Me.buttonGo.TabIndex = 0
        Me.buttonGo.Text = ":)"
        '
        'labelChannel
        '
        Me.labelChannel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.labelChannel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.labelChannel.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelChannel.Location = New System.Drawing.Point(7, 257)
        Me.labelChannel.Name = "labelChannel"
        Me.labelChannel.Size = New System.Drawing.Size(714, 19)
        Me.labelChannel.TabIndex = 5
        Me.labelChannel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'comboRealm
        '
        Me.comboRealm.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.comboRealm.ItemHeight = 13
        Me.comboRealm.Items.AddRange(New Object() {"uswest.battle.net", "useast.battle.net", "asia.battle.net", "europe.battle.net"})
        Me.comboRealm.Location = New System.Drawing.Point(107, 49)
        Me.comboRealm.Name = "comboRealm"
        Me.comboRealm.Size = New System.Drawing.Size(341, 21)
        Me.comboRealm.TabIndex = 1
        Me.comboRealm.Text = "uswest.battle.net"
        '
        'labelBotServer
        '
        Me.labelBotServer.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelBotServer.Location = New System.Drawing.Point(7, 21)
        Me.labelBotServer.Name = "labelBotServer"
        Me.labelBotServer.Size = New System.Drawing.Size(93, 21)
        Me.labelBotServer.TabIndex = 31
        Me.labelBotServer.Text = "War3 Path"
        '
        'labelFirstChannel
        '
        Me.labelFirstChannel.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelFirstChannel.Location = New System.Drawing.Point(7, 187)
        Me.labelFirstChannel.Name = "labelFirstChannel"
        Me.labelFirstChannel.Size = New System.Drawing.Size(93, 21)
        Me.labelFirstChannel.TabIndex = 28
        Me.labelFirstChannel.Text = "Home Channel"
        '
        'labelPassword
        '
        Me.labelPassword.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelPassword.Location = New System.Drawing.Point(7, 159)
        Me.labelPassword.Name = "labelPassword"
        Me.labelPassword.Size = New System.Drawing.Size(93, 21)
        Me.labelPassword.TabIndex = 27
        Me.labelPassword.Text = "Password"
        '
        'labelUserName
        '
        Me.labelUserName.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelUserName.Location = New System.Drawing.Point(7, 132)
        Me.labelUserName.Name = "labelUserName"
        Me.labelUserName.Size = New System.Drawing.Size(93, 21)
        Me.labelUserName.TabIndex = 26
        Me.labelUserName.Text = "User Name"
        '
        'labelTFTKey
        '
        Me.labelTFTKey.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelTFTKey.Location = New System.Drawing.Point(7, 104)
        Me.labelTFTKey.Name = "labelTFTKey"
        Me.labelTFTKey.Size = New System.Drawing.Size(93, 14)
        Me.labelTFTKey.TabIndex = 25
        Me.labelTFTKey.Text = "TFT Key"
        '
        'labelROCKey
        '
        Me.labelROCKey.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelROCKey.Location = New System.Drawing.Point(7, 76)
        Me.labelROCKey.Name = "labelROCKey"
        Me.labelROCKey.Size = New System.Drawing.Size(93, 21)
        Me.labelROCKey.TabIndex = 24
        Me.labelROCKey.Text = "ROC Key"
        '
        'labelRealm
        '
        Me.labelRealm.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelRealm.Location = New System.Drawing.Point(7, 49)
        Me.labelRealm.Name = "labelRealm"
        Me.labelRealm.Size = New System.Drawing.Size(93, 20)
        Me.labelRealm.TabIndex = 23
        Me.labelRealm.Text = "Realm"
        '
        'txtChannel
        '
        Me.txtChannel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtChannel.Location = New System.Drawing.Point(107, 187)
        Me.txtChannel.Name = "txtChannel"
        Me.txtChannel.Size = New System.Drawing.Size(341, 20)
        Me.txtChannel.TabIndex = 6
        Me.txtChannel.Text = "The Void"
        '
        'txtPassword
        '
        Me.txtPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPassword.Location = New System.Drawing.Point(107, 159)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(341, 20)
        Me.txtPassword.TabIndex = 5
        Me.txtPassword.Text = "blacksheepwall"
        '
        'txtAccount
        '
        Me.txtAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAccount.Location = New System.Drawing.Point(107, 132)
        Me.txtAccount.Name = "txtAccount"
        Me.txtAccount.Size = New System.Drawing.Size(341, 20)
        Me.txtAccount.TabIndex = 4
        Me.txtAccount.Text = "lain"
        '
        'txtTFT
        '
        Me.txtTFT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTFT.Location = New System.Drawing.Point(107, 104)
        Me.txtTFT.MaxLength = 26
        Me.txtTFT.Name = "txtTFT"
        Me.txtTFT.Size = New System.Drawing.Size(341, 20)
        Me.txtTFT.TabIndex = 3
        Me.txtTFT.Text = "XXXXXXXXXXXXXXXXXXXXXXXXXX"
        '
        'txtROC
        '
        Me.txtROC.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtROC.Location = New System.Drawing.Point(107, 76)
        Me.txtROC.MaxLength = 26
        Me.txtROC.Name = "txtROC"
        Me.txtROC.Size = New System.Drawing.Size(341, 20)
        Me.txtROC.TabIndex = 2
        Me.txtROC.Text = "XXXXXXXXXXXXXXXXXXXXXXXXXX"
        '
        'groupParam
        '
        Me.groupParam.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.groupParam.Controls.Add(Me.txtHostPort)
        Me.groupParam.Controls.Add(Me.labelHostPort)
        Me.groupParam.Controls.Add(Me.txtWar3)
        Me.groupParam.Controls.Add(Me.comboRealm)
        Me.groupParam.Controls.Add(Me.labelBotServer)
        Me.groupParam.Controls.Add(Me.labelFirstChannel)
        Me.groupParam.Controls.Add(Me.labelPassword)
        Me.groupParam.Controls.Add(Me.labelUserName)
        Me.groupParam.Controls.Add(Me.labelTFTKey)
        Me.groupParam.Controls.Add(Me.labelROCKey)
        Me.groupParam.Controls.Add(Me.labelRealm)
        Me.groupParam.Controls.Add(Me.txtChannel)
        Me.groupParam.Controls.Add(Me.txtPassword)
        Me.groupParam.Controls.Add(Me.txtAccount)
        Me.groupParam.Controls.Add(Me.txtTFT)
        Me.groupParam.Controls.Add(Me.txtROC)
        Me.groupParam.Location = New System.Drawing.Point(267, 7)
        Me.groupParam.Name = "groupParam"
        Me.groupParam.Size = New System.Drawing.Size(454, 243)
        Me.groupParam.TabIndex = 0
        Me.groupParam.TabStop = False
        Me.groupParam.Text = "Wacraft Fronzen Throne Configuration"
        '
        'txtHostPort
        '
        Me.txtHostPort.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtHostPort.Location = New System.Drawing.Point(107, 215)
        Me.txtHostPort.Name = "txtHostPort"
        Me.txtHostPort.Size = New System.Drawing.Size(341, 20)
        Me.txtHostPort.TabIndex = 7
        Me.txtHostPort.Text = "6000"
        '
        'labelHostPort
        '
        Me.labelHostPort.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelHostPort.Location = New System.Drawing.Point(7, 215)
        Me.labelHostPort.Name = "labelHostPort"
        Me.labelHostPort.Size = New System.Drawing.Size(93, 21)
        Me.labelHostPort.TabIndex = 34
        Me.labelHostPort.Text = "Host Port"
        '
        'txtWar3
        '
        Me.txtWar3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtWar3.Location = New System.Drawing.Point(107, 21)
        Me.txtWar3.Name = "txtWar3"
        Me.txtWar3.Size = New System.Drawing.Size(341, 20)
        Me.txtWar3.TabIndex = 0
        Me.txtWar3.Text = "C:\Program Files\Warcraft III\"
        '
        'labelSummary
        '
        Me.labelSummary.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.labelSummary.BackColor = System.Drawing.Color.DarkGray
        Me.labelSummary.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelSummary.Location = New System.Drawing.Point(561, 546)
        Me.labelSummary.Name = "labelSummary"
        Me.labelSummary.Size = New System.Drawing.Size(167, 62)
        Me.labelSummary.TabIndex = 9
        Me.labelSummary.Text = "-=[Lain]=-"
        '
        'listChannelPeople
        '
        Me.listChannelPeople.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.listChannelPeople.Location = New System.Drawing.Point(561, 284)
        Me.listChannelPeople.Name = "listChannelPeople"
        Me.listChannelPeople.Size = New System.Drawing.Size(160, 225)
        Me.listChannelPeople.TabIndex = 2
        '
        'groupUser
        '
        Me.groupUser.Controls.Add(Me.buttonUserRemove)
        Me.groupUser.Controls.Add(Me.buttonUserAdd)
        Me.groupUser.Controls.Add(Me.txtUserName)
        Me.groupUser.Controls.Add(Me.listUser)
        Me.groupUser.Location = New System.Drawing.Point(53, 7)
        Me.groupUser.Name = "groupUser"
        Me.groupUser.Size = New System.Drawing.Size(94, 243)
        Me.groupUser.TabIndex = 10
        Me.groupUser.TabStop = False
        Me.groupUser.Text = "Admins"
        '
        'buttonUserRemove
        '
        Me.buttonUserRemove.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.buttonUserRemove.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.buttonUserRemove.Location = New System.Drawing.Point(7, 215)
        Me.buttonUserRemove.Name = "buttonUserRemove"
        Me.buttonUserRemove.Size = New System.Drawing.Size(80, 20)
        Me.buttonUserRemove.TabIndex = 3
        Me.buttonUserRemove.Text = "Remove"
        '
        'buttonUserAdd
        '
        Me.buttonUserAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.buttonUserAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.buttonUserAdd.Location = New System.Drawing.Point(7, 187)
        Me.buttonUserAdd.Name = "buttonUserAdd"
        Me.buttonUserAdd.Size = New System.Drawing.Size(80, 20)
        Me.buttonUserAdd.TabIndex = 2
        Me.buttonUserAdd.Text = "Add"
        '
        'txtUserName
        '
        Me.txtUserName.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUserName.Location = New System.Drawing.Point(7, 159)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(80, 20)
        Me.txtUserName.TabIndex = 1
        '
        'listUser
        '
        Me.listUser.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.listUser.Location = New System.Drawing.Point(7, 21)
        Me.listUser.Name = "listUser"
        Me.listUser.Size = New System.Drawing.Size(80, 121)
        Me.listUser.TabIndex = 0
        '
        'groupGame
        '
        Me.groupGame.Controls.Add(Me.buttonGameInfo)
        Me.groupGame.Controls.Add(Me.buttonGameStop)
        Me.groupGame.Controls.Add(Me.listGame)
        Me.groupGame.Location = New System.Drawing.Point(153, 7)
        Me.groupGame.Name = "groupGame"
        Me.groupGame.Size = New System.Drawing.Size(107, 243)
        Me.groupGame.TabIndex = 11
        Me.groupGame.TabStop = False
        Me.groupGame.Text = "Games"
        '
        'buttonGameInfo
        '
        Me.buttonGameInfo.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.buttonGameInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.buttonGameInfo.Location = New System.Drawing.Point(7, 187)
        Me.buttonGameInfo.Name = "buttonGameInfo"
        Me.buttonGameInfo.Size = New System.Drawing.Size(93, 20)
        Me.buttonGameInfo.TabIndex = 4
        Me.buttonGameInfo.Text = "Info"
        '
        'buttonGameStop
        '
        Me.buttonGameStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.buttonGameStop.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.buttonGameStop.Location = New System.Drawing.Point(7, 215)
        Me.buttonGameStop.Name = "buttonGameStop"
        Me.buttonGameStop.Size = New System.Drawing.Size(93, 20)
        Me.buttonGameStop.TabIndex = 3
        Me.buttonGameStop.Text = "Stop"
        '
        'listGame
        '
        Me.listGame.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.listGame.Location = New System.Drawing.Point(7, 21)
        Me.listGame.Name = "listGame"
        Me.listGame.Size = New System.Drawing.Size(93, 147)
        Me.listGame.TabIndex = 0
        '
        'frmLainEthLite
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(728, 606)
        Me.Controls.Add(Me.groupGame)
        Me.Controls.Add(Me.groupUser)
        Me.Controls.Add(Me.listChannelPeople)
        Me.Controls.Add(Me.labelSummary)
        Me.Controls.Add(Me.labelChannel)
        Me.Controls.Add(Me.labelStatus)
        Me.Controls.Add(Me.buttonChat)
        Me.Controls.Add(Me.buttonGo)
        Me.Controls.Add(Me.groupParam)
        Me.Controls.Add(Me.txtChat)
        Me.Controls.Add(Me.txtChatLog)
        Me.Icon = CType(resources.GetObject("trayIcon.Icon"), System.Drawing.Icon)
        Me.Menu = Me.menuMain
        Me.Name = "frmLainEthLite"
        Me.groupParam.ResumeLayout(False)
        Me.groupParam.PerformLayout()
        Me.groupUser.ResumeLayout(False)
        Me.groupUser.PerformLayout()
        Me.groupGame.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "menu"
    Private Sub menuConvertor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuConvertor.Click
        Dim form As frmConvertor

        form = New frmConvertor
        form.Show()
    End Sub
    Private Sub menuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuExit.Click
        ExitProgram()
    End Sub
    Private Sub menuAutoReconnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuAutoReconnect.Click
        If menuAutoReconnect.Checked = False Then
            If bnet.IsSendChatAble Then
                menuAutoReconnect.Checked = True
                aliveTimer.Start()
            Else
                MessageBox.Show("Please Tick This Only After The Bot Is Logged On", ProjectLainVersion, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            menuAutoReconnect.Checked = False
            aliveTimer.Stop()
        End If
    End Sub
    Private Sub menuLENPServer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuLENPServer.Click
        formLENPServer.Show()
        formLENPServer.BringToFront()
    End Sub
    Private Sub menuLENPClient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuLENPClient.Click
        Dim form As frmLENPClient

        form = New frmLENPClient
        form.Show()
    End Sub
#End Region

#Region "xml management"
    Private Sub LoadAll()
        Dim ds As DataSet

        Try
            ds = New DataSet(ProjectLainName)
            If IO.File.Exists(String.Format("{0}\{1}.xml", Application.StartupPath, ProjectLainConfig)) Then
                ds.ReadXml(String.Format("{0}\{1}.xml", Application.StartupPath, ProjectLainConfig))
                If ds.Tables.Contains(ProjectLainConfig) Then
                    LoadConfigurationTable(ds.Tables(ProjectLainConfig))
                End If
            End If

            ds = New DataSet(ProjectLainName)
            If IO.File.Exists(String.Format("{0}\{1}.xml", Application.StartupPath, ProjectLainMap)) = False Then
                ds.Tables.Add(CreateMapTable)
                ds.WriteXml(String.Format("{0}\{1}.xml", Application.StartupPath, ProjectLainMap))
            End If
            ds.ReadXml(String.Format("{0}\{1}.xml", Application.StartupPath, ProjectLainMap))
            If ds.Tables.Contains(ProjectLainMap) Then
                LoadMapTable(ds.Tables(ProjectLainMap))
            End If

            ds = New DataSet(ProjectLainName)
            If IO.File.Exists(String.Format("{0}\{1}.xml", Application.StartupPath, ProjectLainUser)) Then
                ds.ReadXml(String.Format("{0}\{1}.xml", Application.StartupPath, ProjectLainUser))
                If ds.Tables.Contains(ProjectLainUser) Then
                    LoadUserTable(ds.Tables(ProjectLainUser))
                End If
            Else
                listUser.Items.Add("leax")
            End If

            ds = New DataSet(ProjectLainName)
            If IO.File.Exists(String.Format("{0}\{1}.xml", Application.StartupPath, ProjectLainCustomData)) Then
                ds.ReadXml(String.Format("{0}\{1}.xml", Application.StartupPath, ProjectLainCustomData))
                If ds.Tables.Contains(ProjectLainCustomData) Then
                    LoadCustomDataTable(ds.Tables(ProjectLainCustomData))
                End If
            End If

        Catch ex As Exception
        End Try
    End Sub

    Private Function LoadCustomDataTable(ByVal table As DataTable) As Boolean
        Dim row As DataRow
        Dim exeversion As String()
        Dim exeversionhash As String()
        Dim passwordhashtype As String
        Try
            customData = New clsBNETCustomData
            If table.Rows.Count > 0 Then
                row = table.Rows(0)

                If table.Columns.Contains("exeversion") Then
                    exeversion = CStr(row("exeversion")).Split(Convert.ToChar(" "))
                    If exeversion.Length = 4 Then
                        customData.SetExeVersion(New Byte() {CByte(exeversion(0)), CByte(exeversion(1)), CByte(exeversion(2)), CByte(exeversion(3))})
                    End If
                End If

                If table.Columns.Contains("exeversionhash") Then
                    exeversionhash = CStr(row("exeversionhash")).Split(Convert.ToChar(" "))
                    If exeversionhash.Length = 4 Then
                        customData.SetExeVersionHash(New Byte() {CByte(exeversionhash(0)), CByte(exeversionhash(1)), CByte(exeversionhash(2)), CByte(exeversionhash(3))})
                    End If
                End If

                If table.Columns.Contains("passwordhashtype") Then
                    passwordhashtype = CStr(row("passwordhashtype"))
                    If passwordhashtype IsNot Nothing Then
                        customData.SetPasswordHashType(passwordhashtype)
                    End If
                End If

            End If
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function LoadUserTable(ByVal table As DataTable) As Boolean
        Dim row As DataRow
        Try
            If table.Rows.Count > 0 AndAlso table.Columns.Contains("user") Then
                listUser.BeginUpdate()
                listUser.Items.Clear()
                For Each row In table.Rows
                    listUser.Items.Add(CStr(row("user")))
                Next
                listUser.EndUpdate()
                Return True
            End If

            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Function CreateMapTable() As DataTable
        Dim table As DataTable
        Dim column As DataColumn
        Dim row As DataRow
        Try
            table = New DataTable(ProjectLainMap)

            column = New DataColumn("mappath", GetType(String))
            table.Columns.Add(column)
            column = New DataColumn("mapsize", GetType(String))
            table.Columns.Add(column)
            column = New DataColumn("mapinfo", GetType(String))
            table.Columns.Add(column)
            column = New DataColumn("mapcrc", GetType(String))
            table.Columns.Add(column)

            row = table.NewRow

            row("mappath") = "Maps\Download\DotA Allstars v6.49b.w3x"
            row("mapsize") = "34 28 38 0"
            row("mapinfo") = "46 233 147 143"
            row("mapcrc") = "173 166 209 216"

            table.Rows.Add(row)
            Return table
        Catch ex As Exception
            Return New DataTable
        End Try
    End Function
    Private Function LoadMapTable(ByVal table As DataTable) As Boolean
        Dim row As DataRow
        Dim mappath As String
        Dim mapsize As String()
        Dim mapinfo As String()
        Dim mapcrc As String()
        Try
            If table.Rows.Count > 0 Then
                row = table.Rows(0)
                If table.Columns.Contains("mappath") AndAlso table.Columns.Contains("mapsize") AndAlso table.Columns.Contains("mapinfo") AndAlso table.Columns.Contains("mapcrc") Then
                    mappath = CStr(row("mappath"))
                    mapsize = CStr(row("mapsize")).Split(Convert.ToChar(" "))
                    mapinfo = CStr(row("mapinfo")).Split(Convert.ToChar(" "))
                    mapcrc = CStr(row("mapcrc")).Split(Convert.ToChar(" "))

                    If mappath.Length > 0 AndAlso mapsize.Length = 4 AndAlso mapinfo.Length = 4 AndAlso mapcrc.Length = 4 Then
                        map = New clsGameHostMap(mappath, _
                                                New Byte() {CByte(mapsize(0)), CByte(mapsize(1)), CByte(mapsize(2)), CByte(mapsize(3))}, _
                                                New Byte() {CByte(mapinfo(0)), CByte(mapinfo(1)), CByte(mapinfo(2)), CByte(mapinfo(3))}, _
                                                New Byte() {CByte(mapcrc(0)), CByte(mapcrc(1)), CByte(mapcrc(2)), CByte(mapcrc(3))})
                        Return True
                    End If
                End If
            End If
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Function LoadConfigurationTable(ByVal table As DataTable) As Boolean
        Dim row As DataRow
        Try
            If table.Rows.Count > 0 Then
                row = table.Rows(0)
                If table.Columns.Contains("war3path") Then
                    txtWar3.Text = CStr(row("war3path"))
                End If
                If table.Columns.Contains("realm") Then
                    comboRealm.Text = CStr(row("realm"))
                End If
                If table.Columns.Contains("ROCkey") Then
                    txtROC.Text = CStr(row("ROCkey"))
                End If
                If table.Columns.Contains("TFTkey") Then
                    txtTFT.Text = CStr(row("TFTkey"))
                End If
                If table.Columns.Contains("account") Then
                    txtAccount.Text = CStr(row("account"))
                End If
                If table.Columns.Contains("password") Then
                    txtPassword.Text = CStr(row("password"))
                End If
                If table.Columns.Contains("channel") Then
                    txtChannel.Text = CStr(row("channel"))
                End If
                If table.Columns.Contains("hostport") Then
                    txtHostPort.Text = CStr(row("hostport"))
                End If
                Return True
            End If
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Sub SaveAll()
        Dim ds As DataSet
        Try
            ds = New DataSet(ProjectLainName)
            ds.Tables.Add(SaveConfigurationTable)
            ds.WriteXml(String.Format("{0}\{1}.xml", Application.StartupPath, ProjectLainConfig))

            ds = New DataSet(ProjectLainName)
            ds.Tables.Add(SaveUserTable)
            ds.WriteXml(String.Format("{0}\{1}.xml", Application.StartupPath, ProjectLainUser))
        Catch ex As Exception
        End Try
    End Sub

    Private Function SaveUserTable() As DataTable
        Dim table As DataTable
        Dim column As DataColumn
        Dim row As DataRow
        Dim name As String
        Try
            table = New DataTable(ProjectLainUser)

            column = New DataColumn("user", GetType(String))
            table.Columns.Add(column)

            For Each name In listUser.Items
                row = table.NewRow
                row("user") = name
                table.Rows.Add(row)
            Next

            Return table
        Catch ex As Exception
            Return New DataTable
        End Try
    End Function
    Private Function SaveConfigurationTable() As DataTable

        Dim table As DataTable
        Dim column As DataColumn
        Dim row As DataRow
        Try
            table = New DataTable(ProjectLainConfig)

            column = New DataColumn("war3path", GetType(String))
            table.Columns.Add(column)
            column = New DataColumn("realm", GetType(String))
            table.Columns.Add(column)
            column = New DataColumn("ROCkey", GetType(String))
            table.Columns.Add(column)
            column = New DataColumn("TFTkey", GetType(String))
            table.Columns.Add(column)
            column = New DataColumn("account", GetType(String))
            table.Columns.Add(column)
            column = New DataColumn("password", GetType(String))
            table.Columns.Add(column)
            column = New DataColumn("channel", GetType(String))
            table.Columns.Add(column)
            column = New DataColumn("hostport", GetType(String))
            table.Columns.Add(column)

            row = table.NewRow
            row("war3path") = txtWar3.Text
            row("realm") = comboRealm.Text
            row("ROCkey") = txtROC.Text
            row("TFTkey") = txtTFT.Text
            row("account") = txtAccount.Text
            row("password") = txtPassword.Text
            row("channel") = txtChannel.Text
            row("hostport") = txtHostPort.Text

            table.Rows.Add(row)

            Return table
        Catch ex As Exception
            Return New DataTable
        End Try
    End Function
#End Region

#Region "bnet event"

    Private Sub bnet_EventBnetSIDPING() Handles bnet.EventBnetSIDPING
        labelStatus.Invoke(New clsHelper.DelegateControlText(AddressOf clsHelper.ControlText), New Object() {labelStatus, String.Format("Memory Usage : {0} MB, Threads : {1}, Handles : {2}", Math.Round(Diagnostics.Process.GetCurrentProcess.PrivateMemorySize64 / 1024 / 1024, 2), Diagnostics.Process.GetCurrentProcess.Threads.Count, Diagnostics.Process.GetCurrentProcess.HandleCount)})
    End Sub

    Private Sub bnet_EventEngingState() Handles bnet.EventEngingState
        ReflectEngineState()
    End Sub
    Private Sub bnet_EventLogOnStatus(ByVal failpoint As clsBNET.LogOnFailPoints, ByVal info As String) Handles bnet.EventLogOnStatus
        labelSummary.Invoke(New clsHelper.DelegateControlText(AddressOf clsHelper.ControlText), New Object() {labelSummary, info})
    End Sub
    Private Sub bnet_OnEventMessage(ByVal server As clsCommandPacket.PacketType, ByVal msg As String) Handles bnet.EventMessage
        labelStatus.Invoke(New clsHelper.DelegateControlText(AddressOf clsHelper.ControlText), New Object() {labelStatus, String.Format("{0} -> {1}", server, msg)})
    End Sub
    Private Sub bnet_OnEventError(ByVal errorFunction As String, ByVal errorString As String) Handles bnet.EventError
        labelStatus.Invoke(New clsHelper.DelegateControlText(AddressOf clsHelper.ControlText), New Object() {labelStatus, String.Format("{0} -> {1}", "Error", errorFunction)})
        txtChatLog.Invoke(New clsHelper.DelegateTextBoxAppend(AddressOf clsHelper.ControlTextBoxAppend), New Object() {txtChatLog, String.Format("ERROR[{0}]:{1}{2}", errorFunction, errorString, Environment.NewLine)})
    End Sub
    Private Sub bnetchannel_OnEventChannelChange() Handles channel.EventChannelChange
        If channel.GetChannelName.Length = 0 Then
            bnet.GoChannel("The Void")
        Else
            labelChannel.Invoke(New clsHelper.DelegateControlText(AddressOf clsHelper.ControlText), New Object() {labelChannel, channel.GetChannelName})
            listChannelPeople.Invoke(New clsHelper.DelegateControlListBoxDataSource(AddressOf clsHelper.ControlListBoxDataSource), New Object() {listChannelPeople, channel.GetChannelPeopleList})
        End If
    End Sub
    Private Sub bnet_EventIncomingChat(ByVal eventChat As clsIncomingChatChannel) Handles bnet.EventIncomingChat
        channel.Process(eventChat)
        bot.ProcessCommand(eventChat)
    End Sub
    Private Sub bnetchannel_OnEventChannelChatMessage(ByVal msg As String) Handles channel.EventChannelChatMessage
        txtChatLog.Invoke(New clsHelper.DelegateTextBoxAppend(AddressOf clsHelper.ControlTextBoxAppend), New Object() {txtChatLog, String.Format("{0}{1}", msg, Environment.NewLine)})
    End Sub

    Private Sub bnet_EventBnetSIDSTARTADVEX3Result(ByVal isOK As Boolean) Handles bnet.EventBnetSIDSTARTADVEX3Result
        If isOK Then
            'out of chat
        Else
            currentHost.Dispose("Host Game Failed")
            SendChat("/me : Game Failed to be Hosted, try another name")
        End If
    End Sub

#End Region

#Region "host event"
    Private Sub host_EventHostUncreate()
        Debug.WriteLine("uncreated")
        bnet.GameUnCreate()
        Debug.WriteLine("after trysend")
        If bnet.IsSendChatAble = False Then
            bnet.EnterChat()
        End If
        Debug.WriteLine("after enter chat")
    End Sub
    Private Sub host_EventHostDisposed(ByVal host As clsGameHost, ByVal reason As String)
        Try
            If currentHost Is host Then
                currentHost = New clsGameHost
            End If

            If listHost.Contains(host) Then
                listHost.Remove(host)
                UpdateGameList()

                RemoveHandler host.EventHostUncreate, AddressOf host_EventHostUncreate
                RemoveHandler host.EventHostDisposed, AddressOf host_EventHostDisposed
                RemoveHandler host.EventGameWon, AddressOf host_EventGameWon
            End If

            SendChat(String.Format("[{0}] : {1}", host.GetGameName, reason), ProjectLainVersion, True)
        Catch ex As Exception
            Debug.WriteLine(ex)
        End Try

    End Sub
    Private Sub host_EventGameWon(ByVal callerName As String, ByVal gameName As String, ByVal sentinelPlayer() As String, ByVal scourgePlayer() As String, ByVal referee() As String, ByVal winner As String)
        'SendChat(String.Format("Game:[{0}] Result Win -> [{1}]", gameName, winner))
        SendChat(String.Format("{0} wins the game {1}", winner, gameName))
    End Sub
#End Region

#Region "bot channel"
    Private Sub bot_EventBotMap(ByVal isWhisper As Boolean, ByVal owner As String, ByVal mapName As String) Handles bot.EventBotMap
        Dim ds As DataSet
        Dim msg As String

        Try
            ds = New DataSet(ProjectLainName)

            mapName = mapName.ToLower
            If mapName.EndsWith(".xml") Then
                mapName = mapName.Substring(0, mapName.Length - 4)
            End If

            If IO.File.Exists(String.Format("{0}\{1}.xml", Application.StartupPath, mapName)) Then
                ds.ReadXml(String.Format("{0}\{1}.xml", Application.StartupPath, mapName))
                If ds.Tables.Contains(ProjectLainMap) Then
                    If LoadMapTable(ds.Tables(ProjectLainMap)) Then
                        msg = String.Format("Map file [{0}] Loaded Succesfully", mapName)
                    Else
                        msg = String.Format("Map file [{0}] Fail to Load", mapName)
                    End If
                Else
                    msg = String.Format("Map file [{0}] not Valid", mapName)
                End If
            Else
                msg = String.Format("Map file [{0}] not Found", mapName)
            End If

            If isWhisper Then
                SendChat(String.Format("/w {0} {1}", owner, msg))
            Else
                SendChat(String.Format("/me : {0}", msg))
            End If

        Catch ex As Exception
            Debug.WriteLine(ex)
        End Try


    End Sub
    Private Sub bot_EventBotSay(ByVal msg As String) Handles bot.EventBotSay
        SendChat(msg)
    End Sub
    Private Sub bot_EventBotResponse(ByVal msg As String, ByVal isWhisper As Boolean, ByVal owner As String) Handles bot.EventBotResponse
        If isWhisper Then
            SendChat(String.Format("/w {0} {1}", owner, msg))
        Else
            SendChat(String.Format("/me : {0}", msg))
        End If
    End Sub
    Private Sub bot_EventBotHost(ByVal isPublic As Boolean, ByVal numPlayers As Integer, ByVal gameName As String, ByVal callerName As String, ByVal isWhisper As Boolean, ByVal owner As String) Handles bot.EventBotHost
        Dim host As clsGameHost
        Dim state As Byte

        If isPublic Then
            state = GAME_PUBLIC
        Else
            state = GAME_PRIVATE
        End If

        host = New clsGameHost(GetAdminList, state, numPlayers, gameName, bnet.GetUniqueUserName, callerName, bnet.GetHostPort, map.GetMapPath, map.GetMapSize, map.GetMapInfo, map.GetMapCRC, bnet)
        If host.HostStart Then
            listHost.Add(host)
            UpdateGameList()

            currentHost = host
            AddHandler host.EventHostUncreate, AddressOf host_EventHostUncreate
            AddHandler host.EventHostDisposed, AddressOf host_EventHostDisposed
            AddHandler host.EventGameWon, AddressOf host_EventGameWon

            If isWhisper Then
                SendChat(String.Format("/w {0} Creating Game: [{0}] started by [{1}]...", owner, host.GetGameName, host.GetCallerName))
            Else
                SendChat(String.Format("/me : Creating Game: [{0}] started by [{1}]...", host.GetGameName, host.GetCallerName))
            End If

            'Threading.Thread.Sleep(1000)
            bnet.GetFriendList()    'MrJag|0.8c|reserve|get the current list of friends
            bnet.GetClanList()      'MrJag|0.8c|reserve|get the current list of clan members
            bnet.GameCreate(labelChannel.Text, state, numPlayers, gameName, bnet.GetUniqueUserName, 0, map.GetMapPath, map.GetMapCRC)
        Else
            If currentHost.GetGameName.Length > 0 Then
                If isWhisper Then
                    SendChat(String.Format("/w {0} Game: [{1}] is waiting", owner, currentHost.GetGameName))
                Else
                    SendChat(String.Format("/me : Game: [{0}] is waiting", currentHost.GetGameName))
                End If
            Else
                If isWhisper Then
                    SendChat(String.Format("/w {0} Hosting Failed, Possible Port In Use", owner))
                Else
                    SendChat(String.Format("/me : Hosting Failed, Possible Port In Use"))
                End If
            End If
        End If
    End Sub
    Private Sub bot_EventBotUnHost(ByVal isWhisper As Boolean, ByVal owner As String) Handles bot.EventBotUnHost
        Dim gamename As String
        Dim result As String

        If currentHost.GetGameName.Length > 0 Then
            gamename = currentHost.GetGameName

            If currentHost.GetIsCountDownStarted = False Then
                currentHost.Dispose("Game Unhosted Manually")
                result = String.Format("Game [{0}] is stopped", gamename)
            Else
                result = String.Format("Game [{0}] already started", gamename)
            End If
        Else
            result = String.Format("No valid targeted game")
        End If

        If isWhisper Then
            SendChat(String.Format("/w {0} {1}", owner, result))
        Else
            SendChat(String.Format("/me : {0}", result))
        End If
    End Sub
    Private Sub bot_EventBotGetGames(ByVal isWhisper As Boolean, ByVal owner As String) Handles bot.EventBotGetGames
        Dim host As clsGameHost
        Dim result As System.Text.StringBuilder

        result = New System.Text.StringBuilder
        SyncLock listHost.SyncRoot
            For Each host In listHost
                If host.GetIsCountDownStarted = False Then
                    result.Append(String.Format("[{0} : {1} : {2}], ", host.GetGameName, host.GetCallerName, host.GetTotalPlayers))
                Else
                    result.Append(String.Format("[{0}], ", host.GetGameName))
                End If
            Next
        End SyncLock

        If result.Length = 0 Then
            result.Append("No games in progress")
        End If

        If isWhisper Then
            SendChat(String.Format("/w {0} {1}", owner, result))
        Else
            SendChat(String.Format("/me : {0}", result))
        End If

    End Sub

#End Region

    Private Sub frmLainEthLite_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        listHost = ArrayList.Synchronized(New ArrayList)
        currentHost = New clsGameHost
        map = New clsGameHostMap
        customData = New clsBNETCustomData
        bnet = New clsBNET
        bot = New clsBotCommandHostChannel(New String() {})
        channel = New clsBNETChannel

        Me.Text = ProjectLainVersion
        txtChatLog.AppendText(Environment.NewLine)
        trayIcon.ContextMenu = contextMain
        trayIcon.Text = ProjectLainVersion
        trayIcon.Visible = True
        Me.ReflectEngineState()
        Me.buttonGo.Focus()

        aliveTimer = New Timers.Timer
        aliveTimer.Interval = 1000
        aliveCounter = 0
        channelJoinCounter = 0

        LoadAll()

        formLENPServer = New frmLENPServer
        lenp = New clsLENPServer(listHost)

    End Sub
    Private Sub frmLainEthLite_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        e.Cancel = True
        Me.Visible = False
    End Sub

    Private Sub txtChat_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtChat.TextChanged
        If txtChat.Text.Length > 0 Then
            txtChat.Text = TrimNewLine(txtChat.Text)
        End If
    End Sub
    Private Sub txtChat_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtChat.KeyPress
        If txtChat.Text.Length > 0 AndAlso e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            buttonChat_Click("txtChat_KeyPress", New EventArgs)
        End If
    End Sub
    Private Sub txtChatLog_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtChatLog.TextChanged
        If txtChatLog.Text.Length > txtChatLog.MaxLength * 0.9 Then
            txtChatLog.Text = txtChatLog.Text.Substring(CInt(txtChatLog.Text.Length * 0.7))
        End If
    End Sub
    Private Sub txtHostPort_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHostPort.LostFocus
        If IsNumeric(txtHostPort.Text) AndAlso CType(txtHostPort.Text, Integer) > 0 AndAlso CType(txtHostPort.Text, Integer) < 32000 Then
            Return
        Else
            txtHostPort.Text = "6000"
            MessageBox.Show("Invalid Host Port, Reverting to Default = 6000", ProjectLainVersion, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Sub txtWar3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWar3.LostFocus
        Dim file_war3_exe As String = "war3.exe"
        Dim file_storm_dll As String = "storm.dll"
        Dim file_game_dll As String = "game.dll"
        Dim war3path As String

        If txtWar3.Text.EndsWith("\") = False Then
            txtWar3.Text = txtWar3.Text & "\"
        End If

        war3path = txtWar3.Text
        If Directory.Exists(war3path) Then
            If File.Exists(war3path & file_war3_exe) AndAlso File.Exists(war3path & file_storm_dll) AndAlso File.Exists(war3path & file_game_dll) Then
                txtWar3.ForeColor = Color.Black
                Return
            End If
        End If
        txtWar3.ForeColor = Color.Red
    End Sub
    Private Sub buttonChat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonChat.Click
        SendChat(txtChat.Text)
        txtChat.Text = ""
        txtChat.Focus()
    End Sub
    Private Sub trayIcon_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles trayIcon.DoubleClick
        contextShow_Click("trayIcon_DoubleClick", New EventArgs)
    End Sub
    Private Sub contextShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles contextShow.Click
        Me.Visible = True
        Me.WindowState = FormWindowState.Normal
        Me.BringToFront()
    End Sub
    Private Sub contextExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles contextExit.Click
        ExitProgram()
    End Sub
    Private Sub buttonGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonGo.Click
        Start()
    End Sub

    Private Function Start() As Boolean
        Try
            If bnet.IsEngineRunning() Then
                bnet.Stop()
            Else
                listChannelPeople.Invoke(New clsHelper.DelegateControlListBoxDataSource(AddressOf clsHelper.ControlListBoxDataSource), New Object() {listChannelPeople, New Object() {}})
                labelSummary.Invoke(New clsHelper.DelegateControlText(AddressOf clsHelper.ControlText), New Object() {labelSummary, "Connecting to BattleNET..."})

                bot = New clsBotCommandHostChannel(GetAdminList())
                bnet.Start( _
                CType(comboRealm.Invoke(New clsHelper.DelegateControlTextGet(AddressOf clsHelper.ControlTextGet), New Object() {comboRealm}), String), _
                CType(txtWar3.Invoke(New clsHelper.DelegateControlTextGet(AddressOf clsHelper.ControlTextGet), New Object() {txtWar3}), String), _
                CType(txtROC.Invoke(New clsHelper.DelegateControlTextGet(AddressOf clsHelper.ControlTextGet), New Object() {txtROC}), String), _
                CType(txtTFT.Invoke(New clsHelper.DelegateControlTextGet(AddressOf clsHelper.ControlTextGet), New Object() {txtTFT}), String), _
                CType(txtAccount.Invoke(New clsHelper.DelegateControlTextGet(AddressOf clsHelper.ControlTextGet), New Object() {txtAccount}), String), _
                CType(txtPassword.Invoke(New clsHelper.DelegateControlTextGet(AddressOf clsHelper.ControlTextGet), New Object() {txtPassword}), String), _
                CType(txtChannel.Invoke(New clsHelper.DelegateControlTextGet(AddressOf clsHelper.ControlTextGet), New Object() {txtChannel}), String), _
                CType(txtHostPort.Invoke(New clsHelper.DelegateControlTextGet(AddressOf clsHelper.ControlTextGet), New Object() {txtHostPort}), Integer), _
                customData)

            End If
        Catch ex As Exception
            Debug.WriteLine(ex)
        End Try

    End Function
    Private Function GetAdminList() As String()
        Dim listadmin As ArrayList
        Dim name As String

        Try
            listadmin = New ArrayList
            For Each name In listUser.Items
                listadmin.Add(name)
            Next

            Return CType(listadmin.ToArray(GetType(String)), String())
        Catch ex As Exception
            Debug.WriteLine(ex)
            Return New String() {}
        End Try

    End Function
    Private Sub SendChat(ByVal chat As String)
        bnet.SendChatToQueue(New clsBNETChatMessage(TrimNewLine(chat), ProjectLainVersion, False))
    End Sub
    Private Sub SendChat(ByVal chat As String, ByVal owner As String, ByVal isPersistent As Boolean)
        bnet.SendChatToQueue(New clsBNETChatMessage(TrimNewLine(chat), owner, isPersistent))
    End Sub

    Private Sub ReflectEngineState()
        If bnet.IsEngineRunning() Then
            LockControl(True)
            labelStatus.Invoke(New clsHelper.DelegateControlBackColor(AddressOf clsHelper.ControlBackColor), New Object() {labelStatus, Color.CornflowerBlue})
            labelSummary.Invoke(New clsHelper.DelegateControlBackColor(AddressOf clsHelper.ControlBackColor), New Object() {labelSummary, Color.CornflowerBlue})
        Else
            LockControl(False)
            labelStatus.Invoke(New clsHelper.DelegateControlBackColor(AddressOf clsHelper.ControlBackColor), New Object() {labelStatus, Color.DarkGray})
            labelSummary.Invoke(New clsHelper.DelegateControlBackColor(AddressOf clsHelper.ControlBackColor), New Object() {labelSummary, Color.DarkGray})

        End If
    End Sub
    Private Sub LockControl(ByVal lock As Boolean)
        comboRealm.Invoke(New clsHelper.DelegateControlEnabled(AddressOf clsHelper.ControlEnabled), New Object() {comboRealm, Not lock})
        txtWar3.Invoke(New clsHelper.DelegateControlTextBoxReadOnly(AddressOf clsHelper.ControlTextBoxReadOnly), New Object() {txtWar3, lock})
        txtROC.Invoke(New clsHelper.DelegateControlTextBoxReadOnly(AddressOf clsHelper.ControlTextBoxReadOnly), New Object() {txtROC, lock})
        txtTFT.Invoke(New clsHelper.DelegateControlTextBoxReadOnly(AddressOf clsHelper.ControlTextBoxReadOnly), New Object() {txtTFT, lock})
        txtAccount.Invoke(New clsHelper.DelegateControlTextBoxReadOnly(AddressOf clsHelper.ControlTextBoxReadOnly), New Object() {txtAccount, lock})
        txtPassword.Invoke(New clsHelper.DelegateControlTextBoxReadOnly(AddressOf clsHelper.ControlTextBoxReadOnly), New Object() {txtPassword, lock})
        txtChannel.Invoke(New clsHelper.DelegateControlTextBoxReadOnly(AddressOf clsHelper.ControlTextBoxReadOnly), New Object() {txtChannel, lock})
        txtHostPort.Invoke(New clsHelper.DelegateControlTextBoxReadOnly(AddressOf clsHelper.ControlTextBoxReadOnly), New Object() {txtHostPort, lock})
        listChannelPeople.Invoke(New clsHelper.DelegateControlEnabled(AddressOf clsHelper.ControlEnabled), New Object() {listChannelPeople, lock})
        listUser.Invoke(New clsHelper.DelegateControlEnabled(AddressOf clsHelper.ControlEnabled), New Object() {listUser, Not lock})
        txtUserName.Invoke(New clsHelper.DelegateControlTextBoxReadOnly(AddressOf clsHelper.ControlTextBoxReadOnly), New Object() {txtUserName, lock})
        buttonUserAdd.Invoke(New clsHelper.DelegateControlEnabled(AddressOf clsHelper.ControlEnabled), New Object() {buttonUserAdd, Not lock})
        buttonUserRemove.Invoke(New clsHelper.DelegateControlEnabled(AddressOf clsHelper.ControlEnabled), New Object() {buttonUserRemove, Not lock})
    End Sub
    Private Function TrimNewLine(ByVal msg As String) As String
        Try
            Return msg.Replace(Environment.NewLine, "")
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Private Sub ExitProgram()
        Try
            SaveAll()
            bnet.Stop()
            Me.Enabled = False
            Me.Visible = True
            trayIcon.Visible = False
            Me.Dispose()
        Catch ex As Exception
            Debug.WriteLine(ex)
        End Try
    End Sub
    Private Sub aliveTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles aliveTimer.Elapsed
        Static isRunning As Boolean = False
        Try
            If isRunning = False Then
                isRunning = True

                If bnet.IsEngineRunning = False Then
                    If aliveCounter > 0 Then
                        aliveCounter = aliveCounter - 1
                        labelStatus.Invoke(New clsHelper.DelegateControlText(AddressOf clsHelper.ControlText), New Object() {labelStatus, String.Format("Auto-Reconnect will initialise in {0} seconds", aliveCounter)})
                    Else
                        Start()
                        aliveCounter = 60
                    End If
                Else
                    aliveCounter = 60

                    channelJoinCounter = (channelJoinCounter + 1) Mod 60
                    If channelJoinCounter = 0 AndAlso bnet.IsSendChatAble Then
                        If labelChannel.Text.ToLower = "the void" Then
                            bnet.SendChatToQueue(New clsBNETChatMessage(String.Format("/channel {0}", txtChannel.Text), txtAccount.Text, False))
                        Else
                            bnet.SendChatToQueue(New clsBNETChatMessage(String.Format("/rejoin"), txtAccount.Text, False))
                        End If

                    End If

                End If

                isRunning = False
            End If
        Catch ex As Exception
            Debug.WriteLine(ex)
        End Try
    End Sub


    Private Sub listUser_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles listUser.SelectedIndexChanged
        txtUserName.Text = CType(listUser.SelectedItem, String)
    End Sub
    Private Sub buttonUserAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonUserAdd.Click
        Dim name As String
        name = txtUserName.Text.Trim
        If name.Length > 0 AndAlso listUser.Items.Contains(name) = False Then
            listUser.Items.Add(name)
            txtUserName.Text = ""
        End If
    End Sub
    Private Sub buttonUserRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonUserRemove.Click
        Dim name As String
        name = txtUserName.Text.Trim
        If name.Length > 0 AndAlso listUser.Items.Contains(name) Then
            listUser.Items.Remove(name)
        End If
    End Sub
    Private Sub UpdateGameList()
        Dim host As clsGameHost
        Dim list As ArrayList

        list = New ArrayList
        SyncLock listHost.SyncRoot
            For Each host In listHost
                list.Add(host.GetGameName)
            Next
        End SyncLock

        listGame.Invoke(New clsHelper.DelegateControlListBoxDataSource(AddressOf clsHelper.ControlListBoxDataSource), New Object() {listGame, list.ToArray(GetType(String))})

    End Sub
    Private Sub buttonGameInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonGameInfo.Click
        Dim gamename As String
        Dim host As clsGameHost

        If listGame.SelectedItem Is Nothing = False AndAlso CType(listGame.SelectedItem, String).Length > 0 Then
            gamename = CType(listGame.SelectedItem, String)
            For Each host In listHost
                If host.GetGameName = gamename Then

                    MessageBox.Show(String.Format("Game: {0} currently with {1} Players, Game Started -> {2} ", host.GetGameName, host.GetTotalPlayers, host.GetIsCountDownStarted), ProjectLainVersion, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit For
                End If
            Next
        End If
    End Sub
    Private Sub buttonGameStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonGameStop.Click
        Dim gamename As String
        Dim host As clsGameHost

        If listGame.SelectedItem Is Nothing = False AndAlso CType(listGame.SelectedItem, String).Length > 0 Then
            gamename = CType(listGame.SelectedItem, String)
            For Each host In listHost
                If host.GetGameName = gamename Then
                    If MessageBox.Show(String.Format("Are You Sure You Want To Force Stop Game: {0} ?", host.GetGameName), ProjectLainVersion, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                        host.Dispose("Game Stopped Manually")
                    End If
                    Exit For
                End If
            Next
        End If
    End Sub

    Public Function GetLENPServer() As clsLENPServer
        Return lenp
    End Function



End Class







