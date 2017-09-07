Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading.Tasks
Imports System.Text.RegularExpressions ' Namespace untuk manipulasi registry
Imports System.Text

Imports Devart.Data
Imports Devart.Data.Oracle
Imports Devart.Common

Imports DevExpress
Imports DevExpress.XtraSplashScreen
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.BandedGrid
Imports DevExpress.XtraEditors.Repository

Imports AForge.Video

Public Class FrmConfigurationMenu
    Dim WithEvents Client As TCPCam.Client
    Dim WithEvents HOST As TCPCam.Host
    Dim nAction As String = ""
    Dim IdTabel As String = ""

    Dim Parameter As String = ""
    Dim frs As String

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
        'close
        Me.Close()
    End Sub

    Private Sub SimpleButton1_Click_1(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        'save
        If Not IsEmptyCombo({ComboBoxEdit1}) Then
            If Not IsEmptyText({TextEdit1, TextEdit2, TextEdit3, TextEdit4, TextEdit5}) Then
                If ComboBoxEdit1.Text = "LOKAL" Then
                    My.Settings.DBSourceLocal = TextEdit1.Text
                    My.Settings.Save()
                    My.Settings.DBVerLocal = TextEdit2.Text
                    My.Settings.Save()
                    My.Settings.DBNameLocal = TextEdit19.Text
                    My.Settings.Save()
                    My.Settings.DBPortLocal = TextEdit3.Text
                    My.Settings.Save()
                    My.Settings.DBUserLocal = TextEdit4.Text
                    My.Settings.Save()
                    My.Settings.DBPassLocal = TextEdit5.Text
                    My.Settings.Save()

                    CheckConLocal()
                    CloseConnLocal()
                End If
            End If
        End If
    End Sub

    Private Sub ComboBoxEdit1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxEdit1.SelectedIndexChanged
        If ComboBoxEdit1.Text = "LOKAL" Then
            LocalConfig()
        End If
    End Sub
    Private Sub LocalConfig()
        TextEdit1.Text = My.Settings.DBSourceLocal.ToString  'ipadress
        TextEdit3.Text = My.Settings.DBPortLocal.ToString    'ipport

        TextEdit19.Text = My.Settings.DBNameLocal.ToString     'db name
        TextEdit2.Text = My.Settings.DBVerLocal.ToString     'version

        TextEdit4.Text = My.Settings.DBUserLocal.ToString    'user
        TextEdit5.Text = My.Settings.DBPassLocal.ToString    'pass
    End Sub
    Private Sub CheckConLocal()
        GetConfig()
        If OpenConnLocal() = True Then
            MsgBox("Connection Successful", vbInformation, "Conection")
        Else
            MsgBox("Connection Failed", vbInformation, "Conection")
        End If
    End Sub
    Private Sub SimpleButton10_Click(sender As Object, e As EventArgs) Handles SimpleButton10.Click
        'CLOSE
        Me.Close()
    End Sub

    Private Sub SimpleButton13_Click(sender As Object, e As EventArgs) Handles SimpleButton13.Click
        'SAVE GENERAL CONFIG
        If Not IsEmptyText({TextEdit41}) Then
            SaveConfig()  'save ke my setting
            saveGConfig() 'save ke Database
            SimpleButton13.Enabled = False
            MsgBox("Save Succeeded", vbInformation, "Configuration")
            LockAll_GConfig()
        End If
    End Sub
    Private Sub saveGConfig()

        CompanyCode = My.MySettings.Default.CompanyCode
        Company = My.MySettings.Default.Company
        Dim MILLPLANT As String = My.MySettings.Default.Millplant
        Dim LocationSite As String = My.MySettings.Default.LocationSite
        Dim STORELOC1 As String = My.MySettings.Default.StoreLocation1
        Dim STORELOC2 As String = My.MySettings.Default.StoreLocation2
        Dim WBSETTING As String = My.MySettings.Default.ComportSetting
        WBCode = My.MySettings.Default.WBCode
        Dim IP_DERIVER As String = My.MySettings.Default.IPCamera1
        Dim IP_VEHICLE As String = My.MySettings.Default.IPCamera2
        Dim IPINDICATOR As String = My.MySettings.Default.IPIndicator
        Dim LRT As String = My.MySettings.Default.LoadingRampTransit

        Dim CONNECTION_LOC_NAME As String = My.MySettings.Default.DBNameLocal
        Dim CONNECTION_LOC_USER As String = My.MySettings.Default.DBUserLocal
        Dim CONNECTION_LOC_PASS As String = My.MySettings.Default.DBPassLocal
        Dim CONNECTION_LOC_IP As String = My.MySettings.Default.DBSourceLocal
        Dim CONNECTION_LOC_PORT As String = My.MySettings.Default.DBPortLocal
        Dim CONNECTION_LOC_VER As String = My.MySettings.Default.DBVerLocal

        Dim FFBCODE As String = ""
        Dim PKCODE As String = ""
        Dim CPOCODE As String = ""
        Dim SHELLCODE As String = ""
        Dim MM As String = ""
        Dim KTU As String = ""
        SAP = My.MySettings.Default.SAP

        SQL = "SELECT * FROM T_CONFIG WHERE COMPANYCODE='" & TextEdit41.Text & "' "

        If CheckRecord(SQL) = 0 Then
            SQL = "INSERT INTO T_CONFIG " +
            "(CompanyCode,COMPANY,WBCODE,WBSETTING,MILLPLANT, " +
            "STORELOC1,STORELOC2,IP_DERIVER,IP_VEHICLE,FFBCODE,CPOCODE,PKCODE,SHELLCODE,MM,KTU,LRT,SAP, " +
            "CONNECTION_LOC_NAME,CONNECTION_LOC_USER,CONNECTION_LOC_PASS,CONNECTION_LOC_IP,CONNECTION_LOC_PORT,CONNECTION_LOC_VER) " +
            "VALUES " +
            "('" & CompanyCode & "','" & Company & "','" & WBCode & "','" & WBSETTING & "','" & MILLPLANT & "', " +
            "'" & STORELOC1 & "','" & STORELOC2 & "','" & IP_DERIVER & "','" & IP_VEHICLE & "','" & FFBCODE & "','" & CPOCODE & "','" & PKCODE & "','" & SHELLCODE & "','" & MM & "','" & KTU & "','" & LRT & "','" & SAP & "', " +
            "'" & CONNECTION_LOC_NAME & "','" & CONNECTION_LOC_USER & "','" & CONNECTION_LOC_PASS & "','" & CONNECTION_LOC_IP & "','" & CONNECTION_LOC_PORT & "','" & CONNECTION_LOC_VER & "') "
        Else
            SQL = "Update T_CONFIG SET CompanyCode='" & CompanyCode & "', " +
                " Company='" & Company & "'," +
                " WBCode='" & WBCode & "'," +
                " WBSETTING='" & WBSETTING & "', " +
                " MILLPLANT='" & MILLPLANT & "'," +
                " STORELOC1='" & STORELOC1 & "', " +
                " STORELOC2='" & STORELOC2 & "'," +
                " IP_DERIVER='" & IP_DERIVER & "', " +
                " IP_VEHICLE='" & IP_VEHICLE & "'," +
                " FFBCODE='" & FFBCODE & "', " +
                " CPOCODE='" & CPOCODE & "'," +
                " PKCODE='" & PKCODE & "', " +
                " SHELLCODE='" & SHELLCODE & "'," +
                " MM='" & MM & "', " +
                " KTU='" & KTU & "'," +
                " LRT='" & LRT & "', " +
                " SAP='" & SAP & "'," +
                " CONNECTION_LOC_NAME='" & CONNECTION_LOC_NAME & "', " +
                " CONNECTION_LOC_USER='" & CONNECTION_LOC_USER & "'," +
                " CONNECTION_LOC_PASS='" & CONNECTION_LOC_PASS & "', " +
                " CONNECTION_LOC_IP='" & CONNECTION_LOC_PASS & "'," +
                " CONNECTION_LOC_PORT='" & CONNECTION_LOC_PORT & "', " +
                " CONNECTION_LOC_VER='" & CONNECTION_LOC_VER & "'" +
                " WHERE CompanyCode ='" & TextEdit4.Text & "'"
        End If
        ExecuteNonQuery(SQL)
    End Sub


    Private Sub SimpleButton12_Click(sender As Object, e As EventArgs) Handles SimpleButton12.Click
        'CLOSE GENERAL CONFIG
        Close()
    End Sub
    Public Sub LoadConfig()
        TextEdit41.Text = My.MySettings.Default.CompanyCode.Trim.ToString
        TextEdit40.Text = My.MySettings.Default.Company  'My.Settings.Company
        TextEdit39.Text = My.MySettings.Default.Millplant
        TextEdit38.Text = My.MySettings.Default.LocationSite
        TextEdit37.Text = My.MySettings.Default.StoreLocation1
        TextEdit36.Text = My.MySettings.Default.StoreLocation2
        TextEdit35.Text = My.MySettings.Default.ComportSetting
        ComboBoxEdit6.Text = My.MySettings.Default.WBCode

        ComboBoxEdit7.Text = My.MySettings.Default.IPCamera1
        ComboBoxEdit8.Text = My.MySettings.Default.IPCamera2
        TextEdit31.Text = My.MySettings.Default.IPIndicator
        ComboBoxEdit3.Text = My.MySettings.Default.LoadingRampTransit
        ComboBoxEdit4.Text = My.MySettings.Default.SAP
    End Sub

    Public Sub SaveConfig()
        My.Settings.CompanyCode = TextEdit41.Text
        My.Settings.Save()
        My.Settings.Company = TextEdit40.Text
        My.Settings.Save()
        My.Settings.Millplant = TextEdit39.Text
        My.Settings.Save()
        My.Settings.LocationSite = TextEdit38.Text
        My.Settings.Save()
        My.Settings.StoreLocation1 = TextEdit37.Text
        My.Settings.Save()
        My.Settings.StoreLocation2 = TextEdit36.Text
        My.Settings.Save()
        My.Settings.ComportSetting = TextEdit35.Text
        My.Settings.Save()
        My.Settings.WBCode = ComboBoxEdit6.Text
        My.Settings.Save()
        My.Settings.IPCamera1 = ComboBoxEdit7.Text
        My.Settings.Save()
        My.Settings.IPCamera2 = ComboBoxEdit8.Text
        My.Settings.Save()
        My.Settings.IPIndicator = TextEdit31.Text
        My.Settings.Save()
        My.Settings.LoadingRampTransit = ComboBoxEdit3.Text
        My.Settings.Save()
        My.Settings.SAP = ComboBoxEdit4.Text
        My.Settings.Save()
    End Sub

    Private Sub BackstageViewTabItem1_SelectedChanged(sender As Object, e As DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs) Handles BackstageViewTabItem1.SelectedChanged
        LoadConfig()
    End Sub

    Private Sub FrmConfigurationMenu_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = "CONFIGURATION MENU"
        BackstageViewTabItem1.Selected = True
        If TextEdit41.Text <> "" Then LockAll_GConfig()
        If My.MySettings.Default.SAP.ToString = "Y" Then
            FillWB()   'LOAD WB
            FillCctv() 'LOAD CCTV  
        End If
    End Sub
    Private Sub FillWB()
        SQL = "Select DISTINCT NAMA, KDNAMA FROM M_WB ORDER BY KDNAMA"
        FILLComboBoxEdit(SQL, 0, ComboBoxEdit6, False)
    End Sub
    Private Sub FillCctv()
        SQL = "Select DISTINCT NAMA, KDNAMA FROM M_CCTV ORDER BY KDNAMA"
        FILLComboBoxEdit(SQL, 0, ComboBoxEdit7, False)
        FILLComboBoxEdit(SQL, 0, ComboBoxEdit8, False)
    End Sub


    Private Sub SimpleButton15_Click(sender As Object, e As EventArgs) Handles SimpleButton15.Click
        'EDIT GENERAL CONFIG
        If TextEdit41.Text = "" Then SimpleButton11.Text = "Update"
        UnLockAll_GConfig()
    End Sub

    Private Sub UnLockAll_GConfig()
        TextEdit41.Enabled = True
        TextEdit40.Enabled = True
        TextEdit39.Enabled = True
        TextEdit38.Enabled = True
        TextEdit37.Enabled = True
        TextEdit36.Enabled = True
        TextEdit35.Enabled = True
        ComboBoxEdit6.Enabled = True
        ComboBoxEdit7.Enabled = True
        ComboBoxEdit8.Enabled = True
        TextEdit31.Enabled = True
        ComboBoxEdit3.Enabled = True
        ComboBoxEdit4.Enabled = True
        SimpleButton13.Enabled = True 'SAVE
        SimpleButton15.Enabled = False 'EDIT
    End Sub

    Private Sub LockAll_GConfig()
        TextEdit41.Enabled = False
        TextEdit40.Enabled = False
        TextEdit39.Enabled = False
        TextEdit38.Enabled = False
        TextEdit37.Enabled = False
        TextEdit36.Enabled = False
        TextEdit35.Enabled = False
        ComboBoxEdit6.Enabled = False
        ComboBoxEdit7.Enabled = False
        ComboBoxEdit8.Enabled = False
        TextEdit31.Enabled = False
        ComboBoxEdit3.Enabled = False
        ComboBoxEdit4.Enabled = False
        SimpleButton13.Enabled = False 'SAVE
        SimpleButton15.Enabled = True 'EDIT
    End Sub


#Region "Indicator"
    Private Sub AppendOutput(message As String) ''cetak hasil baca data
        If TxtIndikator.InvokeRequired Then
            Dim x As New SetTextCallback(AddressOf AppendOutput)
            Invoke(x, New Object() {(Text)})
        Else
            TxtIndikator.Text = CType(num(message), String)
            If GETwEIGHT = True Then
                WEIGHT = TxtIndikator.Text
            End If
        End If
    End Sub
    Private Sub DoAcceptClient(result As IAsyncResult)
        Dim monitorInfo As MonitorInfo = CType(_ConnectionMontior.AsyncState, MonitorInfo)
        If monitorInfo.Listener IsNot Nothing AndAlso Not monitorInfo.Cancel Then
            Dim info As ConnectionInfo = CType(result.AsyncState, ConnectionInfo)
            monitorInfo.Connections.Add(info)
            info.AcceptClient(result)
            ListenForClient(monitorInfo)
            info.AwaitData()
            Dim doUpdateConnectionCountLabel As New Action(AddressOf UpdateConnectionCountLabel)
            Invoke(doUpdateConnectionCountLabel)
        End If
    End Sub
    Private Sub DoMonitorConnections()
        'Create delegate for updating output display
        Dim doAppendOutput As New Action(Of String)(AddressOf AppendOutput)
        'Create delegate for updating connection count label
        Dim doUpdateConnectionCountLabel As New Action(AddressOf UpdateConnectionCountLabel)

        'Get MonitorInfo instance from thread-save Task instance
        Dim monitorInfo As MonitorInfo = CType(_ConnectionMontior.AsyncState, MonitorInfo)
        'Report progress
        'Implement client connection processing loop
        Do
            'Create temporary list for recording closed connections
            Dim lostCount As Integer = 0
            'Examine each connection for processing
            For index As Integer = monitorInfo.Connections.Count - 1 To 0 Step -1
                Dim info As ConnectionInfo = monitorInfo.Connections(index)
                If info.Client.Connected Then
                    'Process connected client
                    If info.DataQueue.Count > 0 Then
                        'The code in this If-Block should be modified to build 'message' objects
                        'according to the protocol you defined for your data transmissions.
                        'This example simply sends all pending message bytes to the output textbox.
                        'Without a protocol we cannot know what constitutes a complete message, so
                        'with multiple active clients we could see part of client1's first message,
                        'then part of a message from client2, followed by the rest of client1's
                        'first message (assuming client1 sent more than 64 bytes).
                        Dim messageBytes As New List(Of Byte)
                        While info.DataQueue.Count > 0
                            Dim value As Byte
                            If info.DataQueue.TryDequeue(value) Then
                                messageBytes.Add(value)
                            End If
                        End While
                        Invoke(doAppendOutput, System.Text.Encoding.ASCII.GetString(messageBytes.ToArray))
                        ' cacah(System.Text.Encoding.ASCII.GetString(messageBytes.ToArray))
                    End If
                Else
                    'Clean-up any closed client connections
                    monitorInfo.Connections.Remove(info)
                    lostCount += 1
                End If
            Next
            If lostCount > 0 Then
                Invoke(doUpdateConnectionCountLabel)
            End If
            'Throttle loop to avoid wasting CPU time
            _ConnectionMontior.Wait(1)
        Loop While Not monitorInfo.Cancel
        'Close all connections before exiting monitor
        For Each info As ConnectionInfo In monitorInfo.Connections
            info.Client.Close()
        Next
        monitorInfo.Connections.Clear()
        'Update the connection count label and report status
        Invoke(doUpdateConnectionCountLabel)
        'Me.Invoke(doAppendOutput, "Monitor Stopped.")
    End Sub
    Private Sub ListenForClient(monitor As MonitorInfo)
        Dim info As New ConnectionInfo(monitor)
        _Listener.BeginAcceptTcpClient(AddressOf DoAcceptClient, info)
    End Sub
    Private Sub SimpleButton7_Click_1(sender As Object, e As EventArgs)
        Me.Close()
    End Sub



#End Region
End Class

