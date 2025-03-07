Imports System.IO
Imports System.Net.Sockets.Socket
Imports System
Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Imports System.Drawing.Drawing2D

' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
Public Class Form1
    Inherits System.Windows.Forms.Form
    Public receivingUdpClient As UdpClient
    Public RemoteIpEndPoint As New System.Net.IPEndPoint(System.Net.IPAddress.Any, 0)
    Public ThreadReceive As System.Threading.Thread
    Dim SocketNO As Integer
    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    'Public dt As DataTable = GetTable()           ' History table (same structure)
    'Public dt1 As DataTable = GetTable()          ' Decision table (same structure)

    Dim ArrayIndex As Integer

    Public Structure recCarpark
        Public stName As String
        Public stIP As String
        Public stMsg As String
        Public OldPriority As Integer
        Public nPriority As Integer       ' Red, Orange, Green
        Public stID As String
        Public status() As String        ' to keep 0,0,0,0,0,0,0,0,0,0,0,0,0  13 12 bit Status message value for each carpark
        Public update_dt As String       ' to save last update date time 
        Public GroupID As String
        Public GroupIndex As Integer
        Public Operation As String     ' to store Operatin or Not Operation Yet
        Public y_value As Integer
        Public x_value As Integer
        Public lot_available As String 'Integer
    End Structure

    Public Carparks() As recCarpark

    ' ==============================================
    ' Configuration for screen layout
    Dim TotalCarpark As Integer = 1
    Dim n As Integer = 0
    Dim m As Integer = 0
    Dim LeftMargin As Integer = 10
    Dim H_Spacing As Integer = 20
    Dim V_Spacing As Integer = 20
    Dim TopMargin As Integer = 40
    Dim ColPerRow As Integer = 9

    ' ==============================================
    Private B() As ctlCarpark                       ' control array for display carpark
    Private Legend() As ctlCarpark         ' to show legend
    Dim IndexofRed As Integer
    Dim TotalLegend As Integer = 4
    Dim BlinkCounter As Integer = 0
    Dim SoundAlarmCounter As Integer = 0

    Dim MaxCountOfBlink As Integer = 15 '30      think: 30 is to much blinking, so set 15 is ok
    Dim StillBlinking As Boolean = False
    Dim detectRed As Boolean = False

    'Dim random As New Random
    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        ' Call the base class
        MyBase.OnPaint(e)

        Dim I As Integer = 570
        'Dim I As Integer = (11 * B(0).Height) + (V_Spacing * 12)
        If 1 = 2 Then
            e.Graphics.DrawLine(Pens.DarkGreen, 10, picSJS.Height + 15, Me.Width - 10, picSJS.Height + 15)
            e.Graphics.DrawLine(Pens.DarkGreen, 10, I + 20, Me.Width - 10, I + 20)
        End If
        
    End Sub

    '' to store incoming message as the records of variable count
    'Function GetTable() As DataTable
    '    ' Create new DataTable instance.
    '    Dim table1 As New DataTable
    '    ' Create four typed columns in the DataTable.
    '    table1.Columns.Add("CarparkID", GetType(String))     ' CarparkID Assign by SJS
    '    table1.Columns.Add("Message", GetType(String))
    '    table1.Columns.Add("StationName", GetType(String))
    '    table1.Columns.Add("StationID", GetType(Integer))
    '    table1.Columns.Add("Command", GetType(String))
    '    table1.Columns.Add("Type", GetType(String))
    '    table1.Columns.Add("Update_dt", GetType(DateTime))   ' record INSERT/UPDATE dt
    '    table1.Columns.Add("CarparkName", GetType(String))   ' CarparkID Assign by HDB
    '    table1.Columns.Add("Priority", GetType(Integer))  ' Keep respective priority (3Red,2Orange,1Green)

    '    ' Add five rows with those columns filled in the DataTable.
    '    'table.Rows.Add(25, "Indocin", "David", DateTime.Now)
    '    'table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now)
    '    'table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now)
    '    'table.Rows.Add(21, "Combivent", "Janet", DateTime.Now)
    '    'table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now)

    '    Return table1
    'End Function

    Private Sub Read_Setting()
        Dim st() As String
        Dim I As Integer = 0
        Dim K As Integer = 0
        Dim J As Integer = 0
        Dim stCarpark As String = ""
        Dim stTemp As String = ""
        Dim stID As String = ""
        Dim st1 As String = ""
        Dim n As Integer = 0

        gstVersion1 = Me.GetType.Assembly.GetName.Version.ToString
        With System.Configuration.ConfigurationManager.AppSettings

            gstlogprefix = .Get("gstlogprefix")
            Call WriteLog("*******************************************************************")
            Call WriteLog("****************** Application start " & gstVersion1 & " ======================")
            Call WriteLog("*******************************************************************")
            Call WriteLog("gstlogprefix: " & gstlogprefix)
            gnTotalCarpark = .Get("gnTotalCarpark")

            If gnTotalCarpark > 120 Or gnTotalCarpark < 1 Then
                Call WriteLog("Total Carpark Error: " & gnTotalCarpark)
                Call WriteLog("Max limit is 120 for STEE")
                End
            Else
                Call WriteLog("Total Carpark: " & gnTotalCarpark & ", Max limit is 120")
            End If

            gnColPerRow = .Get("gnColPerRow") : Call WriteLog("gnColPerRow: " & gnColPerRow)
            gnCarparkWidth = .Get("gnCarparkWidth") : Call WriteLog("gnCarparkWidth: " & gnCarparkWidth)
            gnCarparkHeight = .Get("gnCarparkHeight") : Call WriteLog("gnCarparkHeight: " & gnCarparkHeight)
            gbHideCarparks = .Get("gbHideCarparks") : Call WriteLog("gbHideCarparks: " & gbHideCarparks)

            LeftMargin = .Get("LeftMargin") : Call WriteLog("LeftMargin: " & LeftMargin)
            H_Spacing = .Get("H_Spacing") : Call WriteLog("H_Spacing: " & H_Spacing)
            V_Spacing = .Get("V_Spacing") : Call WriteLog("V_Spacing: " & V_Spacing)
            TopMargin = .Get("TopMargin") : Call WriteLog("TopMargin: " & TopMargin)
            Radius = .Get("Radius") : Call WriteLog("Radius: " & Radius)

            Call WriteLog("gbHideCarparks: " & gbHideCarparks)

            gstUDPPort = .Get("gstUDPPort")
            Call WriteLog("gstUDPPort: " & gstUDPPort)
            gbDebugPrint = .Get("gbDebugPrint")
            Call WriteLog("gbDebugPrint: " & gbDebugPrint)
            gbDetailPrint = .Get("gbDetailPrint")
            Call WriteLog("gbDetailPrint: " & gbDetailPrint)

            gnTotalMessage = .Get("gnTotalMessage")
            Call WriteLog("gnTotalMessage: " & gnTotalMessage)
            gnWaitingMinute = .Get("gnWaitingMinute")
            Call WriteLog("gnWaitingMinute: " & gnWaitingMinute)
            Intercom_Minute = .Get("Intercom_Minute")
            Call WriteLog("Intercom_Minute: " & Intercom_Minute)

            gbCarparkTimerEnable = .Get("gbCarparkTimerEnable")
            tmrCarpark.Enabled = gbCarparkTimerEnable
            Call WriteLog("gbCarparkTimerEnable: " & gbCarparkTimerEnable)
            gbCarparkTimerInterval = .Get("gbCarparkTimerInterval")
            Call WriteLog("gbCarparkTimerInterval: " & gbCarparkTimerInterval)
            tmrCarpark.Interval = gbCarparkTimerInterval
            Call WriteLog("Carpark Timer Interval: " & tmrCarpark.Interval.ToString)
            gbNoConnectionShowRedEnable = .Get("gbNoConnectionShowRedEnable")
            Call WriteLog("gbNoConnectionShowRedEnable: " & gbNoConnectionShowRedEnable)

            gstRGB_Value_Orange = .Get("gstRGB_Value_Orange")
            Call WriteLog("gstRGB_Value_Orange" & gstRGB_Value_Orange)
            st = Split(gstRGB_Value_Orange, ",") : Orange1.R = st(0) : Orange1.G = st(1) : Orange1.B = st(2)

            gstRGB_Value_Red = .Get("gstRGB_Value_Red")
            Call WriteLog("gstRGB_Value_Red" & gstRGB_Value_Red)
            st = Split(gstRGB_Value_Red, ",") : Red1.R = st(0) : Red1.G = st(1) : Red1.B = st(2)

            gstRGB_Value_Black = .Get("gstRGB_Value_Black")
            Call WriteLog("gstRGB_Value_Black" & gstRGB_Value_Black)
            st = Split(gstRGB_Value_Black, ",") : Black1.R = st(0) : Black1.G = st(1) : Black1.B = st(2)

            gstRGB_Value_Green1 = .Get("gstRGB_Value_Green1")
            Call WriteLog("gstRGB_Value_Green1" & gstRGB_Value_Green1)
            st = Split(gstRGB_Value_Green1, ",") : Green1.R = st(0) : Green1.G = st(1) : Green1.B = st(2)

            gstRGB_Value_Green2 = .Get("gstRGB_Value_Green2")
            Call WriteLog("gstRGB_Value_Green2" & gstRGB_Value_Green2)
            st = Split(gstRGB_Value_Green2, ",") : Green2.R = st(0) : Green2.G = st(1) : Green2.B = st(2)

            gstRGB_Value_Green3 = .Get("gstRGB_Value_Green3")
            Call WriteLog("gstRGB_Value_Green3" & gstRGB_Value_Green3)
            st = Split(gstRGB_Value_Green3, ",") : Green3.R = st(0) : Green3.G = st(1) : Green3.B = st(2)

            TotalGroup = .Get("TotalGroup")
            If TotalGroup > 3 Or TotalGroup < 1 Then
                Call WriteLog("TotalGroup Error")
                End
            End If

            For I = 0 To 2
                If I < TotalGroup Then
                    st1 = .Get("Group_" & (I + 1).ToString)
                    GroupID(I) = st1
                Else
                    GroupID(I) = ""
                End If
            Next

            gstMapFolder = .Get("gstMapFolder")
            gstServer = .Get("gstServer")
            gstDatabase = .Get("gstDatabase")
            gstUltraVNC_Path = .Get("gstUltraVNC_Path")

            lblNotify1 = .Get("lblNotify1")
            btnReset1 = .Get("btnReset1")
            btnPath1 = .Get("btnPath1")
            txtPath1 = .Get("txtPath1")

            ' Connection String for Central DB
            LocalconStr = "Data Source=" & gstServer & ";Password=yzhh2007;"
            LocalconStr = LocalconStr & "User ID=sa;Initial Catalog=" & gstDatabase & ""

            'If DB_Test() Then
            '    MsgBox("DB OK")
            'Else
            '    MsgBox("DB ERROR")
            'End If

            Call WriteLog("Server: " & gstServer)
            Call WriteLog("Database: " & gstDatabase)
            Call WriteLog("gstUltraVNC_Path: " & gstUltraVNC_Path)
            Call WriteLog("gstMapFolder: " & gstMapFolder)

            Call WriteLog("***************************************************")
            ' Alarm message have 2 columns: (1) Message Value (2) Priority
            Call WriteLog("Reading Alarm Message Setting")
            Call WriteLog("gnTotalMessage: " & gnTotalMessage)
            ReDim ArrayAlarmMessage(gnTotalMessage - 1)
            For I = 0 To gnTotalMessage - 1
                stTemp = .Get("gstMessage_" & (I + 1).ToString)
                Call WriteLog("Message " & (I + 1).ToString & " : " & stTemp)
                st = Split(stTemp, ",")
                ArrayAlarmMessage(I).Message = st(0)
                ArrayAlarmMessage(I).Priority = st(1)
            Next

            gnTotalStatus = .Get("gnTotalStatus")
            Call WriteLog("***************************************************")
            ' Status message have 2 columns: 

            Call WriteLog("Reading Status Setting")
            Call WriteLog("gnTotalStatus: " & gnTotalStatus)
            ' gnTotalStatus = 12, ReDim array size = 12 index from 0 to 12
            ' Index 0 is for No Problem
            ReDim ArrayStatusMessage(gnTotalStatus) ' 12
            ' index from 0 to 11 for 12 Status
            ArrayStatusMessage(0).Value1 = "0" ' Zero is ok 
            ' All Non-Zero value are remark as ERROR
            ArrayStatusMessage(0).Message = "No Error"
            For I = 1 To gnTotalStatus
                stTemp = .Get("gstStatus_" & I.ToString)
                Call WriteLog("Status " & I.ToString & " : " & stTemp)
                ' st = Split(stTemp, ",")
                ArrayStatusMessage(I).Value1 = "0" ' Zero is ok
                ' All Non-Zero value are remark as ERROR
                ArrayStatusMessage(I).Message = stTemp
            Next
        End With

        txtSocket.Text = gstUDPPort

        TotalCarpark = gnTotalCarpark

        ReDim Carparks(TotalCarpark - 1)        ' Note: Index start with zero. 0 to TotalCarpark -1
        Call WriteLog("***************************************************")
        Call WriteLog("Reading Total Carparks")
        Call WriteLog("gnTotalCarpark: " & gnTotalCarpark)
        With System.Configuration.ConfigurationManager.AppSettings
            For I = 0 To TotalCarpark - 1
                stTemp = .Get("gstCarpark_" & (I + 1).ToString)

                Call WriteLog("Carpark " & (I + 1).ToString & " : " & stTemp)
                st = Split(stTemp, ",") ' index start from 1 in setting file

                If st(0) = "" Then
                    MsgBox("Invalid Carpark name found for Carpark: " & I + 1, MsgBoxStyle.Critical, "Setting Error: App stop.")
                    Call WriteLog("Invalid Carpark name found for Carpark: " & I + 1)
                    End
                End If
                ' Some validation need to do here
                ' If st is blank, show error and let them check setting file
                ' also need to check duplicate entry of carpark name and ip address
                stCarpark = UCase(st(0))              ' Carpark Name  (eg. U6, BBM7, etc.........)
                stID = UCase(st(2))                   ' Carpark ID  (eg. 761, 793, etc........)
                If Not IsNumeric(stID) Then
                    MsgBox("Invalid Carpark ID : " & stID & " found at " & (I + 1).ToString, MsgBoxStyle.Critical, "Setting Error: App stop.")
                    End
                End If

                For J = 0 To I
                    If Carparks(J).stName = stCarpark Then
                        MsgBox("Found duplicate carpark name: " & stCarpark, MsgBoxStyle.Critical, "Setting Error: App stop.")
                        End
                    End If
                Next

                Carparks(I).GroupID = UCase(st(3))
                For n = 0 To TotalGroup - 1
                    If Carparks(I).GroupID = GroupID(n) Then
                        Carparks(I).GroupIndex = n
                        Exit For
                    End If
                Next
                Carparks(I).Operation = UCase(st(4))          ' Operation or Not Operation Yet
                Carparks(I).stName = stCarpark          ' carpark name change to Upper Case
                'st1 = UCase(Mid(stCarpark, 1, 1)) & LCase(Mid(stCarpark, 2, stCarpark.Length - 1))
                'Carparks(I).stName = st1

                ' change of total STATUS msg is very important (from 12 to 13)
                ReDim Carparks(I).status(12) '(11)   ' Note: 12 bit status index from 0 to 11
                For K = 0 To 12 '11                ' 0 is OK, non-zero is ERROR
                    Carparks(I).status(K) = 0         ' init by 0 is OK, non zero means ERROR
                Next
                'Carparks(I).status(6) = -1    ' test
                For J = 0 To I
                    If Carparks(J).stID = stID Then
                        MsgBox("Found duplicate carpark id: " & stID, MsgBoxStyle.Critical, "Setting Error: App stop.")
                        End
                    End If
                Next

                Carparks(I).stID = stID          ' carpark name change to Upper Case
                ' also need to check whether IP Address is valid or not by using regular expression
                Carparks(I).stIP = st(1)            ' ip address
                Carparks(I).stMsg = "ok" 'st(2)           ' initial message
                Carparks(I).OldPriority = 1
                Carparks(I).nPriority = 1   ' value init to 1      ' 1 is green , 2 for orange, 3 for red
                Carparks(I).update_dt = Now

                Carparks(I).y_value = st(5)
                Carparks(I).x_value = st(6)
                Carparks(I).lot_available = st(7)

                'Call Save_To_Connection_Table(stID, 0)   ' 0 for PMS Computer
                If Carparks(I).Operation = "OPERATION" Then
                    dtConn.Rows.Add(stID, 0, Now, stCarpark)    ' Every carpark must have station 0 (PMS)
                    dtConn.Rows.Add(stID, 1, Now, stCarpark)    ' Assumption: Every carpark must have station 1
                    dtConn.Rows.Add(stID, 2, Now, stCarpark)    ' Assumpiton: Every carpark must have station 2
                    ' Assumption: Every carpark must have at least 2 stations and 1 PMS
                End If
            Next

            ' Below is newly improved array to handle status message (13 numbers)
            Call WriteLog("Reading newly improved Status array")
            Total_Status_Count = .Get("Total_Status_Count")
            Call WriteLog("Total_Status_Count = " & Total_Status_Count)
            Status_Array_Size = .Get("Status_Array_Size")
            Call WriteLog("Status_Array_Size = " & Status_Array_Size)
            ReDim NewStatusArray(Status_Array_Size - 1)

            For I = 0 To Status_Array_Size - 1
                st1 = .Get("Status_" & (I + 1).ToString)
                st = Split(st1, ",")
                NewStatusArray(I).status1 = st(0)
                NewStatusArray(I).value1 = st(1)
                NewStatusArray(I).message1 = st(2)
                NewStatusArray(I).priority1 = st(3)
                Call WriteLog((I + 1).ToString & " : " & st1)
            Next

            If 1 = 2 Then
                Sorting_Carparks()
            End If

        End With
        ColPerRow = gnColPerRow
    End Sub

    Private Sub Locate2()
        Dim st As String = ""

        st = gstVersion1.Substring(0, gstVersion1.Length - 4)
        Me.Text = Me.Text & "Version " & gstVersion1   'st
        'Me.BackColor = ColorTranslator.FromOle(RGB(Black1.R, Black1.G, Black1.B)) 'Color.Black '
        'Me.BackColor = ColorTranslator.FromOle(RGB(38, 38, 38)) 'Color.Black '
        'Me.BackColor = Color.DarkGray

        With lblTitle
            .Left = (Me.Width - .Width + (picSJS.Width)) / 2
            .Top = 5
        End With

        With picSJS
            .Top = 2
            .Left = lblTitle.Left - .Width - 5
            .Visible = True
        End With

        With lblColorCode
            .Top = Me.Height - .Height - 130 '140
            .Left = 1000 '5
        End With

        With lstConnection
            .Top = lblColorCode.Top + lblColorCode.Height + 25
            .Left = 10 ' (Me.Width - .Width) / 2
        End With
        With lstIntercom
            .Top = lblColorCode.Top + lblColorCode.Height + 25
            .Left = lstConnection.Width + 400 '200   '(Me.Width - .Width) / 2
        End With
    End Sub

    Private Sub Form1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click


        Dim J As Integer
        Dim st() As String
        ' Update lot info in label, but keep the same carpark name also
        If 1 = 2 Then
            If 1 = 2 Then
                st = Split(B(0).G(0).lblCarpark.Text, " ")
                For J = 0 To TotalGroup + 2    '1
                    B(0).G(J).lblCarpark.Text = st(0).Trim & " 391"
                    'B(0).G(J).Tag 
                Next
            Else
                MsgBox("X = " & Cursor.Position.X, MsgBoxStyle.Information, "Y = " & Cursor.Position.Y)
            End If
        End If

    End Sub

    Private Sub update_Carpark_Lot_Available_Number(cpindex As Integer, lot2 As Integer)
        Dim J As Integer
        Dim st2 As String = ""
        Dim st3 As String = ""
        Dim st() As String
        Dim K As Integer = 0
        Dim X As Integer = 0
        Dim Y As Integer = 0
        Dim lastIndex As Integer = TotalGroup + 2
        For J = 0 To TotalGroup + 1    '2    ' 1
            st2 = B(cpindex).G(J).lblCarpark.Text.Trim
            st = Split(st2, " ")
            st3 = st(0).Trim
            K = st3.Length   ' check carpark length, max is 4 characters
            ' NOTE: Instead of handling like this, it may be better to use vbCrLf in between Carpark Name and Lot Number (should try)
            If K = 2 Then
                X = 1
                Y = 1
            End If
            B(cpindex).G(J).lblCarpark.Text = Space(X) & st3 & Space(Y) & " " & lot2.ToString   '391"
            'B(0).G(J).Tag 
        Next
        B(cpindex).G(lastIndex).lblCarpark.Text = Space(X) & st3 & Space(Y) & " NA" '& lot2.ToString   '391"
    End Sub

    Private Sub read_Carpark_Lot_Available_Number()
        Dim st As String = ""
        Dim sql As String = ""
        Dim I As Integer = 0
        Dim b1 As Boolean = True
        Dim dt As DataTable = Nothing
        Dim count1 As Integer = 0
        Dim carpark1 As Integer = 0
        Dim lot1 As Integer = 0

        ' IMPORTANT NOTE: View_Carpark_Name_Index_Map_Look_Up is hardcoded, 
        ' If more new carparks are added, this VIEW also need to update accordingly
        ' Because we Use INNER JOIN, so if not update the new carpark may not be in the list
        ' Carpark_Index value in the view must be the same order which appear in the config file of this application
        ' =========================================================================
        'sql = "SELECT a.[Carpark_Number], a.[Lots_Available], b.[carpark_index] "
        'sql = sql & "FROM [Carpark_Lots_Availablity] a "
        'sql = sql & "INNER JOIN [View_Carpark_Name_Index_Map_Look_Up] b "
        'sql = sql & "ON a.Carpark_Number = b.[carpark_name] "
        'sql = sql & "ORDER BY b.[carpark_index]"

        ' above sql query statement is changed to View and read from this View directly
        ' and in future, we can modify view directly without affecting source code

        sql = "SELECT [Carpark_Number], [Lots_Available], [carpark_index] "
        sql = sql & "FROM [View_JTC_Workstation]"    ' [View_1_for_Test]"

        dt = GetData(sql, LocalconStr)
        If dt Is Nothing Then
            Call WriteLog("Datatable is Nothing, EXIT SUB")
            Call WriteLog("sql: " & sql)
            Exit Sub
        Else
            count1 = dt.Rows.Count
            If count1 > 0 Then
                ' update 
                Call WriteLog("Update lot start, total carpark = " & count1)
                'Call WriteLog("SQL = " & sql)

                For I = 0 To count1 - 1
                    carpark1 = dt.Rows.Item(I).Item(2)   '("Carpark_Index")     ' no need to use Carpark_Number in this case
                    lot1 = dt.Rows.Item(I).Item("Lots_Available")
                    Call update_Carpark_Lot_Available_Number(carpark1 - 1, lot1)
                Next
                Call WriteLog("Update lot finish")
            Else
                Call WriteLog("Total retrieve row(s) = " & count1)
                Call WriteLog("sql: " & sql)
            End If
        End If
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            End
        End If
    End Sub

    Private Sub Locate_New_Control()
        Dim position() As String

        position = Split(lblNotify1, ",")
        lblNotify.Left = position(0)
        lblNotify.Top = position(1)

        position = Split(btnReset1, ",")
        btnReset.Left = position(0)
        btnReset.Top = position(1)

        position = Split(btnPath1, ",")
        btnPath.Left = position(0)
        btnPath.Top = position(1)

        position = Split(txtPath1, ",")
        txtPath.Left = position(0)
        txtPath.Top = position(1)

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim I As Integer = 0
        Dim J As Integer = 0
        Dim K As Integer = 0

        KeyPreview = True
        'Me.Visible = False
        'Dim s = New SplashScreen1()
        's.Show()
        ''Do processing here or thread.sleep to illustrate the concept
        'System.Threading.Thread.Sleep(9000)
        's.Close()
        'Me.Visible = True

        Call Read_Setting()          ' Carparks Array is handled in this Routine
        Call Locate_New_Control()  ' lblNotify, btnReset, btnPath, txtPath
        Call Locate2()

        If gbHideCarparks Then            ' hide carparks is only for testing to see UDP incoming message
            GroupBox1.Visible = True
            TextBox2.Visible = True
            chkBit.Checked = True
        Else
            chkBit.Checked = False
            GroupBox1.Visible = False
            TextBox2.Visible = False
            Button1_Click(sender, e)
        End If

        Call Show_Carpark()   ' min = 1, max = 100
        'B(0).image1 = My.Resources.Resource1.red
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        ' below is showing legend at the bottom of screen

        If False Then
            Call Show_Legend()
        End If

        'Me.BackgroundImage = Image.FromFile("d:\abc\map1.png")    'jpg")  'png")   'jpg")
        Me.BackgroundImage = Image.FromFile(gstMapFolder)  ' & "\map1.png")
        Me.BackgroundImageLayout = ImageLayout.Stretch
        'Call Show_Blinking()   ' for testing

        If 1 = 2 Then
            Call Sync_Carpark_Index_Between_Setting_File_And_DB_Lot_Available()

            Call read_Carpark_Lot_Available_Number()
        Else
            tmrUpdate.Enabled = False
        End If
        picSJS.Visible = False
        btnPath.Enabled = False
    End Sub

    Private Sub Sync_Carpark_Index_Between_Setting_File_And_DB_Lot_Available()
        Dim I As Integer
        Dim sql As String
        Dim st As String
        Dim st2 As String = ""
        Dim err1 As String = ""
        Dim n As Integer = 0
        Dim count1 As Integer = 0

        Try
            For I = 0 To TotalCarpark - 1
                st = Carparks(I).stName             ' this is carpark name
                'st2 = Carparks(I).stID              ' this is ticketSiteID from Parameter_mst
                sql = "UPDATE Carpark_Lots_Availablity SET [Carpark_Index] = '" & I + 1 & "' "
                sql = sql & "WHERE [Carpark_Number] = '" & st & "'"
                err1 = ExecuteNonQuery(sql)
                If err1 <> "Error" Then
                    If IsNumeric(err1) Then
                        n = Val(err1)
                        If n = 0 Then
                            Call WriteLog("(0) row(s) affected: This carpark " & st & " is not in the database")
                        Else
                            count1 += 1
                        End If
                    Else
                        Call WriteLog("Invalid return value: " & err1)
                    End If
                Else
                    Call WriteLog("Error sql: " & sql)
                End If
            Next
            Call WriteLog("Total Carpark(s): " & TotalCarpark & " and " & count1 & " Carpark(s) updated success")
        Catch ex As Exception
            Call WriteLog("Error: " & ex.Message)
        End Try

    End Sub

    Private Sub Show_Carpark()
        Dim I As Integer
        Dim J As Integer
        Dim K As Integer = 0
        ' index start from ZERO
        Dim gIndex As Integer = 0
        
        ReDim B(TotalCarpark - 1)           ' create instances for controls

        ' 0 for Green, 1 for Orange, 2 for Red, 3 for Black

        ' TotalCarpark > 0 and TotalCarpark <= 100
        ' If less than 100, the rest will not be created
        For I = 0 To TotalCarpark - 1
            B(I) = New ctlCarpark

            Me.Controls.Add(B(I))

            B(I).Height = gnCarparkHeight
            B(I).Width = gnCarparkWidth

            B(I).Relocate1()

            'B.Height = 30
            'B.Width = 40

            ' B(I).Left = LeftMargin + n + m     ' move left margin by n + m

            n = n + B(I).Width          ' added width to n
            m = m + H_Spacing           ' added space to m

            ' B(I).Top = TopMargin

            B(I).Left = Carparks(I).x_value - 20          ' offset value
            B(I).Top = Carparks(I).y_value - 40           ' offset value

            If I > 0 And (I + 1) Mod ColPerRow = 0 Then
                TopMargin = TopMargin + B(I).Height + V_Spacing   ' reset vertical movement
                n = 0
                m = 0
            End If

            B(I).text1 = Carparks(I).stName & " " & Carparks(I).lot_available.ToString            ' label1 for carpark name
            B(I).text2 = Carparks(I).stMsg '(I + 1).ToString & Carparks(I).stMsg   ' label2 for message
            B(I).IP1 = Carparks(I).stIP                 ' store ip address
            B(I).OldPrior1 = Carparks(I).OldPriority
            B(I).Prior1 = Carparks(I).nPriority         ' store priority level
            B(I).idx1 = I               ' array index also put in tag of label1
            B(I).idx2 = I               ' array index also put in tag of label2
            B(I).CarparkID1 = Carparks(I).stID
            B(I).Update_dt = Carparks(I).update_dt
            B(I).Label1.Text = B(I).text1
            B(I).Label1.Visible = False
            B(I).Group1 = Carparks(I).GroupID
            B(I).GroupIndex1 = Carparks(I).GroupIndex
            B(I).Operation1 = Carparks(I).Operation

            For J = 0 To TotalGroup + 2     '1   ' Change on 11-12-2017 for JTC Special version to show non-operation carparks with G instead of using Lable1
                With B(I).G(J)
                    '.R1 = Green1.R
                    '.G1 = Green1.G
                    '.B1 = Green1.B
                    .Left = 1
                    .Top = 1
                    .Width = B(I).Width - 2
                    .Height = B(I).Height - 2
                    .lblCarpark.Left = 0
                    .lblCarpark.Top = 0
                    .lblCarpark.Width = .Width
                    .lblCarpark.Height = .Height
                    .lblCarpark.BringToFront()
                    .lblCarpark.Text = B(I).text1    ' CarparkID
                    '.lblCarpark.BackColor = Color.Green   ' default is green
                    '.Visible = False
                End With

                AddHandler B(I).G(J).lblCarpark.DoubleClick, _
                AddressOf DoubleClickHandler2
            Next

            'If TotalGroup > 1 Then      '20141015
            If B(I).Operation1 = "OPERATION" Then
                Select Case B(I).GroupIndex1
                    Case 0 : Show_Carpark_Color("Carpark", I, "Green1")
                    Case 1 : Show_Carpark_Color("Carpark", I, "Green2")
                    Case 2 : Show_Carpark_Color("Carpark", I, "Green3")
                End Select
            Else
                Show_Carpark_Color("Carpark", I, "Black")
            End If

            'End If
            'B(I).btnCarpark.Text = B(I).text1
            'B(I).btnCarpark.BackColor = ColorTranslator.FromOle(RGB(Green1.R, Green1.G, Green1.B))
            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            ' just for test showing (later need to remove this section)
            If 1 = 2 Then
                Select Case I
                    Case 43, 50 'B(I).btnCarpark.BackColor = ColorTranslator.FromOle(RGB(Orange1.R, Orange1.G, Orange1.B))
                        ' set to Orange
                        Call Show_Carpark_Color("Carpark", I, "Orange")
                    Case 30, 61 'B(I).btnCarpark.BackColor = ColorTranslator.FromOle(RGB(Red1.R, Red1.G, Red1.B))
                        ' set to Red
                        Call Show_Carpark_Color("Carpark", I, "Red")
                        'Case 70, 71, 72, 73 'B(I).btnCarpark.BackColor = ColorTranslator.FromOle(RGB(Black1.R, Black1.G, Black1.B))
                        '    ' set to Black (Transparent)
                        '    Call Show_Carpark_Color("Carpark", I, "Black")
                End Select
                ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            End If

            B(I).Tag = I
            If gbHideCarparks Then B(I).Visible = False

            AddHandler B(I).Click, AddressOf ClickHandler          ' test for green color
            AddHandler B(I).Label1.Click, AddressOf ClickHandler2  ' test for red color
            AddHandler B(I).Label2.Click, AddressOf ClickHandler3  ' test for orange color
            'AddHandler B(I).DoubleClick, AddressOf DoubleClickHandler

        Next
    End Sub

    Private Sub Show_Carpark_Color(ByVal flag As String, ByVal n As Integer, ByVal st As String)
        ' n is index of B, the rest need to hide
        Dim I As Integer = 0
        Dim K As Integer = 0
        Dim J As Integer

        K = TotalGroup + 2 ' 1   change on 11-12-2017 to show non-operation carparks with G instead of using Label1

        If flag = "Carpark" Then
            For I = 0 To K
                B(n).G(I).Visible = False   ' Hide all firstly
            Next
            Select Case TotalGroup
                Case 1
                    Select Case st
                        Case "Green1" : B(n).G(0).Visible = True
                        Case "Orange" : B(n).G(1).Visible = True
                        Case "Red" : B(n).G(2).Visible = True
                            'Case "Black" : B(n).Label1.Visible = True
                        Case "Black" : B(n).G(K).Visible = True       ' change on 11-12-2017 for JTC Special Version
                            ' Label1 is used to show Carpark (only for NOT YET OPERATION)
                    End Select
                Case 2
                    Select Case st
                        Case "Green1" : B(n).G(0).Visible = True
                        Case "Green2" : B(n).G(1).Visible = True
                        Case "Orange" : B(n).G(2).Visible = True
                        Case "Red" : B(n).G(3).Visible = True
                            'Case "Black" : B(n).Label1.Visible = True
                        Case "Black" : B(n).G(K).Visible = True       ' change on 11-12-2017 for JTC Special Version
                            ' Label1 is used to show Carpark (only for NOT YET OPERATION)
                    End Select
                Case 3
                    Select Case st
                        Case "Green1" : B(n).G(0).Visible = True
                        Case "Green2" : B(n).G(1).Visible = True
                        Case "Green3" : B(n).G(2).Visible = True
                        Case "Orange" : B(n).G(3).Visible = True
                        Case "Red" : B(n).G(4).Visible = True
                            'Case "Black" : B(n).Label1.Visible = True
                        Case "Black" : B(n).G(K).Visible = True       ' change on 11-12-2017 for JTC Special Version
                            ' Label1 is used to show Carpark (only for NOT YET OPERATION)
                    End Select
            End Select
            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            ' BELOW is for displaying LEGEND ICON at the bottom of the screen
            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Else     ' legend handling are below
            For I = 0 To K
                Legend(n).G(I).Visible = False
            Next
            Select Case TotalGroup
                Case 1
                    Select Case st
                        Case "Green1"
                            J = 0
                            Legend(n).G(J).Visible = True
                            Legend(n).G(J).lblCarpark.Text = "OK"  ' (" & GroupID(0) & ")"    ' "OK" only is for special version JTC
                            Legend(n).text2 = st

                        Case "Orange"
                            J = 1
                            Legend(n).G(J).Visible = True
                            Legend(n).G(J).lblCarpark.Text = "Non Critical"   ' Problem"
                            Legend(n).text2 = st

                        Case "Red"
                            J = 2
                            Legend(n).G(J).Visible = True
                            Legend(n).G(J).lblCarpark.Text = "Critical"  ' Problem"
                            Legend(n).text2 = st

                        Case "Black"
                            Legend(n).Label1.Visible = False ' True   ' the value False is only for Special version JTC
                            Legend(n).G(K).Visible = True
                            Legend(n).G(K).lblCarpark.Text = "Not Yet"  ' 
                            Legend(n).Label1.Text = "Not yet Operation"
                            Legend(n).text2 = st
                    End Select
                Case 2
                    Select Case st

                        Case "Green1"
                            J = 0
                            Legend(n).G(J).Visible = True
                            Legend(n).G(J).lblCarpark.Text = "OK (" & GroupID(0) & ")"
                            Legend(n).text2 = st

                        Case "Green2"
                            J = 1
                            Legend(n).G(J).Visible = True
                            Legend(n).G(J).lblCarpark.Text = "OK (" & GroupID(1) & ")"
                            Legend(n).text2 = st

                        Case "Orange"
                            J = 2
                            Legend(n).G(J).Visible = True
                            Legend(n).G(J).lblCarpark.Text = "No Critical Problem"
                            Legend(n).text2 = st

                        Case "Red"
                            J = 3
                            Legend(n).G(J).Visible = True
                            Legend(n).G(J).lblCarpark.Text = "Critical Problem"
                            Legend(n).text2 = st

                        Case "Black"
                            Legend(n).Label1.Visible = False ' True
                            Legend(n).G(K).Visible = True
                            Legend(n).G(K).lblCarpark.Text = "Not Yet"
                            Legend(n).Label1.Text = "Not yet Operation"
                            Legend(n).text2 = st
                    End Select
                Case 3
                    Select Case st
                        Case "Green1"
                            J = 0
                            Legend(n).G(J).Visible = True
                            Legend(n).G(J).lblCarpark.Text = "OK (" & GroupID(0) & ")"
                            Legend(n).text2 = st

                        Case "Green2"
                            J = 1
                            Legend(n).G(J).Visible = True
                            Legend(n).G(J).lblCarpark.Text = "OK (" & GroupID(1) & ")"
                            Legend(n).text2 = st

                        Case "Green3"
                            J = 2
                            Legend(n).G(J).Visible = True
                            Legend(n).G(J).lblCarpark.Text = "OK (" & GroupID(2) & ")"
                            Legend(n).text2 = st

                        Case "Orange"
                            J = 3
                            Legend(n).G(J).Visible = True
                            Legend(n).G(J).lblCarpark.Text = "No Critical Problem"
                            Legend(n).text2 = st

                        Case "Red"
                            J = 4
                            Legend(n).G(J).Visible = True
                            Legend(n).G(J).lblCarpark.Text = "Critical Problem"
                            Legend(n).text2 = st

                        Case "Black"
                            Legend(n).Label1.Visible = False ' True
                            Legend(n).G(K).Visible = True
                            Legend(n).G(K).lblCarpark.Text = "Not Yet"
                            Legend(n).Label1.Text = "Not yet Operation"
                            Legend(n).text2 = "Black"
                    End Select

            End Select

        End If
    End Sub

    Private Sub Show_Blinking()
        Dim I As Integer = 0
        Dim K As Integer = 0

        StillBlinking = True          ' init flag (boolean)
        BlinkCounter = 0              ' set counter value to zero
        SoundAlarmCounter = 0

        K = TotalLegend + TotalGroup
        For I = 0 To K - 2
            If Legend(I).text2 = "Red" Then
                IndexofRed = I                 ' find the exact location of red color button (Legend)
                Exit For               ' because its index may change based on the number of group (1,2,3)
            End If
        Next
        tmrBlink.Enabled = True         ' start to let blink timer
        'MsgBox(tmrBlink.Enabled)
    End Sub

    Private Sub Show_Legend()
        Dim I As Integer
        Dim J As Integer
        Dim K As Integer
        Dim fontName As FontFamily = lblDummyFont.Font.FontFamily '// Get the Current Font Name used.
        'TextBox1.Font = New Font(fontName, 10) 'TrackBar1.Value)

        K = TotalLegend + TotalGroup
        ' eg.  4 + 2 = 6 - 1 = 5;  green1, gree2, orange, red, black

        ReDim Legend(K - 2)         ' to show legend
        n = 0
        m = 0
        LeftMargin = lblColorCode.Left + lblColorCode.Width + 10
        TopMargin = lblColorCode.Top - 5

        For I = 0 To K - 2  'TotalCarpark - 1
            Legend(I) = New ctlCarpark
            Me.Controls.Add(Legend(I))

            Legend(I).Height = gnCarparkHeight + 15       ' the same control is used for both Carpark and Legend, but Legend have bigger size
            Legend(I).Width = gnCarparkWidth + 35

            Legend(I).Relocate1()

            'B.Height = 30
            'B.Width = 40

            Legend(I).Left = LeftMargin + n + m     ' move left margin by n + m
            n = n + Legend(I).Width          ' added width to n
            m = m + H_Spacing           ' added space to m

            Legend(I).Top = TopMargin

            If I > 0 And (I + 1) Mod ColPerRow = 0 Then
                TopMargin = TopMargin + Legend(I).Height + V_Spacing   ' reset vertical movement
                n = 0
                m = 0
            End If

            Legend(I).text1 = "A" 'Carparks(I).stName            ' label1 for carpark name
            Legend(I).text2 = "B" 'Carparks(I).stMsg '(I + 1).ToString & Carparks(I).stMsg   ' label2 for message
            Legend(I).IP1 = "C" 'Carparks(I).stIP                 ' store ip address
            Legend(I).Prior1 = 0 'Carparks(I).nPriority         ' store priority level
            Legend(I).idx1 = I               ' array index also put in tag of label1
            Legend(I).idx2 = I               ' array index also put in tag of label2
            Legend(I).CarparkID1 = "D" 'Carparks(I).stID
            Legend(I).Update_dt = Now ' Carparks(I).update_dt
            Legend(I).Label1.Text = Legend(I).text1
            Legend(I).Label1.Visible = False
            Legend(I).Label1.Font = New Font(fontName, 12)
            Legend(I).Group1 = "No Need"
            Legend(I).GroupIndex1 = 1
            Legend(I).Operation1 = "No Need"

            For J = 0 To TotalGroup + 2 ' 1     ' change on 11-12-2017 for the JTC Special version to show non-operation carparks instead of using Lable1
                With Legend(I).G(J)
                    '.R1 = Green1.R
                    '.G1 = Green1.G
                    '.B1 = Green1.B
                    .Left = 1
                    .Top = 1
                    .Width = Legend(I).Width - 2
                    .Height = Legend(I).Height - 2
                    .lblCarpark.Left = 0
                    .lblCarpark.Top = 0
                    .lblCarpark.Width = .Width
                    .lblCarpark.Height = .Height
                    .lblCarpark.BringToFront()
                    .lblCarpark.Text = Legend(I).text1
                    .lblCarpark.Font = New Font(fontName, 12)
                End With
            Next

            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            Select Case TotalGroup
                Case 1
                    Select Case I
                        Case 0 : Show_Carpark_Color("Legend", I, "Green1")
                        Case 1 : Show_Carpark_Color("Legend", I, "Orange")
                        Case 2 : Show_Carpark_Color("Legend", I, "Red")
                        Case 3 : Show_Carpark_Color("Legend", I, "Black")
                    End Select
                Case 2
                    Select Case I
                        Case 0 : Show_Carpark_Color("Legend", I, "Green1")
                        Case 1 : Show_Carpark_Color("Legend", I, "Green2")
                        Case 2 : Show_Carpark_Color("Legend", I, "Orange")
                        Case 3 : Show_Carpark_Color("Legend", I, "Red")
                        Case 4 : Show_Carpark_Color("Legend", I, "Black")
                    End Select
                Case 3
                    Select Case I
                        Case 0 : Show_Carpark_Color("Legend", I, "Green1")
                        Case 1 : Show_Carpark_Color("Legend", I, "Green2")
                        Case 2 : Show_Carpark_Color("Legend", I, "Green3")
                        Case 3 : Show_Carpark_Color("Legend", I, "Orange")
                        Case 4 : Show_Carpark_Color("Legend", I, "Red")
                        Case 5 : Show_Carpark_Color("Legend", I, "Black")
                    End Select
            End Select

            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            Legend(I).Tag = I
            If gbHideCarparks Then Legend(I).Visible = False

            AddHandler Legend(I).Click, AddressOf ClickHandler          ' test for green color
            AddHandler Legend(I).Label1.Click, AddressOf ClickHandler2  ' test for red color
            AddHandler Legend(I).Label2.Click, AddressOf ClickHandler3  ' test for orange color

        Next

    End Sub

    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    ' xxxxxxxxxxxxxxxxxxxx EVENT HANDLERS xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    Public Sub ClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)

        'MsgBox("I am Carpark #" & CType(sender, ctlCarpark).Text)
        'ArrayIndex = Array.IndexOf(B, sender)
        'MsgBox(ArrayIndex)
        'just_test_only(ArrayIndex, "OK", 1)  ' to test GREEN COLOR

    End Sub

    Public Sub DoubleClickHandler2(ByVal sender As Object, ByVal e As System.EventArgs)

        'MsgBox("I am Carpark #" & CType(sender, ctlCarpark).Text)
        'ArrayIndex = Array.IndexOf(B, sender)
        Dim st As String = ""
        st = CType(sender, Label).Text
        'MsgBox("Double Click !", MsgBoxStyle.Information, ArrayIndex)
        Dim idx As Integer = Node_Index(st)
        ' Don't have Checked Value because those control are not Checked Boxes

        'MsgBox(st, MsgBoxStyle.Information, idx)

        Call Assign_Node(st, idx)
        'Call VNC_Connection(st)
        'MsgBox(ArrayIndex)
        'just_test_only(ArrayIndex, "OK", 1)  ' to test GREEN COLOR
    End Sub

    Private Sub Assign_Node(ByVal node1 As String, ByVal index1 As Integer)
        Dim I As Integer = 0
        Dim J As Integer = 0
        Dim K As Integer = 0
        Dim dummy1 As Integer = 123

        If order1 >= 2 Then
            MsgBox("Sorry you already selected 2 nodes for this path")
            Exit Sub
        End If
        Select Case order1

            Case 0
                order1 = order1 + 1
                Selected_Node(index1) = order1
                Color_Change(index1, 2)   ' 2 = Red Color

            Case 1
                order1 = order1 + 1
                Selected_Node(index1) = order1
                Color_Change(index1, 2)
            Case Else
        End Select
        If order1 = 2 Then
            change_Orange_The_Rest()
            btnPath.Enabled = True
        End If
        dummy1 = 124
    End Sub

    Private Sub change_Orange_The_Rest()
        Dim I As Integer = 0
        For I = 0 To 4
            If Selected_Node(I) = 0 Then
                Color_Change(I, 1)    ' 1 = Orange, 0 = Green
            End If
        Next
    End Sub

    Private Sub Color_Change(ByVal idx1 As Integer, ByVal colorIdx As Integer)
        Dim I As Integer = 0
        Dim J As Integer = 0
        For I = 0 To 4
            If I = idx1 Then
                For J = 0 To 2
                    B(I).G(J).Visible = False
                Next
            Else

            End If
            
        Next
        B(idx1).G(colorIdx).Visible = True
    End Sub

    Dim Node_Label() As String = {"A", "B", "C", "D", "E"}
    Dim Selected_Node() As Integer = {0, 0, 0, 0, 0}
    Dim order1 As Integer = 0

    Dim vertices As Integer = 5
    Dim graph(,) As Integer = {{0, 2, 0, 6, 4}, _
                               {2, 0, 8, 0, 3}, _
                               {0, 8, 0, 5, 0}, _
                               {6, 0, 5, 0, 0}, _
                               {4, 3, 0, 0, 0}}

    Dim labels() As String = {"A", "B", "C", "D", "E"}

    Dim gnDestination As String = "A"
    Dim gnSource As String = "C"

    Dim inf1 As Integer = 9999
    Dim source_index As Integer ' String
    Dim dest_index As Integer

    Dim dist(4) As Integer
    Dim spt_set(4) As Integer
    Dim parent1(4) As Integer

    Private Function Node_Index(ByVal st As String) As Integer
        Node_Index = 0
        Dim temp As String = ""
        st = st.Trim
        Dim dummy1 As Integer = 0
        For i As Integer = 0 To 4
            temp = Node_Label(i).Trim

            If temp = st Then
                'Return i
                Node_Index = i
                Exit For
            End If
        Next
        dummy1 = 1
    End Function

    Private Sub VNC_Connection(ByVal cp As String)
        Dim stCPName As String
        Dim stCPID As String
        Dim stIPAddress As String
        Dim ans As Integer
        Dim I As Integer = 0
        Dim st As String = ""
        Dim stName1 As String = ""
        Dim cp1 As String = ""

        With Carparks
            cp1 = cp.Substring(0, cp.IndexOf(" "))
            For I = 0 To .Length - 1

                stName1 = Carparks(I).stName
                If cp1.ToUpper = stName1.ToUpper Then
                    global_CarparkID = I
                    Exit For
                End If
            Next
        End With
        stIPAddress = Carparks(global_CarparkID).stIP
        ' Calling Ultra VNC
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        ' Shell("C:\Program Files\uvnc bvba\UltraVNC\vncviewer.exe 192.168.116.133 /password suntest", AppWinStyle.NormalFocus)
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

        stCPName = Carparks(global_CarparkID).stName
        stCPID = Carparks(global_CarparkID).stID
        ans = vbYes   '= MsgBox("Connect to " & stCPName, MsgBoxStyle.YesNo, stCPID)
        ' do remote connection by using UltraVNC     *VNC*
        ' DEMO Setting: MNO PC for Carpark RCB and TMS PC
        'If stIPAddress = "192.168.2.35" Or stIPAddress = "192.168.5.6" Then
        'Else
        '    Exit Sub
        'End If
        If ans = vbYes Then
            Try
                'stIPAddress = "192.168.2.100"   ' hardcode for testing
                'st = "C:\Program Files\uvnc bvba\UltraVNC\vncviewer.exe "
                st = gstUltraVNC_Path & "\vncviewer.exe "
                st = st & stIPAddress & " /password sunsjs"
                ' MsgBox(st, MsgBoxStyle.Information, stCPID & ", " & stCPName)


                I = Shell(st, AppWinStyle.NormalFocus)
            Catch ex As Exception
                Call WriteLog("Routine: VNC_Connection, Error:" & ex.Message)
            End Try
        Else
            Exit Sub
        End If

        'Form1.Process_Incoming_Message("12,3")
    End Sub

    Public Sub ClickHandler2(ByVal sender As Object, ByVal e As System.EventArgs)
        'MsgBox("I am Label1 # Tag:" & CType(sender, Label).Tag)
        'just_test_only(CType(sender, Label).Tag, "Barrier drop", 3)   ' to test RED COLOR
    End Sub

    Public Sub ClickHandler3(ByVal sender As Object, ByVal e As System.EventArgs)
        'MsgBox("I am Label2 # Tag:" & CType(sender, Label).Tag)
        'just_test_only(CType(sender, Label).Tag, "Printer error", 2)    ' to test ORANGE COLOR
        global_CarparkID = CType(sender, Label).Tag
    End Sub

    Private Sub just_test_only(ByVal idx4 As Integer, ByVal msg As String, ByVal p As Integer)

        Carparks(idx4).nPriority = p
        Carparks(idx4).stMsg = msg

        Rearrange_Carparks()    ' call sorting procedure
    End Sub

    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    ' Below is very important function, need to handle carefully
    Private Function Get_Priority(ByVal st As String, ByVal cp1 As String, ByRef ShowMsg As String, _
                ByVal OldPriority As Integer, ByVal OldMsg As String) As Integer
        ' Add 2 new parameters, OldMsg and OldPriority is to handle INTERCOM TIMEOUT
        ' Printer Paper Low (ORANGE) must be overwrite Intercom (RED) before its timeout
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Dim p As Integer = 1
        Dim I As Integer
        Dim J As Integer = 0
        Dim maxvalue As Integer = 0
        Dim totalrow As Integer = 0
        Dim st1 As String = ""
        Dim p1 As Integer = 1
        Dim msg1 As String = ""
        Dim income_msg As String = ""

        If gbDetailPrint Then WriteLog("Function: Get_Priority Start ...............")
        If gbDetailPrint Then WriteLog("Start Msg = " & st)
        ShowMsg = st                  ' this msg is not so sure to be display when it is OK msg

        For I = 0 To gnTotalMessage - 1

            ' income msg may be more wordings than stored msg in this
            ' application, so cannot match exactly by using UCase and = Sign
            ' Just check stored msg is inside income_msg
            income_msg = UCase(st)                        ' income message
            msg1 = UCase(ArrayAlarmMessage(I).Message)    ' stored message

            If msg1 = income_msg OrElse income_msg.Contains(msg1) Then
                'If msg1 = income_msg AndAlso income_msg.Contains(msg1) Then
                p = ArrayAlarmMessage(I).Priority
                Get_Priority = p
                'Exit Function
                'Exit For
            End If
        Next

        If gbDetailPrint Then WriteLog("Initial Priority = " & p)

        ' Note: more logic are as per below section
        ' If you don't want (to avoid) it, just go outside here
        ' Just show latest priority logic applied here ---------------------------
        If 1 = 1 Then Return p ' don't want to execute below section (messy logic)
        ' below logic is no longer used and replaced it by DECISION table feature
        ' ------------------------------------------------------------------------
        ' Need to process further as below 
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Dim dr() As DataRow         ' Array of DataRow

        ' Further enhancement are as below
        ' Above process is only for just incoming message;
        ' It is very important to check all the previous status of all stations
        Select Case p
            Case 3           ' For 3, no further checking is needed
                Return 3         ' Higest priority; RED Color, just return the value 
                '(no need to check the history)
                If gbDetailPrint Then WriteLog("Return Priority = 3")
            Case 2, 1                 ' If Orange or Green color, check dt array again to make sure
                ' set lastest priority p to be maxvalue
                maxvalue = p          ' Scan the dt array and find the more important error message
                If gbDetailPrint Then WriteLog("Start maxvalue = " & maxvalue)
                With dt          ' Global memory table
                    totalrow = .Rows.Count
                    If totalrow = 0 Then
                        ' nothing to do, just return 
                        Return p
                    End If

                    ' Select the rows only for specific CarparkID
                    dr = .Select("CarparkID = '" & cp1 & "'")
                    For Each row1 As DataRow In dr
                        st1 = row1(1)           ' Read message 
                        ShowMsg = st1
                        If gbDetailPrint Then WriteLog("Proposed ShowMsg = " & ShowMsg)
                        For I = 0 To gnTotalMessage - 1           ' and check its priority
                            If st1 = ArrayAlarmMessage(I).Message Then
                                p1 = ArrayAlarmMessage(I).Priority
                                Exit For
                            End If
                        Next
                        If p1 > maxvalue Then
                            maxvalue = p1
                            If maxvalue = 3 Then
                                If gbDetailPrint Then WriteLog("ShowMsg = " & ShowMsg & ", Priority = 3")
                                Return 3         ' if found highest priority, no need to check further

                            End If
                        End If
                    Next
                    ' scan dt to find existing error msg
                End With
        End Select
    End Function

    Public Sub Process_Incoming_Message(ByVal msg As String)
        Dim st() As String
        Dim I As Integer = 0
        Dim J As Integer = 0
        Dim K1 As Integer = 0
        Dim K2 As Integer = 0
        Dim CPID As String = ""
        Dim CarparkName As String = ""
        Dim Priority As Integer
        Dim FinalPriority As Integer
        Dim Message As String = ""
        Dim stStationName As String = ""        ' stStation is added on 19-9-2013
        Dim nStationID As Integer = 0
        Dim stCommand As String = ""
        Dim MsgType As String = ""
        Dim status1() As String             ' msg is splited by | and store in array
        Dim count1 As Integer = 0
        Dim ShowMsg1 As String = ""
        Dim stLog As String = ""

        ' Below is sample ALARM Message (Original and Processed message)
        '25-10-2013 12:44:44 : (5) [761|Exit|3|00|Stopped|]
        '25-10-2013 12:44:44 : (5) 761|Stopped|Exit|3|00|ALARM
        '25-10-2013 12:44:44 : (1) 761|Stopped|Exit|3|00|ALARM
        ' Below is sample STATUS Message  (Original and Processed message)
        '25-10-2013 12:44:46 : (5) [761|BBM7_X|2|00|0,0,0,0,0,0,0,0,0,0,0,1,1;0,20131023175703,|]
        '25-10-2013 12:44:46 : (5) 761|Open|BBM7_X|2|00|0,0,0,0,0,0,0,0,0,0,0,1,1
        '25-10-2013 12:44:46 : (1) 761|Open|BBM7_X|2|00|0,0,0,0,0,0,0,0,0,0,0,1,1
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Try
            Try
                If gbDebugPrint Then Call WriteLog("(1) " & msg) ' before it is log ready
                'st = Split(msg, ",")             ' Message Format =>  carpark_id, priority

                st = Split(msg, "|")             ' Message Format =>  carpark_id, priority, message_type
                ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                ' Incoming message having 6 Info for each station of each carpark 
                CPID = st(0)            ' 230, 456, 876, etc. .........
                CarparkName = Get_CarparkName(CPID)

                Message = st(1)         ' Antena Error, Database Error, Reader Error, etc. .........
                stStationName = st(2)   ' Exit 1, Entry 1, etc. ...........
                nStationID = Val(st(3)) ' the value must be integer
                stCommand = st(4)
                MsgType = st(5)         ' = ALARM is alarm (or) <> ALARM is STATUS (0,0,0,0,0,0,0,0,0,0,0,0)
                ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                'table.Rows.Add(25, "Indocin", "David", DateTime.Now)

                ' Before saving; need to check the same CarparkID + StationID are already exists or not
                ' If exist, do UPDATE that row only, otherwise INSERT new row
                ' ====================================================================
                'Call Save_To_Datatable(CPID, Message, stStationName, _
                '                nStationID, stCommand, MsgType, CarparkName)

                'If gbDebugPrint Then Call WriteLog("(2) CPID: " & CPID & ", Priority: " & Priority)
                ' Find the correct array index where carpark_id = input carpark_id
                If UCase(Message) = "PMS OK" Then
                    stLog = "Process_Incoming_Message: Exit Sub, CpID: " & CPID
                    stLog = stLog & "(" & CarparkName & "), Msg: " & Message & ", StnName: " & stStationName
                    stLog = stLog & ", StnID: " & nStationID & ", Cmd: " & stCommand
                    stLog = stLog & ", Type: " & MsgType
                    'Save_To_Connection_Table
                    Call Save_To_Connection_Table(CPID, 0, CarparkName)   ' PMS PC OK - Update_dt
                    Call WriteLog(stLog)
                    Exit Sub
                End If
            Catch ex As Exception
                Call WriteLog("Process Incoming Message")
                Call WriteLog("Error #1: " & ex.Message)
            End Try

            For I = 0 To TotalCarpark - 1
                '111:
                'If gbDetailPrint Then Call WriteLog("(3) " & I.ToString & " : " & CPID & " =? " & Priority)
                If Carparks(I).stID = CPID Then             ' Find the corret carparkID in ARRAY
                    If gbDetailPrint Then
                        Call WriteLog("Match CarparkID: Array = " & Carparks(I).stID & " and Income = " & CPID)
                    End If
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    If MsgType <> "ALARM" Then     ' STATUS MESSAGE handling
                        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

                        'MessageBox.Show("test")
                        ' Last 2 parameters is to help INTERCOM msg handling
                        Try
                            ' Save to Decision Table
                            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                            Priority = Get_Priority(Message, CPID, ShowMsg1, _
                                    Carparks(I).nPriority, Carparks(I).stMsg)    ' get related priority level based on message
                            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                            If gbDetailPrint Then
                                stLog = "Initial Priority = " & Priority & ", "
                                stLog = stLog & Message & ", " & CPID
                                Call WriteLog("Process STATUS msg: " & stLog)
                            End If

                            Try
                                Call Save_To_Datatable(CPID, Message, stStationName, _
                                        nStationID, stCommand, MsgType, CarparkName, Priority)
                            Catch ex As Exception
                                Call WriteLog("Process Incoming Message")
                                Call WriteLog("Error #6: " & ex.Message)
                            End Try

                            Try
                                'Save_To_Connection_Table
                                Call Save_To_Connection_Table(CPID, nStationID, CarparkName)
                            Catch ex As Exception
                                Call WriteLog("Process Incoming Message")
                                Call WriteLog("Error #7: " & ex.Message)
                            End Try

                            Try
                                ' Decision table (need to print for reference)
                                Call Save_To_Decision_Table(CPID, Message, stStationName, _
                                        nStationID, stCommand, MsgType, CarparkName, Priority)
                            Catch ex As Exception
                                Call WriteLog("Process Incoming Message")
                                Call WriteLog("Error #8: " & ex.Message)
                            End Try

                            ' if new priority is less than existing priority
                            ' just keep existing STATE; 
                            ' EXCEPT old msg is INTERCOM;

                            ' Make Final Decision based on DECISION table --------------------
                            FinalPriority = Make_Final_Decision(CPID, nStationID)
                            If gbDebugPrint Then
                                Call WriteLog("STATUS: Final Priority = " & FinalPriority)
                            End If
                            Carparks(I).OldPriority = Carparks(I).nPriority  ' Keep old priority before overwrite
                            Carparks(I).nPriority = FinalPriority       ' Overwrit OLD msg by new Msg
                            Carparks(I).stMsg = ShowMsg1           ' Store Message Value
                            Carparks(I).update_dt = Now            ' Store receiving time
                        Catch ex As Exception
                            Call WriteLog("Process Incoming Message")
                            Call WriteLog("Error #5: " & ex.Message)
                        End Try

                        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                        ' Below is for VIEW - Form (also need to check dt array)
                        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                        If gbDebugPrint Then Call WriteLog("MsgType: " & MsgType)
                        status1 = Split(MsgType, ",")
                        K1 = LBound(status1)
                        K2 = UBound(status1)
                        If gbDetailPrint Then Call WriteLog("Status1: Lower Bound = " & K1 & _
                                " and Upper Bound = " & K2)

                        If Not (K1 = 0 And K2 = 12) Then '11) Then
                            Call WriteLog("Status array size error")
                            Exit Sub
                        End If
                        count1 = 0
                        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                        ' NOTE: Explanation for hardcoding below
                        ' The meaning of bit values are already translated to ALARM msg accordingly 
                        ' and 
                        ' corresponding priority is calculated ready
                        ' So all zero maintain to ZERO (this mean NO ERROR) and 
                        ' All non-zero is converted to 1 (this mean ERROR)
                        ' So setting the values to 0 (no error) and 1 (error) is not effected 
                        ' the operation
                        ' zero and non-zeros are handle 
                        status1(12) = 0        ' so this hardcode is OK; no worry
                        'MessageBox.Show(count1)
                        ' Note: 
                        Try
                            For J = 0 To 12 '11
                                Carparks(I).status(J) = status1(J)
                                ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                                ' zero and non-zeros are handle 
                                If J = 11 Then
                                    If status1(11) = 3 Or status1(11) = 2 Then
                                        status1(11) = "1"
                                    Else
                                        status1(11) = "0"
                                    End If
                                End If
                                ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                                If status1(J) = "0" Then count1 = count1 + 1
                            Next
                        Catch ex As Exception
                            Call WriteLog("Process Incoming Message")
                            Call WriteLog("Error #2: " & ex.Message)
                        End Try

                        ' MessageBox.Show(count1)
                        ' below block is commented because it may cause decision error
                        ' because it set prior 1 regardless of station id
                        'If count1 = 13 Then           '12 Then         ' clear ERROR
                        '    Carparks(I).nPriority = 1
                        '    Carparks(I).stMsg = "OK1"
                        'End If
                        ' stop commenting -----------------------------------------------------
                        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    Else
                        ' ALARM MESSAGE handling - 
                        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                        Try
                            Priority = Get_Priority(Message, CPID, ShowMsg1, _
                                        Carparks(I).nPriority, Carparks(I).stMsg)
                            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                            If gbDetailPrint Then
                                stLog = "Initial Priority = " & Priority & ", "
                                stLog = stLog & Message & ", " & CPID
                                Call WriteLog("Process ALARM msg: " & stLog)
                            End If

                            Call Save_To_Datatable(CPID, Message, stStationName, _
                                            nStationID, stCommand, MsgType, CarparkName, Priority)
                            'Save_To_Connection_Table
                            Call Save_To_Connection_Table(CPID, nStationID, CarparkName)

                            ' Decision table (need to print for reference)
                            Call Save_To_Decision_Table(CPID, Message, stStationName, _
                                    nStationID, stCommand, MsgType, CarparkName, Priority)

                            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                            ' Last 2 parameters is to help INTERCOM msg handling
                            ' get relative priority level based on message
                            ' if new priority is less than existing priority
                            ' just keep existing STATE; EXCEPT old msg is INTERCOM; 

                            ' Make Final Decision based on DECISION table ----------------------
                            FinalPriority = Make_Final_Decision(CPID, nStationID)
                            If gbDebugPrint Then
                                Call WriteLog("ALARM: Final Priority = " & FinalPriority)
                            End If

                            Carparks(I).OldPriority = Carparks(I).nPriority   ' Keep old prior before overwrite
                            Carparks(I).nPriority = FinalPriority    ' Overwrite old priority with new value
                            Carparks(I).stMsg = ShowMsg1       ' Overwrite old msg with new value
                            Carparks(I).update_dt = Now
                        Catch ex As Exception
                            Call WriteLog("Process Incoming Message")
                            Call WriteLog("Error #3: " & ex.Message)
                        End Try

                        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                        ' Below is for VIEW - Form
                        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                        Try
                            For J = 1 To 12   '11
                                If ArrayStatusMessage(J).Message = Message Then
                                    Carparks(I).status(J - 1) = "-9"
                                Else
                                    Carparks(I).status(J - 1) = "0"
                                End If
                            Next
                        Catch ex As Exception
                            Call WriteLog("Process Incoming Message")
                            Call WriteLog("Error #4: " & ex.Message)
                        End Try
                    End If
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    If gbDetailPrint Then Call WriteLog("Call Rearrange_Carparks")
                    Rearrange_Carparks() ' call sorting procedure and change color
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    If gbDetailPrint Then Call WriteLog("(4) CarparkID = " & CPID & "(" & CarparkName & ")")
                    Exit Sub
                End If
            Next
            Call WriteLog("CarparkID = " & CPID & " NOT found")
        Catch ex As Exception
            Call WriteLog("Process Incoming Message")
            Call WriteLog("Error: " & ex.Message)
        End Try
    End Sub

    ' Make Final Decision for Priority (1 = Green, 2 = Orange, 3 = Red)
    Private Function Make_Final_Decision(ByVal cp As String, ByVal stn As Integer) As Integer
        ' By using decision table dt1
        Dim dr2() As DataRow
        Dim count1 As Integer = 0
        Dim dr As DataRow
        Dim ans As Integer = 1
        Dim stMsg As String = "No decision record"

        If dt1.Rows.Count = 0 Then
            ans = 1             ' No record means no error at all
        Else
            dr2 = dt1.Select("CarparkID = '" & cp & "'", "Priority DESC")
            count1 = 0
            For Each dr In dr2
                count1 += 1
                If count1 >= 1 Then
                    stMsg = dr(1).ToString
                    ans = dr(8)           ' Take largest value for priority
                    Exit For
                End If
            Next
            If count1 = 0 Then
                ans = 1             ' No record means no error at all
            End If
        End If
        Make_Final_Decision = ans
        If gbDetailPrint Then
            Call WriteLog("Make Final Decision: CarparkID: " & cp & ", StationID: " & stn & _
                                ", Msg: " & stMsg & ", Final Priority = " & ans)
        End If
    End Function

    Private Function Get_CarparkName(ByVal cid As Integer) As String
        Dim I As Integer = 0
        Dim st As String = ""

        For I = 0 To gnTotalCarpark - 1
            If Carparks(I).stID = cid Then
                st = Carparks(I).stName
                Exit For
            End If
        Next
        Get_CarparkName = st
    End Function

    Private Sub Save_To_Decision_Table(ByVal cp_id As String, ByVal msg_value As String, _
                    ByVal stn_name As String, ByVal stn_id As Integer, _
                    ByVal cmd As String, ByVal msg_type As String, _
                    ByVal cpName As String, ByVal prior1 As Integer)

        Dim rcount As Integer = 0
        'Dim dr() As DataRow
        Dim count1 As Integer = 0
        Dim I As Integer = 0
        Dim sql As String = ""
        Dim deleted1 As Integer = 0
        Dim msg1 As String = ""

        ' cpName = Get_CarparkName(cp_id)
        ' all the new logic comes here
        Try
            If msg_type <> "ALARM" Then          ' Incoming STATUS message
                Select Case prior1
                    Case 1
                        ' Delete all rows with the same carpark and station
                        If gbDetailPrint Then
                            Call WriteLog("STATUS msg: Delete all rows with the same carpark and station")
                            Call WriteLog("Prior: 1, CarparkID: " & cp_id & " and StnID: " & stn_id)
                        End If
                        count1 = dt1.Rows.Count
                        If gbDetailPrint Then Call WriteLog("Before delete count = " & count1)

                        deleted1 = 0
                        Try
                            If count1 > 0 Then
                                For I = 0 To count1 - 1                        ' Use for loop with index
                                    If dt1.Rows(I).Item(0).ToString = cp_id And _
                                       dt1.Rows(I).Item(3) = stn_id Then
                                        dt1.Rows(I).Delete()
                                        deleted1 += 1
                                        If count1 = deleted1 Then
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                            If gbDetailPrint Then
                                Call WriteLog("Delete: SUCCESS, Total: " & count1 & ", Deleted = " & deleted1 & ", New: " & dt1.Rows.Count)
                            End If
                        Catch ex As Exception
                            If gbDetailPrint Then
                                Call WriteLog("Delete: FAIL, Total: " & count1 & ", Deleted = " & deleted1 & ", New: " & dt1.Rows.Count)
                                Call WriteLog("Error msg: " & ex.Message)
                            End If
                        End Try

                        'For Each dr As DataRow In dt1.Rows
                        '    If dr.Item(0).ToString = cp_id And _
                        '        dr.Item(3) = stn_id Then
                        '        dr.Delete()
                        '    End If
                        'Next

                    Case 2  ' Delete all rows with the same carpark and station, then add ORANGE message
                        If gbDetailPrint Then
                            Call WriteLog("STATUS msg: Delete all rows with the same carpark and station")
                            Call WriteLog("Prior: 2, CarparkID: " & cp_id & " and StnID: " & stn_id)
                        End If
                        deleted1 = 0
                        count1 = dt1.Rows.Count
                        If gbDetailPrint Then Call WriteLog("Before delete count = " & count1)
                        Try
                            If count1 > 0 Then
                                For I = 0 To count1 - 1
                                    If dt1.Rows(I).Item(0).ToString = cp_id And _
                                       dt1.Rows(I).Item(3) = stn_id Then 'And _
                                        'dt1.Rows(I).Item(8) = 2 Then
                                        dt1.Rows(I).Delete()
                                        deleted1 += 1
                                        If count1 = deleted1 Then
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                            If gbDetailPrint Then
                                Call WriteLog("Delete: SUCCESS, Total: " & count1 & ", Deleted = " & deleted1 & ", New: " & dt1.Rows.Count)
                            End If
                        Catch ex As Exception
                            If gbDetailPrint Then
                                Call WriteLog("Delete: FAIL, Total: " & count1 & ", Deleted = " & deleted1 & ", New: " & dt1.Rows.Count)
                                Call WriteLog("Error msg: " & ex.Message)
                            End If
                        End Try
                        If gbDetailPrint Then
                            Call WriteLog("Add new ORANGE record")
                            msg1 = "Cid: " & cp_id & ", Msg: " & msg_value & ", Stn: " & stn_name
                            msg1 = msg1 & ", StnID: " & stn_id & ", Cmd: " & cmd
                            msg1 = msg1 & ", Type: " & msg_type & ", CpName: " & cpName & ", Prior: " & prior1
                            Call WriteLog(msg1)
                        End If
                        dt1.Rows.Add(cp_id, msg_value, stn_name, stn_id, cmd, msg_type, _
                                                                    DateTime.Now, cpName, prior1)
                        If gbDetailPrint Then
                            Call WriteLog("SUCCESS: Add new ORANGE record, Total: " & dt1.Rows.Count)
                        End If
                    Case 3 ' Delete all rows with the same carpark and station, then add RED message
                        If gbDetailPrint Then
                            Call WriteLog("STATUS msg: Delete all rows with the same carpark and station")
                            Call WriteLog("Prior: 3, CarparkID: " & cp_id & " and StnID: " & stn_id)
                        End If
                        deleted1 = 0
                        count1 = dt1.Rows.Count
                        If gbDetailPrint Then Call WriteLog("Before delete count = " & count1)
                        Try
                            If count1 > 0 Then
                                For I = 0 To count1 - 1
                                    If dt1.Rows(I).Item(0).ToString = cp_id And _
                                       dt1.Rows(I).Item(3) = stn_id Then
                                        dt1.Rows(I).Delete()
                                        deleted1 += 1
                                        If count1 = deleted1 Then
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                            If gbDetailPrint Then
                                Call WriteLog("Delete: SUCCESS, Total: " & count1 & ", Deleted = " & deleted1 & ", New: " & dt1.Rows.Count)
                            End If
                        Catch ex As Exception
                            If gbDetailPrint Then
                                Call WriteLog("Delete: FAIL, Total: " & count1 & ", Deleted = " & deleted1 & ", New: " & dt1.Rows.Count)
                                Call WriteLog("Error msg: " & ex.Message)
                            End If
                        End Try
                        If gbDetailPrint Then
                            Call WriteLog("Add new RED record")
                            msg1 = "Cid: " & cp_id & ", Msg: " & msg_value & ", Stn: " & stn_name
                            msg1 = msg1 & ", StnID: " & stn_id & ", Cmd: " & cmd
                            msg1 = msg1 & ", Type: " & msg_type & ", CpName: " & cpName & ", Prior: " & prior1
                            Call WriteLog(msg1)
                        End If
                        dt1.Rows.Add(cp_id, msg_value, stn_name, stn_id, cmd, msg_type, _
                                        DateTime.Now, cpName, prior1)
                        If gbDetailPrint Then
                            Call WriteLog("SUCCESS: Add new RED record, Total: " & dt1.Rows.Count)
                        End If
                End Select
            Else
                Select Case prior1
                    Case 1
                        ' Ignore ALARM message GREEN color
                        If gbDetailPrint Then
                            Call WriteLog("ALARM: msg, Prior: 1, Action: Ignore ALARM message GREEN color")
                        End If
                    Case 2  ' Delete all rows with the same carpark and station, then add ORANGE message
                        If gbDetailPrint Then
                            Call WriteLog("ALARM msg: Delete all rows with the same carpark and station")
                            Call WriteLog("Prior: 2, CarparkID: " & cp_id & " and StnID: " & stn_id)
                        End If
                        deleted1 = 0
                        count1 = dt1.Rows.Count
                        If gbDetailPrint Then Call WriteLog("Before delete count = " & count1)
                        Try
                            If count1 > 0 Then
                                For I = 0 To count1 - 1
                                    If dt1.Rows(I).Item(0).ToString = cp_id And _
                                       dt1.Rows(I).Item(3) = stn_id Then ' And _
                                        'dt1.Rows(I).Item(8) = 2 Then
                                        dt1.Rows(I).Delete()
                                        deleted1 += 1
                                        If count1 = deleted1 Then
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                            If gbDetailPrint Then
                                Call WriteLog("Delete: SUCCESS, Total: " & count1 & ", Deleted = " & deleted1 & ", New: " & dt1.Rows.Count)
                            End If
                        Catch ex As Exception
                            If gbDetailPrint Then
                                Call WriteLog("Delete: FAIL, Total: " & count1 & ", Deleted = " & deleted1 & ", New: " & dt1.Rows.Count)
                                Call WriteLog("Error msg: " & ex.Message)
                            End If
                        End Try
                        If gbDetailPrint Then
                            Call WriteLog("Add new ORANGE record")
                            msg1 = "Cid: " & cp_id & ", Msg: " & msg_value & ", Stn: " & stn_name
                            msg1 = msg1 & ", StnID: " & stn_id & ", Cmd: " & cmd
                            msg1 = msg1 & ", Type: " & msg_type & ", CpName: " & cpName & ", Prior: " & prior1
                            Call WriteLog(msg1)
                        End If
                        dt1.Rows.Add(cp_id, msg_value, stn_name, stn_id, cmd, msg_type, _
                                        DateTime.Now, cpName, prior1)
                        If gbDetailPrint Then
                            Call WriteLog("SUCCESS: Add new ORANGE record, Total: " & dt1.Rows.Count)
                        End If
                    Case 3 ' Delete all rows with the same carpark and station, then add RED message
                        If gbDetailPrint Then
                            Call WriteLog("ALARM msg: Delete all rows with the same carpark and station")
                            Call WriteLog("Prior: 3, CarparkID: " & cp_id & " and StnID: " & stn_id)
                        End If
                        deleted1 = 0
                        count1 = dt1.Rows.Count
                        If gbDetailPrint Then Call WriteLog("Before delete count = " & count1)
                        Try
                            If count1 > 0 Then
                                For I = 0 To count1 - 1
                                    If dt1.Rows(I).Item(0).ToString = cp_id And _
                                       dt1.Rows(I).Item(3) = stn_id Then
                                        dt1.Rows(I).Delete()
                                        deleted1 += 1
                                        If count1 = deleted1 Then
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                            If gbDetailPrint Then
                                Call WriteLog("Delete: SUCCESS, Total: " & count1 & ", Deleted = " & deleted1 & ", New: " & dt1.Rows.Count)
                            End If
                        Catch ex As Exception
                            If gbDetailPrint Then
                                Call WriteLog("Delete: FAIL, Total: " & count1 & ", Deleted = " & deleted1 & ", New: " & dt1.Rows.Count)
                                Call WriteLog("Error msg: " & ex.Message)
                            End If
                        End Try
                        If gbDetailPrint Then
                            Call WriteLog("Add new RED record")
                            msg1 = "Cid: " & cp_id & ", Msg: " & msg_value & ", Stn: " & stn_name
                            msg1 = msg1 & ", StnID: " & stn_id & ", Cmd: " & cmd
                            msg1 = msg1 & ", Type: " & msg_type & ", CpName: " & cpName & ", Prior: " & prior1
                            Call WriteLog(msg1)
                        End If
                        dt1.Rows.Add(cp_id, msg_value, stn_name, stn_id, cmd, msg_type, _
                                        DateTime.Now, cpName, prior1)
                        If gbDetailPrint Then
                            Call WriteLog("SUCCESS: Add new RED record, Total: " & dt1.Rows.Count)
                        End If
                End Select
            End If
        Catch ex As Exception
            Call WriteLog("Error in Save to Decision Table: " & ex.Message)
        End Try

        Call write_dt_to_Decision_log() ' write the whole table to verify data
    End Sub

    Private Sub Save_To_Datatable(ByVal cp_id As String, ByVal msg_value As String, _
    ByVal stn_name As String, ByVal stn_id As Integer, _
    ByVal cmd As String, ByVal msg_type As String, ByVal cpName As String, ByVal prior1 As Integer)

        Dim rcount As Integer = 0
        Dim dr() As DataRow
        Dim count1 As Integer = 0
        Dim I As Integer = 0
        Dim sql As String = ""

        'cpName = Get_CarparkName(cp_id)

        With dt             ' dt is dec lared by globally
            rcount = .Rows.Count
            'If gbDebugPrint Then MsgBox("Total Record = " & rcount)
            If gbDebugPrint Then Call WriteLog("Total Record = " & rcount)

            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            ' for this table, carparkID + stationID = composite key
            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            If rcount = 0 Then                ' There is no record in table, so DO INSERT new record
                ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                '                     INSERT (8 columns)
                ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                dt.Rows.Add(cp_id, msg_value, stn_name, stn_id, cmd, msg_type, DateTime.Now, cpName, prior1)
                ' Insert new record to dt

                'If gbDebugPrint Then MsgBox("Add new record (the first record)")
                If gbDetailPrint Then WriteLog("Add new record (the first record)")

                'If gbDebugPrint Then MsgBox("Total (new) = " & .Rows.Count)
                If gbDebugPrint Then WriteLog("Total (new) = " & .Rows.Count)

            Else
                ' ================================================================
                'table1.Columns.Add("CarparkID", GetType(String))    ' (0)      ' Primary KEY
                'table1.Columns.Add("Message", GetType(String))      ' (1)
                'table1.Columns.Add("StationName", GetType(String))  ' (2)
                'table1.Columns.Add("StationID", GetType(Integer))   ' (3)      ' Primary KEY
                'table1.Columns.Add("Command", GetType(String))      ' (4)      ' Primary KEY
                'table1.Columns.Add("Type", GetType(String))         ' (5)
                'table1.Columns.Add("Update_dt", GetType(DateTime))  ' (6)
                'table1.Columns.Add("CarparkName", GetType(String))  ' (7)
                'table1.Columns.Add("Priority", GetType(Integer))    ' (8)      ' 
                ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

                ' below is WHERE clause, filter by  primary KEY

                ' All incoming messages 
                ' with different Carpark, Station, Command, Message
                ' are stored in memory for later use
                ' accordingly

                ' To avoid duplicate records in History table, use the following keys for searching
                sql = "CarparkID = '" & cp_id & "' "                    ' Different carpark
                sql = sql & " AND StationID = '" & stn_id & "' "        ' Different station
                sql = sql & " AND Command = '" & cmd & "' "             ' Different command
                sql = sql & " AND Message = '" & msg_value & "' "       ' Different Message
                ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                dr = dt.Select(sql)    ' select record

                count1 = 0
                For Each row As DataRow In dr
                    count1 = count1 + 1              ' find the number of rows in DataRow
                Next

                ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                '                          INSERT
                ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                If count1 = 0 Then        ' no existing record, DO INSERT one more record
                    'If gbDebugPrint Then MsgBox("Insert one more record again")
                    If gbDetailPrint Then
                        WriteLog("Insert one more record again")
                    End If

                    dt.Rows.Add(cp_id, msg_value, stn_name, stn_id, cmd, msg_type, DateTime.Now, cpName, prior1)
                    ' add new record to dt

                    'If gbDebugPrint Then MsgBox("Total (new) = " & .Rows.Count)
                    If gbDetailPrint Then
                        WriteLog("Total (new) = " & .Rows.Count)
                    End If

                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    '                       UPDATE
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                Else                    ' record already exists, DO UPDATE (How to update?)
                    'If gbDebugPrint Then MsgBox("Update existing record again")
                    If gbDetailPrint Then WriteLog("Update existing record again")

                    ' carparkid, stationid, stationname are no need to update
                    ' =======================================================
                    'If gbDetailPrint Then Call WriteLog("Old value = " & dr(0).Item(1) & " New value = " & msg_value)
                    'dr(0).Item(1) = msg_value         ' save new value for message
                    ' =======================================================
                    'If gbDetailPrint Then Call WriteLog("Old value = " & dr(0).Item(4) & ", New value = " & cmd)
                    'dr(0).Item(4) = cmd         ' save new value for command
                    ' =======================================================
                    If gbDetailPrint Then
                        Call WriteLog("Old value = " & dr(0).Item(5) & " New value = " & msg_type)
                    End If

                    dr(0).Item(5) = msg_type         ' save new value for msg_type
                    ' =======================================================
                    dr(0).Item(6) = DateTime.Now
                    'If gbDebugPrint Then MsgBox("Total (no change) = " & .Rows.Count)
                    If gbDetailPrint Then WriteLog("Total (no change) = " & .Rows.Count)
                End If
            End If
        End With
        If gbDetailPrint Then Call write_dt_to_log() ' write the whole table to verify data
    End Sub

    Private Sub Save_To_Connection_Table(ByVal cp_id As String, ByVal stn_id As Integer, ByVal cpname As String)
        Dim rcount As Integer = 0
        Dim dr() As DataRow
        Dim count1 As Integer = 0
        Dim I As Integer = 0
        Dim sql As String = ""
        Dim cp1 As String = ""
        Dim stn1 As String = ""
        Dim row1 As DataRow

        With dtConn             ' dt is declared by globally
            rcount = .Rows.Count
            If gbDebugPrint Then Call WriteLog("Total Record = " & rcount)
            If rcount = 0 Then                ' There is no record in table, so DO INSERT new record
                Try
                    .Rows.Add(cp_id, stn_id, DateTime.Now, cpname)
                Catch ex As Exception
                    Call WriteLog("Save_To_Connection_Table")
                    Call WriteLog("Error #11: " & ex.Message)
                End Try
            Else
                Try
                    Try
                        count1 = 0
                        Try
                            'Call WriteLog("cp_id = " & cp_id & ", StationID = " & stn_id)

                            sql = "CarparkID = '" & cp_id & "' "  ' Different carpark
                            sql = sql & " AND StationID = '" & stn_id & "'"     ' Different station
                            ' All integer value must be used within single quite
                            ' to avoid runtime error
                            dr = .Select(sql)    ' select record
                        Catch ex As Exception
                            Call WriteLog("Save_To_Connection_Table")
                            Call WriteLog("Error #16: " & ex.Message)
                        End Try

                        Try
                            For Each row1 In dr
                                count1 = count1 + 1              ' find the number of rows in DataRow
                            Next
                        Catch ex As Exception
                            Call WriteLog("Save_To_Connection_Table")
                            Call WriteLog("Error #15: " & ex.Message)
                        End Try
                    Catch ex As Exception
                        Call WriteLog("Save_To_Connection_Table")
                        Call WriteLog("Error #14: " & ex.Message)
                    End Try

                    If count1 = 0 Then        ' no existing record, DO INSERT one more record
                        Try
                            .Rows.Add(cp_id, stn_id, DateTime.Now, cpname)
                        Catch ex As Exception
                            Call WriteLog("Save_To_Connection_Table")
                            Call WriteLog("Error #13: " & ex.Message)
                        End Try
                    Else                    ' record already exists, DO UPDATE (How to update?)
                        Try
                            dr(0).Item(2) = DateTime.Now
                        Catch ex As Exception
                            Call WriteLog("Save_To_Connection_Table")
                            Call WriteLog("Error #12: " & ex.Message)
                        End Try
                    End If
                Catch ex As Exception
                    Call WriteLog("Save_To_Connection_Table")
                    Call WriteLog("Error #10: " & ex.Message)
                End Try

            End If
            ' For every station message, also need to update PMS time
            Try
                For Each row1 In .Rows
                    If row1.Item(0) = cp_id And row1.Item(1) = 0 Then
                        row1.Item(2) = Now
                    End If
                Next
            Catch ex As Exception
                Call WriteLog("Save_To_Connection_Table")
                Call WriteLog("Error #9: " & ex.Message)
            End Try

        End With

    End Sub

    ' below subroutine is just for testing only 
    Private Sub write_dtConn_to_Listbox()
        Dim I As Integer = 0
        Dim J As Integer = 0
        Dim dr As DataRow
        Dim st As String = ""
        Dim dr2() As DataRow    ' DataRow ARRAY
        Dim n As Integer = 0
        Dim st1() As String

        Try
            With dtConn
                ' Sorting by CarparkID, StationID, Update_dt
                dr2 = .Select("", "CarparkID ASC, StationID ASC")
                lstConnection.Items.Clear()

                'I = 0
                For Each dr In dr2

                    'st = "|"
                    'For Each item As Object In dr.ItemArray
                    '    st = st & item.ToString & "|"
                    'Next

                    n = DateDiff(DateInterval.Minute, dr.Item(2), Now)
                    'Call WriteLog((I + 1).ToString & " => " & st)
                    'st = st & n.ToString & "|"      ' Calculate Duraction

                    'st2 = (I + 1).ToString & " : " & st
                    If n > gnWaitingMinute Then
                        st1 = Split(dr.Item(3), " ")    ' split carpark name and lot available (Special version for JTC)

                        If dr.Item(1) = 0 Then
                            'st = dr.Item(3) & " PMS connection lost for "
                            st = st1(0) & " PMS connection lost for "
                            st = st & n.ToString & " minutes"
                        Else
                            'st = dr.Item(3) & " Station " & dr.Item(1) & " connection lost for "
                            st = st1(0) & " Station " & dr.Item(1) & " connection lost for "
                            st = st & n.ToString & " minutes"
                        End If

                        lstConnection.Items.Add(st)
                    End If
                    'I += 1
                Next

            End With
        Catch ex As Exception
            Call WriteLog("Error handling in Sub: write_dtConn_to_Listbox")
            Call WriteLog("Error message = " & ex.Message)
        End Try
    End Sub

    ' below subroutine is just for testing only 
    Private Sub write_dt_to_log()
        Dim I As Integer = 0
        Dim J As Integer = 0
        Dim dr As DataRow
        Dim st As String = ""
        Dim dr2() As DataRow    ' DataRow ARRAY

        'Call WriteLog("==============================================================")
        'Call WriteLog("The content of dt - Message History for each station per carpark")
        'Call WriteLog("==============================================================")

        With dt
            ' Sorting by CarparkID, StationID, Update_dt
            dr2 = .Select("", "CarparkID ASC, StationID ASC, Update_dt DESC")

            'Call WriteLog("Total History record = " & .Rows.Count)
            'Call WriteLog("============================================================")
            'For I = 0 To .Rows.Count - 1
            '    dr = dt.Rows(I)              ' take each row
            '    st = "|"
            '    For Each item As Object In dr.ItemArray
            '        st = st & item.ToString & "|"
            '    Next
            '    Call WriteLog((I + 1).ToString & " => " & st)
            'Next
            I = 0
            For Each dr In dr2
                st = "|"
                For Each item As Object In dr.ItemArray
                    st = st & item.ToString & "|"
                Next
                'Call WriteLog((I + 1).ToString & " => " & st)
                I += 1
            Next
        End With
        'Call WriteLog("================== Printing Message History End / Intercom Start ===========")

        
        ' Below is only message for INTERCOM in order of latest one on top
        ' -----------------------------------------------------------------
        Dim dr1() As DataRow         ' DataRow ARRAY
        Dim K As Integer = 1
        'Call WriteLog("================== Printing Intercom History ===========") and Fireengine History
        'dr1 = dt.Select("Message = 'Intercom'", "Update_dt DESC")
        'dr1 = dt.Select("Message = 'Intercom' or Message = 'Fireengine'", "Update_dt DESC")
        dr1 = dt.Select("Message = 'Intercom' or Message = 'Fireengine' or Message = 'Season Sync Error'", "Update_dt DESC")

        lstIntercom.Items.Clear()
        For Each dr In dr1
            st = ""
            'For Each item As Object In dr.ItemArray
            '    st = st & item.ToString & "|"
            'Next
            st = st & dr(6).ToString & vbTab               ' msg received time
            st = st & dr(7).ToString & vbTab               ' cpname
            'st = st & dr(2).ToString & vbTab & "Intercom"             ' station
            st = st & dr(2).ToString & vbTab & dr(1).ToString            ' station
            'Call WriteLog(K.ToString & " : " & st)
            lstIntercom.Items.Add(st)
            K += 1
        Next
        'Call WriteLog("================== Printing Intercom History Finish ===========")
    End Sub

    ' below subroutine is just for testing only 
    Private Sub write_dt_to_Decision_log()
        Dim I As Integer = 0
        Dim J As Integer = 0
        Dim dr As DataRow
        Dim st As String = ""
        Dim dr2() As DataRow    ' DataRow ARRAY

        Try
            If gbDetailPrint Then
                'Call WriteLog("**********************************************************")
                'Call WriteLog("The content of dt1 - Decision for each station per carpark")
                'Call WriteLog("***********************************************************")

                With dt1
                    ' Sorting by CarparkID, StationID, Update_dt
                    dr2 = .Select("", "CarparkID ASC, StationID ASC, Update_dt ASC")

                    'Call WriteLog("Total Decision record(s) = " & .Rows.Count)
                    'Call WriteLog("**********************************************************")

                    I = 0
                    For Each dr In dr2
                        st = "|"
                        For Each item As Object In dr.ItemArray
                            st = st & item.ToString & "|"
                        Next
                        'Call WriteLog((I + 1).ToString & " => " & st)
                        I += 1
                    Next
                End With
                'Call WriteLog("**************** Printing Decision Table Finish ****************")
            End If
        Catch ex As Exception
            Call WriteLog("Printing Decision Table: " & ex.Message)
        End Try
    End Sub

    Private Sub Rearrange_Carparks()
        Dim I As Integer = 0
        Dim st As String = ""
        Dim J As Integer = 0
        Dim K As Integer = 0
        Dim anyChanges As Integer = 0

        ' Incoming message are come from UDP port
        ' upon request basis
        'If Sorting_Carparks() Or 1 Then    ' step 1: sort carparks array based on priority level
        If 1 Then    ' step 1: sort carparks array based on priority level
            ' if there is any changes, need to refresh screen display to be the same as Carparks array
            ' step 2: synchronize with B array to show on screen

            gnGreenCount = 0   ' Total of Green Color carpark
            For I = 0 To TotalCarpark - 1               ' Index start from zero
                B(I).text1 = Carparks(I).stName            ' label1 for carpark name
                B(I).text2 = Carparks(I).stMsg '(I + 1).ToString & Carparks(I).stMsg   ' label2 for message
                B(I).IP1 = Carparks(I).stIP                          ' store ip address

                B(I).OldPrior1 = Carparks(I).OldPriority          ' Old priority to compare color changes
                B(I).Prior1 = Carparks(I).nPriority                  ' store priority level

                B(I).CarparkID1 = Carparks(I).stID        ' Note: Carpark ID is not array Index
                B(I).Update_dt = Carparks(I).update_dt

                If B(I).Operation1 <> "OPERATION" Then
                    Call Show_Carpark_Color("Carpark", I, "Black")
                    gnGreenCount += 1          ' Consider as GREEN to be tally by TotalCarpark
                Else
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    ' Check color changes (from which color to which color?) and write log
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    Try
                        If B(I).OldPrior1 = B(I).Prior1 Then
                            ' No color changes
                        Else
                            J = B(I).OldPrior1        ' from Old Priority
                            K = B(I).Prior1           ' to New Priority
                            st = ""
                            st = st & B(I).text1 & "("
                            st = st & B(I).CarparkID1 & ") change from " & ColorCode(J - 1) 'J.ToString
                            st = st & " to " & ColorCode(K - 1)   'K.ToString
                            st = st & ", Msg: " & B(I).text2
                            anyChanges += 1
                            Call WriteLog(st)
                        End If
                    Catch ex As Exception
                        Call WriteLog("Error in Check color changes: " & ex.Message)
                    End Try

                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    Select Case B(I).Prior1
                        Case 1
                            'B(I).image1 = My.Resources.Resource1.green
                            Select Case B(I).GroupIndex1
                                Case 0 : Call Show_Carpark_Color("Carpark", I, "Green1")
                                Case 1 : Call Show_Carpark_Color("Carpark", I, "Green2")
                                Case 2 : Call Show_Carpark_Color("Carpark", I, "Green3")
                            End Select
                            ' Check all Carpark (Operation) are GREEN, if YES, stop blinking
                            gnGreenCount += 1   ' Increase GREEN Count

                        Case 2
                            'B(I).image1 = My.Resources.Resource1.orange
                            Call Show_Carpark_Color("Carpark", I, "Orange")

                            gnGreenCount += 1   ' Increase GREEN Count
                        Case 3
                            'B(I).image1 = My.Resources.Resource1.red
                            Call Show_Carpark_Color("Carpark", I, "Red")
                            detectRed = True        ' this Flag is always monitering by Timer2
                            ' if TRUE, Timer2 invoke tmrBlink accordingly
                        Case Else
                            'B(I).image1 = My.Resources.Resource1.green
                            Select Case B(I).GroupIndex1
                                Case 0 : Call Show_Carpark_Color("Carpark", I, "Green1")
                                Case 1 : Call Show_Carpark_Color("Carpark", I, "Green2")
                                Case 2 : Call Show_Carpark_Color("Carpark", I, "Green3")
                            End Select
                            ' Check all Carpark (Operation) are GREEN, if YES, stop blinking
                            gnGreenCount += 1   ' Increase GREEN Count
                    End Select
                End If
            Next
        Else

        End If
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        If gbDetailPrint Then
            If anyChanges > 0 Then
                Call WriteLog(anyChanges & " Carpark(s) have been changed its color")
            Else
                Call WriteLog("No changes in color")
            End If
        End If

        If gnGreenCount < TotalCarpark Then
            ' Let it continue blinking
            ' =============================================

        Else     ' greenCount = TotalCarpark then
            ' Stop blinking
            BlinkCounter = MaxCountOfBlink + 10
        End If
        gnGreenCount = 0
    End Sub

    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    ' B array and Carparks array must be exactly the same data
    Private Function Sorting_Carparks() As Boolean
        ' sort carparks array based on priority level
        Dim I As Integer = 0
        Dim J As Integer = 0
        Dim K As Integer = 0
        Dim temp As recCarpark
        Dim bChanged As Boolean = False
        Dim LastIndex As Integer
        ReDim temp.status(12)             '(11)
        For K = 0 To 12               '11
            temp.status(K) = 0
        Next
        LastIndex = TotalCarpark - 2
        ' Sorting Phase: 1/2
        For J = 0 To LastIndex
            For I = 0 To LastIndex - J
                If Carparks(I).nPriority < Carparks(I + 1).nPriority Then
                    ' ===========================================================
                    temp.nPriority = Carparks(I).nPriority
                    temp.stIP = Carparks(I).stIP
                    temp.stMsg = Carparks(I).stMsg
                    temp.stName = Carparks(I).stName
                    temp.stID = Carparks(I).stID           ' Note: Carpark ID is not Array Index
                    temp.update_dt = Carparks(I).update_dt
                    ' ===========================================================
                    For K = 0 To 12           '11
                        temp.status(K) = Carparks(I).status(K)
                    Next
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    Carparks(I).nPriority = Carparks(I + 1).nPriority
                    Carparks(I).stIP = Carparks(I + 1).stIP
                    Carparks(I).stMsg = Carparks(I + 1).stMsg
                    Carparks(I).stName = Carparks(I + 1).stName
                    Carparks(I).stID = Carparks(I + 1).stID
                    Carparks(I).update_dt = Carparks(I + 1).update_dt
                    ' ============================================================
                    For K = 0 To 12              '11
                        Carparks(I).status(K) = Carparks(I + 1).status(K)
                    Next
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    Carparks(I + 1).nPriority = temp.nPriority
                    Carparks(I + 1).stIP = temp.stIP
                    Carparks(I + 1).stMsg = temp.stMsg
                    Carparks(I + 1).stName = temp.stName
                    Carparks(I + 1).stID = temp.stID
                    Carparks(I + 1).update_dt = temp.update_dt
                    For K = 0 To 12              '11
                        Carparks(I + 1).status(K) = temp.status(K)
                    Next
                    ' ==========================================================
                    bChanged = True
                End If
            Next
        Next
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        ' Sorting Phase: 2/2
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        For J = 0 To LastIndex
            For I = 0 To LastIndex - J
                If Carparks(I).stName > Carparks(I + 1).stName And _
                        Carparks(I).nPriority = Carparks(I + 1).nPriority Then
                    ' =======================================================
                    temp.nPriority = Carparks(I).nPriority
                    temp.stIP = Carparks(I).stIP
                    temp.stMsg = Carparks(I).stMsg
                    temp.stName = Carparks(I).stName
                    temp.stID = Carparks(I).stID
                    temp.update_dt = Carparks(I).update_dt
                    ' =======================================================
                    For K = 0 To 12              '11
                        temp.status(K) = Carparks(I).status(K)   ' status array also need to move
                    Next
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    Carparks(I).nPriority = Carparks(I + 1).nPriority
                    Carparks(I).stIP = Carparks(I + 1).stIP
                    Carparks(I).stMsg = Carparks(I + 1).stMsg
                    Carparks(I).stName = Carparks(I + 1).stName
                    Carparks(I).stID = Carparks(I + 1).stID
                    Carparks(I).update_dt = Carparks(I + 1).update_dt
                    ' =======================================================
                    For K = 0 To 12                  '11
                        Carparks(I).status(K) = Carparks(I + 1).status(K)
                    Next
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    Carparks(I + 1).nPriority = temp.nPriority
                    Carparks(I + 1).stIP = temp.stIP
                    Carparks(I + 1).stMsg = temp.stMsg
                    Carparks(I + 1).stName = temp.stName
                    Carparks(I + 1).stID = temp.stID
                    Carparks(I + 1).update_dt = temp.update_dt
                    For K = 0 To 12              '11
                        Carparks(I + 1).status(K) = temp.status(K)
                    Next
                    ' -----------------------------------------------------
                    bChanged = True
                End If
            Next
        Next

        Return bChanged
    End Function

    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    '                 Below are UDP Incoming Messages Reading Routines
    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    Public Sub ReceiveMessages()
        Try
            Dim receiveBytes As [Byte]() = receivingUdpClient.Receive(RemoteIpEndPoint)
            txtIP.Text = RemoteIpEndPoint.Address.ToString
            'Dim BitDet As BitArray
            'BitDet = New BitArray(receiveBytes)

            'Dim strReturnData As String = System.Text.Encoding.Unicode.GetString(receiveBytes)
            'Console.WriteLine("A message is received and being processed")
            'TextBox1.Text = TextBox1.Text & vbCrLf & "A message is received and being processed"
            'Console.WriteLine("The length of the message is ")
            'TextBox1.Text = TextBox1.Text & vbCrLf & "The length of the message is "
            'Console.WriteLine(receiveBytes.Length)
            'TextBox1.Text = TextBox1.Text & receiveBytes.Length
            'Console.WriteLine(strReturnData)
            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            'TextBox2.Text = strReturnData
            'Process_Incoming_Message(strReturnData)
            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            'TextBox1.Text = TextBox1.Text + vbCrLf + "The message received is """
            'TextBox1.Text = TextBox1.Text & Encoding.ASCII.GetChars(receiveBytes) + """"
            'TextBox1.Text = TextBox1.Text & vbCrLf
            'TextBox2.Text = Encoding.ASCII.GetChars(receiveBytes)
            Dim tmpMsg As String = Encoding.ASCII.GetChars(receiveBytes)
            Dim tmpMsg2 As String = Winsock1_DataArrival(tmpMsg)

            If gbDebugPrint Then Call WriteLog("(5) Income Message (Raw): " & tmpMsg) ' log original message

            'TextBox3.Text = Winsock1_DataArrival(Encoding.ASCII.GetChars(receiveBytes))

            If gbDebugPrint Then Call WriteLog("(6) " & tmpMsg2) ' log original message
            ' this routine is moved from Gateway application
            Process_Incoming_Message(tmpMsg2)
            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            ' Comment start: no need to show on screen bit by bit
            'Dim tempStr As String = ""
            'Dim tempStr2 As String = ""
            'If chkBit.Checked = True Then
            '    Dim i As Integer
            '    i = 0
            '    Dim j As Integer
            '    j = 0
            '    Dim line As Integer
            '    line = 0
            '    TextBox1.Text = TextBox1.Text + line.ToString & ") "
            '    For j = 0 To BitDet.Length - 1
            '        If BitDet(j) = True Then
            '            Console.Write("1 ")
            '            tempStr2 = tempStr
            '            tempStr = " 1" + tempStr2
            '        Else
            '            Console.Write("0 ")
            '            tempStr2 = tempStr
            '            tempStr = " 0" + tempStr2
            '        End If
            '        i += 1
            '        If i = 8 And j <= (BitDet.Length - 1) Then
            '            line += 1
            '            i = 0
            '            TextBox1.Text = TextBox1.Text + tempStr
            '            tempStr = ""
            '            tempStr2 = ""
            '            TextBox1.Text = TextBox1.Text + vbCrLf
            '            If j <> (BitDet.Length - 1) Then
            '                TextBox1.Text = TextBox1.Text + line.ToString & ") "
            '                Console.WriteLine()
            '            End If
            '        End If
            '    Next
            'End If
            'TextBox1.Text = TextBox1.Text & vbCrLf
            ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            ' Comment end: no need to show on screen bit by bit
            NewInitialize()
        Catch e As Exception
            Console.WriteLine(e.Message)
        End Try
    End Sub

    ' below routine is moved from gateway
    'Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Private Function Winsock1_DataArrival(ByVal bytesTotal As String) As String
        Dim strData As String
        Dim I As Integer
        Dim J As Integer
        Dim station1 As String              ' station Name
        Dim station1_id As Integer        ' station ID
        Dim description1 As String
        Dim carpark1 As Integer
        Dim carpark2 As String
        Dim finalmsg As String
        Dim status1() As String
        Dim counting1 As Integer
        Dim priority1 As Integer
        Dim Command1 As String = ""

        On Error Resume Next
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        'Winsock1.GetData(strData, vbString)
        strData = bytesTotal
        'Call WriteLog(strData)

        station1 = ""
        description1 = ""

        ' Process_Message function handle both SATAUS and ALARM message based on the value of command
        carpark1 = Process_Message(strData, station1, description1, station1_id, Command1)    ' carpark1 = Carpark ID
        gstCarparkID = carpark1           ' for broadcasting
        gstStationName = station1
        gstStationID = station1_id
        gstCommand = Command1

        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

        carpark2 = "" 'UCase(INIRead("Carpark", Str(carpark1), INIFile))   ' carpark2 = Carpark Name

        'If Len(carpark2) = 0 Then
        '    Exit Function
        'End If

        If gstMsgType = "ALARM" Then
            'finalmsg = carpark2 & "  " & station1 & "   " & description1
            gstMessage = description1           ' for broadcasting message
            gstStatus = "ALARM"         ' pass ALARM to SCMS-Remote Control
            ' if value is not ALARM, it is STATUS message
        Else
            status1 = Split(description1, ",")
            counting1 = 0
            priority1 = 0
            'UCase(ArrayStatus(I).status)
            gstMessage = ""
            gstStatus = description1 '""
            For I = 0 To 12 'gnTotalStatus - 1             ' All NON zero means ERROR
                '                Select Case I
                '                    Case 10, 11, 12: status1(I) = 0
                '                End Select

                If "0" <> status1(I) Then    ' Note: Show the most priority message
                    ' finalmsg is for display

                    ' below is for broadcasting
                    ' If there are more than one message, need to compare priority LEVEL
                    ' Show the most priority message
                    For J = 0 To Status_Array_Size - 1
                        If NewStatusArray(J).status1 = I + 1 And _
                            NewStatusArray(J).value1 = Val(status1(I)) Then
                            If NewStatusArray(J).priority1 > priority1 Then
                                gstMessage = NewStatusArray(J).message1
                                priority1 = NewStatusArray(J).priority1
                            End If

                            'Exit For
                        End If
                    Next

                    '                    If ArrayStatus(I).status > priority1 Then
                    '                        gstMessage = ArrayStatus(I).alarm_message
                    '                        gstStatus = description1
                    '                        priority1 = ArrayStatus(I).status
                    '                        finalmsg = carpark2 & "  " & station1 & "   " & ArrayStatus(I).alarm_message
                    '                    End If
                    '
                    counting1 = counting1 + 1          ' count total error bit (non-zero values)
                End If
            Next

            If counting1 = 0 Then             ' all zeros mean no error, OK
                gstMessage = "OK"
                ' below also change to 13 zeros
                gstStatus = "0,0,0,0,0,0,0,0,0,0,0,0,0"           '13 12 zeros means no error (OK)
            End If
            'finalmsg = "[" & carpark2 & "|" & station1 & "|" & station1_id & "|" & gstMessage & "|" & priority1 & "]"
        End If

        finalmsg = Command1_Click_New()          ' Broadcast message to workstation
        Return finalmsg        ' finally return processed message for remote control application

    End Function

    Private Function Command1_Click_New() As String
        Dim stSendString As String
        'With Winsock2
        '    .Close()
        '    .Protocol = sckUDPProtocol
        '    .RemoteHost = "255.255.255.255"
        '    .RemotePort = UDPBroadCast
        '  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        '  broadcasting carparkid and message to workstation via udp port
        ' Sending message have 3 parts as follow
        ' (1) carpark_id (2) message body (3) message type separated by vertical line
        stSendString = ""

        stSendString = stSendString & gstCarparkID & "|"
        stSendString = stSendString & gstMessage & "|"
        stSendString = stSendString & gstStationName & "|"
        stSendString = stSendString & gstStationID & "|"
        stSendString = stSendString & gstCommand & "|"
        stSendString = stSendString & gstStatus

        ' Note: gstStatus = "ALARM", it is considered ALARM message and
        ' gstStatus <> "ALARM", it is considered STATUS message and the value is itself.
        ' it is a string of 13 digits separated by comma

        ' gstStation is added in sending message on 19-09-2013

        '.SendData(stSendString)
        Return stSendString
        '' respective carparkid and message are also
        '' stored in Setting File of Work Station application
        'End With
        'With RichTextBox2
        '    .Text = Format(Now, "hh:mm:ss") & " Broadcast: " & "<" & stSendString & ">" & vbCrLf & .Text
        '    ' Note: < and > are not in sending message (only for display)
        '    .SelStart = 0
        '    .SelLength = Len("Broadcast: " & stSendString) + 11
        '    .SelColor = vbRed
        'End With
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

    End Function

    Private Function Process_Message(ByVal msg As String, ByRef station As String, _
                                    ByRef Description As String, ByRef station_id As Integer, _
                                    ByRef stCommand As String) As Integer
        Dim st() As String
        Dim I As Integer
        Dim J As Integer
        Dim st1 As String
        Dim K As Integer
        Dim st2 As String
        Dim Command1 As String
        Dim stationid As String
        Dim st3() As String
        Dim message1 As String
        Dim FoundSemiColon As Boolean

        On Error Resume Next
        K = Len(msg)
        FoundSemiColon = False
        For I = 1 To K
            If Mid(msg, I, 1) = ";" Then
                FoundSemiColon = True
                Exit For
            End If
        Next I
        ' can do some validation here also
        If Mid(msg, 1, 1) = "[" And Mid(msg, K, 1) = "]" Then
            st1 = ""             ' [13|Exit 2|7|08|Antenna Error|]
            For I = 1 To K
                st2 = Mid(msg, I, 1)
                Select Case st2
                    Case "[", "]"
                    Case Else
                        st1 = st1 & st2
                End Select
            Next I
        Else
            Process_Message = -999
            Exit Function
        End If
        st = Split(st1, "|", -1, vbTextCompare)

        Process_Message = st(0)            ' return value is carpark id to be compare carpark_id from INI file
        station = st(1)                 ' station name
        stationid = st(2)               ' station id

        station_id = stationid          ' pass by ref station_id

        Command1 = st(3)                ' command
        stCommand = Command1
        ' ==============================================================================
        ' Status message have comma separated value in description
        ' Note: Actually 00 is for STATUS msg (come from Station) and Non 00 is for ALARM msg (come from PMS)
        ' But PMS also send 00, I cannot differentiate between STATUS and ALARM
        ' Finally use semi-colon to make sure BETWEEN Status and Alarm
        ' ==============================================================================
        If FoundSemiColon Then   ' Command1 = "00" Then         '     [374|Exit 2|7|00|0,0,0,0,0,0,0,0,0,0,0,2;0,20061023000418,|]
            gstMsgType = "STATUS"
            message1 = st(4)
            st3 = Split(message1, ";")
            Description = st3(0)
        Else
            gstMsgType = "ALARM"       ' [13|Exit 2|7|08|Antenna Error|]
            Description = st(4)        ' alarm message value is return by passing reference
        End If

    End Function

    Public Sub NewInitialize()
        Console.WriteLine("Thread *Thread Receive* reinitialized")
        ThreadReceive = New System.Threading.Thread(AddressOf ReceiveMessages)
        ThreadReceive.Start()
    End Sub

    Private Sub Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Clear.Click
        TextBox1.Text = "INFORMATION"
    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            receivingUdpClient.Close()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try                                 ' Stop button
            ThreadReceive.Abort()
            receivingUdpClient.Close()
            TextBox1.Text = "INFORMATION"
            TextBox1.Enabled = False
            Button2.Enabled = False
            Button1.Enabled = True
            txtIP.Text = ""
            txtSocket.ReadOnly = False
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try                                 ' Start button
            SocketNO = txtSocket.Text
            Call WriteLog("Start Button1_Click: SocketNO = " & SocketNO)

            receivingUdpClient = New System.Net.Sockets.UdpClient(SocketNO)
            ThreadReceive = New System.Threading.Thread(AddressOf ReceiveMessages)

            ThreadReceive.Start()
            TextBox1.Enabled = True
            Button2.Enabled = True
            Button1.Enabled = False
            txtSocket.ReadOnly = True
        Catch x As Exception
            'Console.WriteLine(x.Message)
            Call WriteLog(x.Message)
            TextBox1.Text = TextBox1.Text & vbCrLf & x.Message
        End Try
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox("Connect")
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox("View")
    End Sub

    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    ' tmrCarpark is used to check connection timeout for every carpark
    Private Sub tmrCarpark_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrCarpark.Tick
        ' Every n seconds (1.5 minutes) seconds) all carparks->update_dt
        ' Read Update_dt and then compare it with current time
        ' now - last update_dt > 15 minutes; show "NO CONNECTION" Message on the carpark
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        Dim I As Integer
        Dim n As Integer
        Dim st As String = ""
        Dim dt1 As DateTime
        Dim stn1 As String = ""
        Dim stLog As String = ""

        'Dim nWaitingSecond As Integer
        tmrCarpark.Enabled = False
        Exit Sub
        'nWaitingSecond = 1 'gnWaitingMinute * 60

        If gbDetailPrint Then Call WriteLog("tmrCarpark start to check time duration")

        Try
            For I = 0 To TotalCarpark - 1
                ' Carpark must be in operation (no need to check for NOT YET OPERATION Carpark)
                If UCase(B(I).Operation1) = "OPERATION" Then
                    ' Calculate the time duration between last Update_dt and Now (How many minutes?)
                    n = DateDiff(DateInterval.Minute, B(I).Update_dt, Now)
                    If gbDebugPrint Then
                        st = "Carpark: " & B(I).Label1.Text & "(" & B(I).CarparkID1 & "), Last update_dt = "
                        st = st & B(I).Update_dt & ", Current Time = " & Now
                        If gbDetailPrint Then Call WriteLog(st)

                        If gbDetailPrint Then Call WriteLog("Duration between last msg and current time = " & n)
                    End If

                    If n > gnWaitingMinute Then    'nWaitingSecond Then
                        st = "Carpark: " & B(I).Label1.Text & "(" & B(I).CarparkID1 & ") "
                        st = st & " doesn't receive any message from carpark within "
                        st = st & gnWaitingMinute & " minutes. Then  set to RED color. Msg = No Connection"
                        If gbDetailPrint Then Call WriteLog(st)

                        If gbNoConnectionShowRedEnable Then
                            'B(I).image1 = My.Resources.Resource1.red

                            Call Show_Carpark_Color("Carpark", I, "Red")
                            B(I).text2 = "No Connection"   ' label2 for message
                            Carparks(I).stMsg = "No Connection"
                            B(I).Prior1 = 3
                            Carparks(I).nPriority = 3
                            detectRed = True
                        Else
                            If gbDetailPrint Then Call WriteLog("gbNoConnectionShowRedEnable: FALSE")
                        End If

                    Else
                        ' Checking each station timeout in details
                        ' =========================================
                        ' Retrieve connection info based on Carpark
                        Dim dr2() As DataRow
                        Dim dr As DataRow
                        Dim count1 As Integer = 0

                        Try
                            dr2 = dtConn.Select("CarparkID = '" & B(I).CarparkID1 & "'") ', "Update_dt DESC")
                            count1 = 0
                            dt1 = Now     ' Find the station which received message with earliest time stamp
                            stn1 = -999

                            For Each dr In dr2          ' At least one station is timeout, set error to RED color
                                count1 += 1
                                If dr.Item(2) < dt1 Then    ' check with current time
                                    dt1 = dr.Item(2)
                                    stn1 = dr.Item(1)
                                End If
                            Next
                        Catch ex As Exception
                            Call WriteLog("Error handle 3 in tmrCarpark_Tick SUB")
                            Call WriteLog("Error message: " & ex.Message)
                        End Try

                        If gbDetailPrint Then
                            stLog = "CarparkID = " & B(I).CarparkID1 & "(" & B(I).Label1.Text & ")"
                            stLog = stLog & " Count1: " & count1 & ", Stn: " & stn1
                            Call WriteLog(stLog)
                        End If

                        If count1 > 0 Then                ' 0:cpid, 1:stn, 2:dt, 3:cpname
                            'dt1 = dr2(0).Item(2)
                            'stn1 = dr2(0).Item(1)
                            Try
                                n = DateDiff(DateInterval.Minute, dt1, Now)

                                If gbDetailPrint Then Call WriteLog("StationID = " & stn1 & " have " & n & " minutes")

                                If n > gnWaitingMinute Then
                                    ' Even one station timeout, show Error to RED color
                                    st = "Carpark: " & B(I).Label1.Text & "(" & B(I).CarparkID1 & ") "
                                    st = st & " StationID " & stn1 & " : "
                                    st = st & " doesn't receive any message from carpark within "
                                    st = st & gnWaitingMinute & " minutes. Set to RED. Msg = No Connection"

                                    If gbDetailPrint Then Call WriteLog(st)
                                    If gbNoConnectionShowRedEnable Then
                                        'B(I).image1 = My.Resources.Resource1.red

                                        Call Show_Carpark_Color("Carpark", I, "Red")
                                        B(I).text2 = "Stn: " & stn1 & " No Connection"   ' label2 for message
                                        Carparks(I).stMsg = "Stn: " & stn1 & " No Connection"
                                        B(I).Prior1 = 3
                                        Carparks(I).nPriority = 3
                                        detectRed = True
                                    Else
                                        If gbDetailPrint Then Call WriteLog("gbNoConnectionShowRedEnable: FALSE")
                                    End If
                                End If
                            Catch ex As Exception
                                Call WriteLog("Error handling 2 in tmrCarpark Sub")
                                Call WriteLog("Error message: " & ex.Message)
                            End Try

                        End If
                    End If

                    ' Carparks(I).nPriority = Priority
                    ' Carparks(I).stMsg = ShowMsg1           ' Store Message Value
                    ' Carparks(I).update_dt = Now            ' Store receiving time
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    '                  HANDLING INTERCOM TIMEOUT
                    ' If RED msg for Intercom is display more than 2 minutes; let it back to GREEN
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    ' Note: INTERCOM is no longer set to RED color according to Boss' instruction
                    ' so purposely don't let below codes for execution (commented)

                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    'If B(I).Prior1 = 3 And UCase(B(I).text2) = "INTERCOM" Then
                    '    If n > Intercom_Minute Then
                    '        st = "Carpark: " & B(I).Label1.Text & "(" & B(I).CarparkID1 & ") "
                    '        st = st & " set back from RED to GREEN after intercom timeout "
                    '        st = st & Intercom_Minute & " minutes"
                    '        If gbDebugPrint Then Call WriteLog(st)

                    '        B(I).image1 = My.Resources.Resource1.green
                    '        Select Case B(I).GroupIndex1
                    '            Case 0 : Show_Carpark_Color("Carpark", I, "Green1")
                    '            Case 1 : Show_Carpark_Color("Carpark", I, "Green2")
                    '            Case 2 : Show_Carpark_Color("Carpark", I, "Green3")
                    '        End Select
                    '        B(I).text2 = "No Problem"   ' label2 for message
                    '        Carparks(I).stMsg = "No Problem"
                    '        B(I).Prior1 = 1
                    '        Carparks(I).nPriority = 1
                    '    End If
                    'End If
                    ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                End If
            Next
            Call WriteLog("tmrCarpark Loop Finish successfully: For I = 0 To TotalCarpark - 1 ")
        Catch ex As Exception
            Call WriteLog("Error handling in Sub: trmCarpark_Tick")
            Call WriteLog("Error message = " & ex.Message)
        End Try

        Call write_dtConn_to_Listbox()

        tmrCarpark.Enabled = False
    End Sub

    Private Sub tmrBlink_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrBlink.Tick
        ' If gbDetailPrint Then Call WriteLog("BlinkCounter = " & BlinkCounter)
        ' Show Forecolor alternately to simulate Blinking effect
        ' If BlinkCounter Mod 2 = 0 Then
        '    Legend(IndexofRed).G(IndexofRed).lblCarpark.ForeColor = Color.Black
        ' Else
        '    Legend(IndexofRed).G(IndexofRed).lblCarpark.ForeColor = Color.Yellow
        ' End If
        tmrBlink.Enabled = False
        Exit Sub
        If Legend(IndexofRed).G(IndexofRed).lblCarpark.ForeColor = Color.Yellow Then
            Legend(IndexofRed).G(IndexofRed).lblCarpark.ForeColor = Color.Black
        Else
            Legend(IndexofRed).G(IndexofRed).lblCarpark.ForeColor = Color.Yellow
        End If

        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        If SoundAlarmCounter Mod 4 = 0 Then         ' Every 2 seconds, make sound
            sndPlaySound("C:\WINDOWS\Media\tada.wav", 1)

        End If
        SoundAlarmCounter += 1
        If SoundAlarmCounter > 9999 Then
            SoundAlarmCounter = 0
        End If
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

        ' BlinkCounter += 1 ' increase counter to prevent non-stop blinking

        ' Above line is commented let blinking forever, only stop there is ALL GREEN
        ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        If BlinkCounter > MaxCountOfBlink Then
            Legend(IndexofRed).G(IndexofRed).lblCarpark.ForeColor = Color.White
            tmrBlink.Enabled = False       ' make sure to stop blinking
            BlinkCounter = 0               ' reset blink counter
            SoundAlarmCounter = 0
            StillBlinking = False          ' reset blinking flag
        End If

        tmrBlink.Enabled = False
    End Sub

    ' Timer2 is always running and checking the boolean flag - detectRed
    ' If he found detectRed and Not blink yet, it call Blink function
    ' Use Timer2 and detectRed is why? the reason 
    ' is that it is not working by calling directly tmrBlink
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False
        Exit Sub
        If detectRed Then
            detectRed = False
            ' Then need to reset detectRed Flag to prevent further blinking again
            If Not StillBlinking Then    ' check

                Call Show_Blinking()

                'Button3_Click(Nothing, Nothing)
            End If
            'MsgBox(detectRed)
            'detectRed = False
        End If
    End Sub

    Private Sub picSJS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picSJS.Click
        Dim ans As Integer = 0
        Dim st As String = ""
        Dim bFlag As Boolean = False
        'st = InputBox("Password?", "View", "*")
        'MsgBox(st)
        Exit Sub

        Password1.ShowDialog()
        st = stPassword
        bFlag = Check_Password(st)
        If bFlag Then
            ' load the from to show Decision table and History record
            'lstConnection.Visible = True
            Message_History.ShowDialog()
            'lstConnection.Visible = False
            'Else
            '    MsgBox("Password wrong")
        End If
    End Sub

    Private Function Check_Password(ByVal ps As String) As Boolean

        If UCase(ps) = "SROS2014" Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub tmrUpdate_Tick(sender As Object, e As EventArgs) Handles tmrUpdate.Tick
        tmrUpdate.Enabled = False
        Exit Sub
        Try
            ' Read available lots from database - table with the join statement of View - UNION - select (Hardcoded)
            Call read_Carpark_Lot_Available_Number()
        Catch ex As Exception
            Call WriteLog("Error in reading Carpark Lot Available Number")
        End Try
        tmrUpdate.Enabled = False
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox("a")

    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Dim I As Integer = 0
        Dim J As Integer = 0

        For I = 0 To 4
            Selected_Node(I) = 0
        Next

        For I = 0 To 4
            For J = 0 To 2
                B(I).G(J).Visible = False
            Next
            B(I).G(0).Visible = True
        Next
        order1 = 0
        btnPath.Enabled = False
    End Sub

    Private Sub lblTitle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTitle.Click

    End Sub

    Private Sub UserControl21_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPath.Click
        'MsgBox("The Results are as follows:" & vbCrLf)
        Dim dummy1 As Integer = 123

        txtPath.Text = ""
        For i As Integer = 0 To 4
            'gnSource = "c"
            'gnDestination = "a"
            If Selected_Node(i) = 1 Then
                gnSource = labels(i)
            End If
            If Selected_Node(i) = 2 Then
                gnDestination = labels(i)
            End If
        Next

        Call Preprocessing()

        txtPath.Text = txtPath.Text & "Shortest path:" & vbCrLf

        dest_index = find_index_mapping(gnDestination)

        print_path(dest_index)

        txtPath.Text = txtPath.Text & vbCrLf

        dummy1 = dist(dest_index)

        txtPath.Text = txtPath.Text & "Shortest Distance = " & dummy1
    End Sub

    Private Sub print_path(ByVal j As Integer)

        'If parent1(i) = -1 Then
        'MsgBox("a" & i)97
        'Convert(char, 97+i)))

        'Return
        'End If 
        If parent1(j) = -1 Then
            'Console.Write(Char.ConvertFromUtf32(Asc("a"c) + i) & " -> ")
            txtPath.Text = txtPath.Text & Char.ConvertFromUtf32(Asc("a"c) + j) & " -> "
            Return
        End If
        print_path(parent1(j))

        If dest_index = j Then
            txtPath.Text = txtPath.Text & Char.ConvertFromUtf32(Asc("a"c) + j)
        Else
            txtPath.Text = txtPath.Text & Char.ConvertFromUtf32(Asc("a"c) + j) & " -> "
        End If


    End Sub

    Private Sub Preprocessing()
        'vertices = InputBox("How many vertices?")
        'MsgBox(vertices)
        Dim dummy1 As Integer = 200
        Dim count1 As Integer
        Dim u As Integer
        Dim v As Integer

        source_index = find_index_mapping(gnSource)

        'MsgBox(source_index)
        'TextBox1.Text = "Source Index = " & source_index & vbCrLf

        init1()
        dist(source_index) = 0
        For count1 = 0 To 3
            u = path_choice()
            spt_set(u) = 1
            For v = 0 To vertices - 1
                If (spt_set(v) = 0) And (graph(u, v) > 0) And (dist(u) <> inf1) And (dist(u) + graph(u, v) < dist(v)) Then
                    dist(v) = dist(u) + graph(u, v)
                    parent1(v) = u

                End If
            Next

        Next

    End Sub

    Private Function path_choice() As Integer
        Dim min As Integer = inf1
        Dim min_index As Integer = -1
        Dim i As Integer
        For i = 0 To vertices - 1
            If (spt_set(i) = 0) And (dist(i) <= min) Then
                min = dist(i)
                min_index = i

            End If
        Next
        Return min_index

    End Function

    Private Sub init1()
        Dim i As Integer
        For i = 0 To 4
            dist(i) = inf1
            spt_set(i) = 0
            parent1(i) = -1
        Next
    End Sub

    Private Function find_index_mapping(ByVal inputLabel As String) As Integer
        find_index_mapping = -999
        Dim label1 As String
        Dim i As Integer

        For i = 0 To 4
            label1 = labels(i)
            If label1 = inputLabel Then
                Return i
            End If
        Next
    End Function

    Private Sub lblNotify_Click(sender As Object, e As EventArgs) Handles lblNotify.Click

    End Sub
End Class

