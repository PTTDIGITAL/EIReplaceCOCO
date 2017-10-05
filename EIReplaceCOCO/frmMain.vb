Imports EIReplaceCOCO.Org.Mentalis.Files
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices

Public Class frmMain

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function FindWindowEx(ByVal parentHandle As IntPtr, ByVal childAfter As IntPtr, ByVal lclassName As String, ByVal windowTitle As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As Integer, ByVal lParam As String) As Integer
    End Function

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        writeLogResult("Start Program")
        Timer1.Interval = 100
        Timer1.Start()
        txtTransLog.TabStop = False

        Try
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
        Catch ex As Exception
            MsgBox("ไม่สามารถเชื่อมต่อฐานข้อมูลได้", MsgBoxStyle.OkOnly)
            writeLogResult("Connection State is close.")
            Application.Exit()
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        lblDateTime.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
    End Sub

    Private Sub pbExit_Click(sender As Object, e As EventArgs) Handles pbExit.Click

        Using New Centered_MessageBox(Me)
            Dim confirm As DialogResult = MessageBox.Show("ต้องการปิดโปรแกรมใช่หรือไม่", "", MessageBoxButtons.OKCancel)
            If (confirm.Equals(DialogResult.OK)) Then
                writeLogResult("Exit Program")
                Application.Exit()
            End If
        End Using
    End Sub

    Private Sub pbExport_MouseHover(sender As Object, e As EventArgs) Handles pbExport.MouseHover
        Dim ToolTip1 As New ToolTip
        ToolTip1.SetToolTip(pbExport, "นำออกข้อมูล")
    End Sub

    Private Sub pbImport_MouseHover(sender As Object, e As EventArgs) Handles pbImport.MouseHover
        Dim ToolTip1 As New ToolTip
        ToolTip1.SetToolTip(pbImport, "นำเข้าข้อมูล")
    End Sub

    Private Sub pbExit_MouseHover(sender As Object, e As EventArgs) Handles pbExit.MouseHover
        Dim ToolTip1 As New ToolTip
        ToolTip1.SetToolTip(pbExit, "ปิด")
    End Sub

    Private Sub pbExport_Click(sender As Object, e As EventArgs) Handles pbExport.Click
        Try
            txtTransLog.Text = ""
            '#Save file Result, File Name _Result.txt  in Startup Path
            Dim strFileResule As String = Application.StartupPath & "\" & "_Result_" & DateTime.Now.ToString("yyyyMMdd") & ".txt"
            Dim strdate As String = DateTime.Now.ToString("yyyyMMdd hh:MM:ss") & "  "
            Dim sw_rs As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(strFileResule, True)
            sw_rs.WriteLine(strdate & "เริ่มต้นนำออกข้อมูล")

            Dim frm As New frmConfirmPassword
            frm.Mode = 1
            If frm.ShowDialog(Me) = DialogResult.Cancel Then
                sw_rs.WriteLine(strdate & "Confirm Password : Cancel")
                sw_rs.Close()
                Exit Sub
            Else
                sw_rs.WriteLine(strdate & "Confirm Password : Success")
                txtTransLog.Text = strdate & "เริ่มต้นนำออกข้อมูล"
            End If

            EnableButton(False)

            '#1. Connect Local Server Only
            '#2. Delete All File In Floder Export & Delete _Result.txt
            Dim path As String = Application.StartupPath & "\Export"
            If Directory.Exists(path) Then
                For Each _file As String In Directory.GetFiles(path)
                    File.Delete(_file)
                Next
            Else
                Directory.CreateDirectory(path)
            End If

            '#3. Export text file Name Same Table Name
            Dim arr() As String = {"APP_DATA", "TBPOS_PUMP_ALLOW", "TBMATERIAL", "TBBOM_USAGE", "TBMATTERIAL_SITE",
                "TBCONVERSION", "LKMAT_GROUP3", "LKDIVISION", "TBMAT_RECOMMEND", "APP_CONFIG", "POS_CONFIG", "TBUSER",
                "TSJOURNAL", "TSJOURNAL_DETAIL", "TSJOURNAL_PAYMENT", "TBPERIODS", "TBHOSE_HISTORY", "TBMATERIAL_HISTORY",
                "TBTANK_HISTORY", "TBPAY_IN", "TBPAYIN_PERIOD_LOG"}
            For Each item As String In arr
                Dim table_name As String = item
                Dim sql As String = ""
                Select Case table_name
                    Case "APP_DATA", "TBPOS_PUMP_ALLOW", "TBMATERIAL", "TBBOM_USAGE", "TBMATTERIAL_SITE", "TBCONVERSION", "LKMAT_GROUP3", "LKDIVISION",
                         "TBMAT_RECOMMEND", "APP_CONFIG", "POS_CONFIG", "TBUSER", "TSJOURNAL", "TSJOURNAL_DETAIL", "TSJOURNAL_PAYMENT", "TBPAYIN_PERIOD_LOG"
                        sql &= "SELECT * FROM " & table_name
                    Case "TBPAY_IN"
                        sql &= "SELECT [PAYIN_ID]
                              ,[REFBILL_NO]
                              ,[TRANSFER_DATE]
                              ,[TRAN_DATE]
                              ,[BUS_DATE]
                              ,REPLACE(REPLACE([SHIFT_DESCRIPTION], CHAR(13), '$$'), CHAR(10), '&&') as SHIFT_DESCRIPTION
                              ,[LAST_CLOSE_SHIFT_DT]
                              ,[FILEPATH]
                              ,[FILENAME]
                              ,[TYPE]
                              ,[PAYMENT_TYPE]
                              ,[AMOUNTREC]
                              ,[AMOUNT]
                              ,[AMOUNT_DIFF]
                              ,REPLACE(REPLACE([REMARK], CHAR(13), '$$'), CHAR(10), '&&') REMARK 
                              ,[STATUS_SAP]
                              ,[STATUS]
                              ,[NO_SALE_STATUS]
                              ,[CREATEDATE]
                              ,[CREATEBY]
                              ,[UPDATEDATE]
                              ,[UPDATEBY]
                          FROM " & table_name
                    Case "TBPERIODS"
                        sql &= "SELECT * FROM TBPERIODS WHERE DAY_ID IN (SELECT DISTINCT  DAY_ID FROM TSJOURNAL)"

                    Case "TBHOSE_HISTORY"
                        sql &= "SELECT * FROM TBHOSE_HISTORY WHERE PERIOD_ID IN (SELECT PERIOD_ID FROM TBPERIODS WHERE DAY_ID IN (SELECT TOP 1 DAY_ID FROM TSJOURNAL))"

                    Case "TBMATERIAL_HISTORY"
                        sql &= "SELECT * FROM TBMATERIAL_HISTORY  WHERE PERIOD_ID IN (SELECT PERIOD_ID FROM TBPERIODS WHERE DAY_ID IN (SELECT TOP 1 DAY_ID FROM TSJOURNAL))"

                    Case "TBTANK_HISTORY"
                        sql &= "SELECT * FROM TBTANK_HISTORY  WHERE PERIOD_ID IN (SELECT PERIOD_ID FROM TBPERIODS WHERE DAY_ID IN (SELECT TOP 1 DAY_ID FROM TSJOURNAL))"
                End Select

                Dim da As New SqlDataAdapter(sql, ConnStr)
                Dim dt As New DataTable
                da.Fill(dt)

                '#4. Save text file in floder Application.StartupPath & Export
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then

                    If table_name = "TBUSER" Then
                        For j As Integer = 0 To dt.Rows.Count - 1
                            dt.Rows(j)("PASSWORD") = Decrypt(dt.Rows(j)("USERNAME").ToString, dt.Rows(j)("PASSWORD").ToString)
                        Next
                    End If




                    Dim strFile As String = path & "\" & table_name & ".txt"

                    'Add column.
                    Dim strcolumn As String = ""
                    For Each column As DataColumn In dt.Columns
                        strcolumn += "|" & column.ColumnName
                    Next
                    strcolumn = strcolumn.Substring(1)

                    'Add data.
                    Dim result As New StringBuilder
                    Dim i As Integer = 0
                    For Each row As DataRow In dt.Rows
                        Dim line As String = ""
                        For Each column As DataColumn In dt.Columns
                            line += "|" & row(column.ColumnName)
                        Next
                        result.AppendLine(line.Substring(1))
                        i += 1
                    Next
                    Using sw As StreamWriter = New StreamWriter(strFile)
                        sw.WriteLine(strcolumn)
                        sw.WriteLine(result.ToString)
                        sw.Close()
                    End Using

                    '#Save file Result, File Name _Result.txt  in Startup Path
                    sw_rs.WriteLine(strdate & table_name & "  (" & dt.Rows.Count.ToString & " Row)")

                    Application.DoEvents()
                    txtTransLog.Text = strdate & table_name & "  (" & dt.Rows.Count.ToString & " Row)" & vbCrLf & txtTransLog.Text
                    Threading.Thread.Sleep(100)

                End If
            Next ' end for arr
            txtTransLog.Text = strdate & "สิ้นสุดการนำออกข้อมูล" & vbCrLf & txtTransLog.Text

            '#5. Final Export Save file Result, File Name _Result.txt  in Startup Path
            sw_rs.WriteLine(strdate & "สิ้นสุดการนำออกข้อมูล")
            sw_rs.Close()

            EnableButton(True)

        Catch ex As Exception
            writeLogResult("Export Fail")

            Using New Centered_MessageBox(Me)
                MessageBox.Show(ex.ToString(), "", MessageBoxButtons.OK)
            End Using

        End Try
    End Sub

    Private Sub pbImport_Click(sender As Object, e As EventArgs) Handles pbImport.Click
        Dim strFileResule As String = Application.StartupPath & "\" & "_Result_" & DateTime.Now.ToString("yyyyMMdd") & ".txt"
        Dim sw_rs As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(strFileResule, True)
        Dim strdate As String = DateTime.Now.ToString("yyyyMMdd hh:MM:ss") & "  "
        Try
            txtTransLog.Text = ""
            '#Save file Result, File Name _Result.txt  in Startup Path
            sw_rs.WriteLine(strdate & "เริ่มต้นนำเข้าข้อมูล")

            Dim frm As New frmConfirmPassword
            frm.Mode = 2
            If frm.ShowDialog(Me) = DialogResult.Cancel Then
                sw_rs.WriteLine(strdate & "Confirm Password : Cancel")
                sw_rs.Close()
                Exit Sub
            Else
                sw_rs.WriteLine(strdate & "Confirm Password : Success")
                txtTransLog.Text = strdate & "เริ่มต้นนำเข้าข้อมูล"
            End If

            EnableButton(False)

            Dim dt_appconfig As New DataTable
            Dim path As String = Application.StartupPath & "\Export"
            If Directory.Exists(path) Then
                For Each _file As String In Directory.GetFiles(path)
                    Dim file_name As String = System.IO.Path.GetFileName(_file)
                    Dim str() As String = file_name.Split(".")
                    Dim table_name As String = ""
                    If str.Length > 0 Then
                        table_name = str(0)
                    End If

                    'Delete Data
                    DeleteData(table_name)

                    'Insert Data
                    Dim dt As New DataTable
                    Dim dr As DataRow
                    Dim i As Integer = 0
                    Dim sr As StreamReader = New StreamReader(_file)
                    Dim line As String = ""
                    Do While sr.Peek() >= 0
                        line = sr.ReadLine()
                        If i = 0 AndAlso line <> "" Then
                            Dim strColumn() As String = line.Split("|")
                            For j As Integer = 0 To strColumn.Length - 1
                                dt.Columns.Add(strColumn(j))
                            Next
                        Else
                            Dim strValue() As String = line.Split("|")
                            dr = dt.NewRow
                            If strValue.Length = dt.Columns.Count Then
                                For j As Integer = 0 To strValue.Length - 1
                                    dr(j) = strValue(j)
                                Next
                                If dr(0).ToString <> "" Then
                                    dt.Rows.Add(dr)
                                End If
                            End If
                        End If

                        i = i + 1
                    Loop
                    sr.Close()

                    Dim result As String = ""
                    If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                        result = InsertData(table_name, dt)
                    End If

                    If table_name.ToUpper = "APP_CONFIG" Then
                        dt_appconfig = dt
                    End If


                    If result <> "" Then
                        txtTransLog.Text = strdate & table_name & "ไม่สามารถนำเข้าข้อมูลได้ !" & vbCrLf & result & vbCrLf & txtTransLog.Text
                        sw_rs.WriteLine(strdate & table_name & "ไม่สามารถนำเข้าข้อมูลได้ !" & vbCrLf & result)
                        sw_rs.Close()
                        Exit Sub
                    End If


                    If table_name.ToUpper <> "APP_CONFIG" Then
                        'Save Transection Log
                        Application.DoEvents()
                        txtTransLog.Text = strdate & table_name & "  (" & dt.Rows.Count.ToString & " Row)" & vbCrLf & txtTransLog.Text
                        Threading.Thread.Sleep(100)

                        'Save Result Log
                        sw_rs.WriteLine(strdate & table_name & "  (" & dt.Rows.Count.ToString & " Row)")
                    End If

                Next

                'call sp sp_Import_Product_To_Inventory
                Application.DoEvents()
                If CallSPImportProduct() = False Then
                    txtTransLog.Text = strdate & "พบปัญหาในการนำเข้าข้อมูล : Call sp_Import_Product_To_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(strdate & "พบปัญหาในการนำเข้าข้อมูล : Call sp_Import_Product_To_Inventory")
                Else
                    txtTransLog.Text = strdate & "Call sp_Import_Product_To_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(strdate & "Call sp_Import_Product_To_Inventory")
                End If

                'Update ISSHOWINPOS =0 , ISRECOMMEND=0
                Application.DoEvents()
                Dim ret0 As Integer = UpdateDefaultProduct()
                txtTransLog.Text = strdate & "Update ISSHOWINPOS = 0,ISSHOWINPOS =0 (" & ret0 & " Row)" & vbCrLf & txtTransLog.Text
                sw_rs.WriteLine(strdate & "Update ISSHOWINPOS = 0,ISSHOWINPOS =0 (" & ret0 & " Row)")
                Threading.Thread.Sleep(100)

                'Update ข้อมูลที่ POSDB.dbo.PRODUCTS.ISSHOWINPOS ให้เป็น 1 เฉพาะรหัสผลิตภัณฑ์ที่มีอยู่ในตาราง POSDB.dbo.TBMATERIAL_SITE
                Application.DoEvents()
                Dim ret1 As Integer = UpdateISSHOWINPOS()
                txtTransLog.Text = strdate & "Update ISSHOWINPOS (" & ret1 & " Row)" & vbCrLf & txtTransLog.Text
                sw_rs.WriteLine(strdate & "Update ISSHOWINPOS (" & ret1 & " Row)")
                Threading.Thread.Sleep(100)

                'Update ข้อมูลที่ POSDB.dbo.PRODUCTS.ISRECOMMEND ให้เป็น 1 เฉพาะรหัสผลิตภัณฑ์ที่มีอยู่ในตาราง POSDB.dbo.TBMAT_RECOMMENED
                Application.DoEvents()
                Dim ret2 As Integer = UpdateISRECOMMEND()
                txtTransLog.Text = strdate & "Update ISRECOMMEND (" & ret2 & " Row)" & vbCrLf & txtTransLog.Text
                sw_rs.WriteLine(strdate & "Update ISRECOMMEND (" & ret2 & " Row)")
                Threading.Thread.Sleep(100)

                'Create sp_Initial_LUBE_Stock_Inventory
                Dim retsp As String = ""
                Dim retdropsp As String = CheckExistsSP("sp_Initial_LUBE_Stock_Inventory")
                If retdropsp = "" Then
                    retsp = CreateStoreInitialLUBE()
                    If retsp = "" Then
                        Application.DoEvents()
                        txtTransLog.Text = strdate & "Create sp_Initial_LUBE_Stock_Inventory" & vbCrLf & txtTransLog.Text
                        sw_rs.WriteLine(strdate & "Create sp_Initial_LUBE_Stock_Inventory")
                        Threading.Thread.Sleep(100)
                    End If
                Else
                    txtTransLog.Text = strdate & "พบปัญหาในการนำเข้าข้อมูล : Drop sp_Initial_LUBE_Stock_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(strdate & "พบปัญหาในการนำเข้าข้อมูล :  Drop sp_Initial_LUBE_Stock_Inventory" & retdropsp)
                End If

                If retsp <> "" Then
                    Application.DoEvents()
                    txtTransLog.Text = strdate & "Cant Create sp_Initial_LUBE_Stock_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(strdate & "Cant Create sp_Initial_LUBE_Stock_Inventory" & retsp)
                    Threading.Thread.Sleep(100)
                End If


                'call sp_Initial_LUBE_Stock_Inventory
                Application.DoEvents()
                Dim ret_CallSPInitialLUBE As String = CallSPInitialLUBE()
                If ret_CallSPInitialLUBE <> "" Then
                    txtTransLog.Text = strdate & "พบปัญหาในการนำเข้าข้อมูล : Call sp_Initial_LUBE_Stock_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(strdate & "พบปัญหาในการนำเข้าข้อมูล : Call sp_Initial_LUBE_Stock_Inventory : " & ret_CallSPInitialLUBE)
                Else
                    txtTransLog.Text = strdate & "Call sp_Initial_LUBE_Stock_Inventory" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(strdate & "Call sp_Initial_LUBE_Stock_Inventory")
                End If


                'RunScriptSQL
                Dim script_path As String = Application.StartupPath & "\Scripts"
                If Directory.Exists(script_path) Then
                    For Each _file As String In Directory.GetFiles(script_path)
                        Dim file_name As String = System.IO.Path.GetFileName(_file)

                        If file_name.ToLower <> "sp_Initial_LUBE_Stock_Inventory.sql".ToLower Then
                            Dim ret_RunScriptSQL As String = RunScriptSQL(_file)
                            If ret_RunScriptSQL = "" Then
                                Application.DoEvents()
                                txtTransLog.Text = strdate & "Call " & file_name & vbCrLf & txtTransLog.Text
                                sw_rs.WriteLine(strdate & "Call " & file_name)
                                Threading.Thread.Sleep(100)
                            Else
                                Application.DoEvents()
                                txtTransLog.Text = strdate & "พบปัญหาในการนำเข้าข้อมูล : Call " & file_name & vbCrLf & txtTransLog.Text
                                sw_rs.WriteLine(strdate & "พบปัญหาในการนำเข้าข้อมูล : Call " & file_name & "      " & ret_RunScriptSQL)
                                Threading.Thread.Sleep(100)
                            End If
                        End If
                    Next
                End If

                'Update APP_Config
                If Not dt_appconfig Is Nothing AndAlso dt_appconfig.Rows.Count > 0 Then
                    Dim cnt As Integer = 0
                    For i As Integer = 0 To dt_appconfig.Rows.Count - 1
                        Dim config_key As String = dt_appconfig.Rows(i)("CONFIG_KEY").ToString
                        Dim config_value As String = dt_appconfig.Rows(i)("CONFIG_VALUE").ToString
                        If Update_APP_Config(config_key, config_value) > 0 Then
                            cnt += 1
                        End If
                    Next
                    Application.DoEvents()
                    txtTransLog.Text = strdate & "Update APP_CONFIG" & "  (" & cnt & " Row)" & vbCrLf & txtTransLog.Text
                    sw_rs.WriteLine(strdate & "Update APP_CONFIG" & "  (" & cnt & " Row)")
                    Threading.Thread.Sleep(100)
                End If


                txtTransLog.Text = strdate & "สิ้นสุดการนำเข้าข้อมูล" & vbCrLf & txtTransLog.Text
                sw_rs.WriteLine(strdate & "สิ้นสุดการนำเข้าข้อมูล")

                EnableButton(True)
            Else
                Using New Centered_MessageBox(Me)
                    MessageBox.Show("ไม่พบรายการสำหรับนำเข้าข้อมูล", "", MessageBoxButtons.OK)
                End Using
                EnableButton(True)
            End If

        Catch ex As Exception
            Application.DoEvents()
            txtTransLog.Text = strdate & "พบปัญหาในการนำเข้าข้อมูล " & ex.ToString & vbCrLf & txtTransLog.Text
            sw_rs.WriteLine(strdate & "พบปัญหาในการนำเข้าข้อมูล " & ex.ToString)
            Threading.Thread.Sleep(100)


            Using New Centered_MessageBox(Me)
                MessageBox.Show("ไม่สามารถนำเข้าข้อมูลได้ !" & vbCrLf & ex.ToString, "", MessageBoxButtons.OK)
            End Using
        End Try
        sw_rs.Close()
    End Sub


#Region "Connect Database"   'Connect Database
    'Dim Server As String = "10.195.2.177"
    'Dim Password As String = "pTT!CT01"
    'Public INIFile As String = Application.StartupPath & "\config.ini"
    Public ConnStr As String = getConnectionString()

    Function getConnectionString() As String

        Dim Server As String = "(local)"
        'Dim Server As String = "10.195.2.205"

        'Dim Database As String = "POSDB"
        Dim Database As String = "POSDBD"

        Dim Username As String = "sa"

        Dim Password As String = "1qaz@WSX"
        'Dim Password As String = "pTT!CT01"


        'Dim ini As New IniReader(INIFile)
        'ini.Section = "Setting"
        Return "Data Source=" & Server & ";Initial Catalog=" & Database & ";User ID=" & Username & ";Password=" & Password & ";Connect Timeout=1;"
    End Function


#End Region

#Region "Sub&Function"

    Function Update_APP_Config(CONFIG_KEY As String, CONFIG_VALUE As String) As Integer
        Dim sql As String = "SELECT * FROM APP_Config Where Config_Key = '" & CONFIG_KEY & "'"
        Dim da As New SqlDataAdapter(sql, ConnStr)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As Integer = 0
        If dt.Rows.Count > 0 Then
            sql = "Update APP_Config set CONFIG_VALUE = '" & CONFIG_VALUE & "' where Config_Key = '" & CONFIG_KEY & "' "
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = sql
                .CommandType = CommandType.Text
                .Connection = conn
            End With
            ret = cmd.ExecuteNonQuery()
            conn.Close()
        End If
        Return ret
    End Function

    Function GET_PUMP_ID(HOSE_ID As String, trans As SqlTransaction) As String
        Dim sql As String = "SELECT EH.Pump_ID FROM ENABLERDB.dbo.HOSES AS EH LEFT OUTER JOIN POSDB.dbo.TBMATERIAL AS PM ON EH.Grade_ID = PM.MAT_ID2 WHERE EH.HOSE_ID = " & HOSE_ID & ""
        Dim da As New SqlDataAdapter(sql, ConnStr)
        'da.SelectCommand.Transaction = trans
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As String = ""
        If dt.Rows.Count > 0 Then
            ret = dt.Rows(0)("Pump_ID").ToString
        End If
        Return ret
    End Function

    Function GET_MAT_ID(HOSE_ID As String, trans As SqlTransaction) As String
        Dim sql As String = "SELECT EH.Pump_ID, EH.Tank_ID, PM.MAT_ID FROM ENABLERDB.dbo.HOSES AS EH LEFT OUTER JOIN POSDB.dbo.TBMATERIAL AS PM ON EH.Grade_ID = PM.MAT_ID2 WHERE EH.HOSE_ID = " & HOSE_ID & ""
        Dim da As New SqlDataAdapter(sql, ConnStr)
        'da.SelectCommand.Transaction = trans
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As String = ""
        If dt.Rows.Count > 0 Then
            ret = dt.Rows(0)("MAT_ID").ToString
        End If
        Return ret
    End Function

    Function GET_TANK_ID(HOSE_ID As String, trans As SqlTransaction) As String
        Dim sql As String = "SELECT EH.Pump_ID, EH.Tank_ID, PM.MAT_ID FROM ENABLERDB.dbo.HOSES AS EH LEFT OUTER JOIN POSDB.dbo.TBMATERIAL AS PM ON EH.Grade_ID = PM.MAT_ID2 WHERE EH.HOSE_ID = " & HOSE_ID & ""
        Dim da As New SqlDataAdapter(sql, ConnStr)
        'da.SelectCommand.Transaction = trans
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As String = ""
        If dt.Rows.Count > 0 Then
            ret = dt.Rows(0)("Tank_ID").ToString
        End If
        Return ret
    End Function

    Function GET_MAT_ID_TANK(TANK_ID As String, trans As SqlTransaction) As String
        Dim sql As String = "SELECT PM.MAT_ID FROM ENABLERDB.dbo.Tanks AS ET LEFT OUTER JOIN POSDB.dbo.TBMATERIAL AS PM ON ET.Grade_ID = PM.MAT_ID2 WHERE ET.TANK_ID = " & TANK_ID & ""
        Dim da As New SqlDataAdapter(sql, ConnStr)
        'da.SelectCommand.Transaction = trans
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As String = ""
        If dt.Rows.Count > 0 Then
            ret = dt.Rows(0)("MAT_ID").ToString
        End If
        Return ret
    End Function

    Function GET_TANK_NAME(TANK_ID As String, trans As SqlTransaction) As String
        Dim sql As String = "SELECT ET.Tank_Name FROM ENABLERDB.dbo.Tanks AS ET LEFT OUTER JOIN POSDB.dbo.TBMATERIAL AS PM ON ET.Grade_ID = PM.MAT_ID2 WHERE ET.TANK_ID =" & TANK_ID & ""
        Dim da As New SqlDataAdapter(sql, ConnStr)
        'da.SelectCommand.Transaction = trans
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As String = ""
        If dt.Rows.Count > 0 Then
            ret = dt.Rows(0)("Tank_Name").ToString
        End If
        Return ret
    End Function

    Function GET_TANK_NUMBER(TANK_ID As String, trans As SqlTransaction) As String
        Dim sql As String = "SELECT ET.Tank_Number FROM ENABLERDB.dbo.Tanks AS ET LEFT OUTER JOIN POSDB.dbo.TBMATERIAL AS PM ON ET.Grade_ID = PM.MAT_ID2 WHERE ET.TANK_ID = " & TANK_ID & ""
        Dim da As New SqlDataAdapter(sql, ConnStr)
        'da.SelectCommand.Transaction = trans
        Dim dt As New DataTable
        da.Fill(dt)

        Dim ret As String = ""
        If dt.Rows.Count > 0 Then
            ret = dt.Rows(0)("Tank_Number").ToString
        End If
        Return ret
    End Function

    Sub EnableButton(IsEnable As Boolean)

        pbExport.Enabled = IsEnable
        pbImport.Enabled = IsEnable
        If IsEnable Then
            pbExport.BackgroundImage = My.Resources.export_th__2_
            pbImport.BackgroundImage = My.Resources.Import_th
        Else
            pbExport.BackgroundImage = My.Resources.export_th_dis
            pbImport.BackgroundImage = My.Resources.import_th_dis
        End If

    End Sub

    Sub writeLogResult(strResult As String)
        Try
            Dim strFileResule As String = Application.StartupPath & "\" & "_Result_" & DateTime.Now.ToString("yyyyMMdd") & ".txt"
            Dim strdate As String = DateTime.Now.ToString("yyyyMMdd hh:MM:ss") & "  "
            Dim sw_rs As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(strFileResule, True)
            sw_rs.WriteLine(strdate & strResult)
            sw_rs.Close()
        Catch ex As Exception

        End Try

    End Sub

    Function CallSPImportProduct() As Boolean
        Try
            Dim sql As String = "sp_Import_Product_To_Inventory"
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = sql
                .CommandType = CommandType.StoredProcedure
                .Connection = conn
                .ExecuteNonQuery()
            End With
            conn.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Function UpdateDefaultProduct() As Integer
        Try
            Dim ret As Integer
            Dim sql As String = "Update PRODUCTS set ISSHOWINPOS = 0 , ISRECOMMEND=0 "
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = sql
                .CommandType = CommandType.Text
                .Connection = conn
            End With
            ret = cmd.ExecuteNonQuery()
            conn.Close()
            Return ret
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Function UpdateISSHOWINPOS() As Integer
        Try
            Dim ret As Integer
            Dim sql As String = "Update PRODUCTS set ISSHOWINPOS = 1 where ProductCode in (select MAT_ID from TBMATTERIAL_SITE)"
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = sql
                .CommandType = CommandType.Text
                .Connection = conn
            End With
            ret = cmd.ExecuteNonQuery()
            conn.Close()
            Return ret
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Function UpdateISRECOMMEND() As Integer
        Try
            Dim ret As Integer
            Dim sql As String = "Update PRODUCTS set ISRECOMMEND = 1 where ProductCode in (select MAT_ID from TBMAT_RECOMMEND)"
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = sql
                .CommandType = CommandType.Text
                .Connection = conn

            End With
            ret = cmd.ExecuteNonQuery()
            conn.Close()
            Return ret
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Function CallSPInitialLUBE() As String
        Try
            Dim sql As String = "sp_Initial_LUBE_Stock_Inventory"
            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = sql
                .CommandType = CommandType.StoredProcedure
                .Connection = conn
                .ExecuteNonQuery()
            End With
            conn.Close()
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function

    Function DeleteData(TableName As String) As String
        Dim sql As String = "TRUNCATE TABLE " & TableName
        Dim conn As New SqlConnection(ConnStr)
        conn.Open()
        Dim cmd As New SqlCommand
        With cmd
            .CommandText = sql
            .CommandType = CommandType.Text
            .Connection = conn
            .ExecuteNonQuery()
        End With
        conn.Close()
        Return ""
    End Function

    Function InsertData(TableName As String, dt As DataTable) As String
        Dim conn As New SqlConnection(ConnStr)
        conn.Open()
        Dim trans As SqlTransaction
        trans = conn.BeginTransaction

        Dim modby As String = "IEReplaceCOCO"
        Dim sql As String = ""
        Try

            Select Case TableName
                Case "APP_DATA"
#Region "APP_DATA"
                    Dim AID, DEPOT, SITENAME, SITEADD, SITETEL, SITEFAX, SITEZIPCODE, VATNO, LOCALVAT, IS_COCO, IS_POSCASH, BUS_PLACE, VOLUMEFORMAT, VALUEFORMAT, SITENAME2, C_SERVICE, COM_NAME, COM_BRANCH, LOCAL_DIFFERENCE, TOBACCO_TAX As String


                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandText = sql
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        With dt
                            AID = IIf(.Rows(i)("AID").ToString = "", "null", "'" & .Rows(i)("AID").ToString & "'")
                            DEPOT = IIf(.Rows(i)("DEPOT").ToString = "", "null", "'" & .Rows(i)("DEPOT").ToString & "'")
                            SITENAME = IIf(.Rows(i)("SITENAME").ToString = "", "null", "'" & .Rows(i)("SITENAME").ToString & "'")
                            SITEADD = IIf(.Rows(i)("SITEADD").ToString = "", "null", "'" & .Rows(i)("SITEADD").ToString & "'")
                            SITETEL = IIf(.Rows(i)("SITETEL").ToString = "", "null", "'" & .Rows(i)("SITETEL").ToString & "'")
                            SITEFAX = IIf(.Rows(i)("SITEFAX").ToString = "", "null", "'" & .Rows(i)("SITEFAX").ToString & "'")
                            SITEZIPCODE = IIf(.Rows(i)("SITEZIPCODE").ToString = "", "null", "'" & .Rows(i)("SITEZIPCODE").ToString & "'")
                            VATNO = IIf(.Rows(i)("VATNO").ToString = "", "null", "'" & .Rows(i)("VATNO").ToString & "'")
                            LOCALVAT = IIf(.Rows(i)("LOCALVAT").ToString = "", "null", "'" & .Rows(i)("LOCALVAT").ToString & "'")
                            IS_COCO = IIf(.Rows(i)("IS_COCO").ToString = "", "null", "'" & .Rows(i)("IS_COCO").ToString & "'")
                            IS_POSCASH = IIf(.Rows(i)("IS_POSCASH").ToString = "", "null", "'" & .Rows(i)("IS_POSCASH").ToString & "'")
                            BUS_PLACE = IIf(.Rows(i)("BUS_PLACE").ToString = "", "null", "'" & .Rows(i)("BUS_PLACE").ToString & "'")
                            VOLUMEFORMAT = IIf(.Rows(i)("VOLUMEFORMAT").ToString = "", "null", "'" & .Rows(i)("VOLUMEFORMAT").ToString & "'")
                            VALUEFORMAT = IIf(.Rows(i)("VALUEFORMAT").ToString = "", "null", "'" & .Rows(i)("VALUEFORMAT").ToString & "'")
                            SITENAME2 = IIf(.Rows(i)("SITENAME2").ToString = "", "null", "'" & .Rows(i)("SITENAME2").ToString & "'")
                            C_SERVICE = IIf(.Rows(i)("C_SERVICE").ToString = "", "null", "'" & .Rows(i)("C_SERVICE").ToString & "'")
                            COM_NAME = "null" 'IIf(.Rows(i)("COM_NAME").ToString = "", "null", "'" & .Rows(i)("COM_NAME").ToString & "'")
                            COM_BRANCH = "null" 'IIf(.Rows(i)("COM_BRANCH").ToString = "", "null", "'" & .Rows(i)("COM_BRANCH").ToString & "'")
                            LOCAL_DIFFERENCE = "0" 'IIf(.Rows(i)("LOCAL_DIFFERENCE").ToString = "", "null", "'" & .Rows(i)("LOCAL_DIFFERENCE").ToString & "'")
                            TOBACCO_TAX = "7" 'IIf(.Rows(i)("TOBACCO_TAX").ToString = "", "null", "'" & .Rows(i)("TOBACCO_TAX").ToString & "'")
                        End With


                        sql = "INSERT INTO [dbo].[APP_DATA]
                           ([AID]
                           ,[DEPOT]
                           ,[SITENAME]
                           ,[SITEADD]
                           ,[SITETEL]
                           ,[SITEFAX]
                           ,[SITEZIPCODE]
                           ,[VATNO]
                           ,[LOCALVAT]
                           ,[IS_COCO]
                           ,[IS_POSCASH]
                           ,[CREATEDATE]
                           ,[MODDATE]
                           ,[MODBY]
                           ,[BUS_PLACE]
                           ,[VOLUMEFORMAT]
                           ,[VALUEFORMAT]
                           ,[SITENAME2]
                           ,[C_SERVICE]
                           ,[COM_NAME]
                           ,[COM_BRANCH]
                           ,[LOCAL_DIFFERENCE]
                           ,[TOBACCO_TAX])
                            VALUES
                           (" & AID & "
                           ," & DEPOT & "  
                           ," & SITENAME & "  
                           ," & SITEADD & "  
                           ," & SITETEL & " 
                           ," & SITEFAX & "  
                           ," & SITEZIPCODE & "  
                           ," & VATNO & "  
                           ," & LOCALVAT & "  
                           ," & IS_COCO & "  
                           ," & IS_POSCASH & "  
                           ,getdate()
                           ,getdate()
                           ,'" & modby & "'
                           ," & BUS_PLACE & "  
                           ," & VOLUMEFORMAT & "  
                           ," & VALUEFORMAT & "  
                           ," & SITENAME2 & "  
                           ," & C_SERVICE & "  
                           ," & COM_NAME & "  
                           ," & COM_BRANCH & "  
                           ," & LOCAL_DIFFERENCE & "  
                           ," & TOBACCO_TAX & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "TBPOS_PUMP_ALLOW"
#Region "TBPOS_PUMP_ALLOW"
                    Dim POS_ID, PUMP_ID As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandText = sql
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        POS_ID = dt.Rows(i)("POS_ID").ToString
                        PUMP_ID = dt.Rows(i)("PUMP_ID").ToString

                        sql = "INSERT INTO [dbo].[TBPOS_PUMP_ALLOW]
                           ([POS_ID],[PUMP_ID],CREATEDATE,MODDATE,MODBY)
                           VALUES(" & IIf(POS_ID = "", "null", "'" & POS_ID & "'") & "  
                           ," & IIf(PUMP_ID = "", "null", "'" & PUMP_ID & "'") & "  
                           ,getdate(),getdate(),'" & modby & "')"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "TBMATERIAL"
#Region "TBMATERIAL"
                    Dim MAT_ID, MAT_NAME, MAT_ID2, MAT_NAME2, MAT_NAME3, MAT_BARCODE, QTY, UOM, MOVING_AVG_PRICE, STOCK, STOCK_MIN, STOCK_MAX, STOCK_LOCATION_ID, TAX_CLASS, MAT_GROUP, MAT_GROUP3,
                                          DIVISION_ID, PRICE0, PRICE1, PRICE2, PRICE3, PRICE4, PRICE5, PRICE6, PRICE7, PRICE8, PRICE9, PRICE10, PRICE11, PRICE12, TIMEOFSALE, LAST_SALE, LAST_RECEIVE, BLOCK,
                                          PRICINGDATE, PRICINGMODBY, LOCATION_ID, MATCOLOR, OBJ_ID, OBJ_ID_MAT_GROUP3, OBJ_ID_DIVISION_ID As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        With dt
                            MAT_ID = IIf(.Rows(i)("MAT_ID").ToString = "", "null", "'" & .Rows(i)("MAT_ID").ToString & "'")
                            MAT_NAME = IIf(.Rows(i)("MAT_NAME").ToString = "", "null", "'" & .Rows(i)("MAT_NAME").ToString & "'")
                            MAT_ID2 = IIf(.Rows(i)("MAT_ID2").ToString = "", "null", "'" & .Rows(i)("MAT_ID2").ToString & "'")
                            MAT_NAME2 = IIf(.Rows(i)("MAT_NAME2").ToString = "", "null", "'" & .Rows(i)("MAT_NAME2").ToString & "'")
                            MAT_NAME3 = IIf(.Rows(i)("MAT_NAME3").ToString = "", "null", "'" & .Rows(i)("MAT_NAME3").ToString & "'")
                            MAT_BARCODE = IIf(.Rows(i)("MAT_BARCODE").ToString = "", "null", "'" & .Rows(i)("MAT_BARCODE").ToString & "'")
                            QTY = IIf(.Rows(i)("QTY").ToString = "", "null", "'" & .Rows(i)("QTY").ToString & "'")
                            UOM = IIf(.Rows(i)("UOM").ToString = "", "null", "'" & .Rows(i)("UOM").ToString & "'")
                            MOVING_AVG_PRICE = IIf(.Rows(i)("MOVING_AVG_PRICE").ToString = "", "null", "'" & .Rows(i)("MOVING_AVG_PRICE").ToString & "'")
                            STOCK = IIf(.Rows(i)("STOCK").ToString = "", "null", "'" & .Rows(i)("STOCK").ToString & "'")
                            STOCK_MIN = IIf(.Rows(i)("STOCK_MIN").ToString = "", "null", "'" & .Rows(i)("STOCK_MIN").ToString & "'")
                            STOCK_MAX = IIf(.Rows(i)("STOCK_MAX").ToString = "", "null", "'" & .Rows(i)("STOCK_MAX").ToString & "'")
                            STOCK_LOCATION_ID = IIf(.Rows(i)("STOCK_LOCATION_ID").ToString = "", "null", "'" & .Rows(i)("STOCK_LOCATION_ID").ToString & "'")
                            TAX_CLASS = IIf(.Rows(i)("TAX_CLASS").ToString = "", "null", "'" & .Rows(i)("TAX_CLASS").ToString & "'")
                            MAT_GROUP = IIf(.Rows(i)("MAT_GROUP").ToString = "", "null", "'" & .Rows(i)("MAT_GROUP").ToString & "'")
                            MAT_GROUP3 = IIf(.Rows(i)("MAT_GROUP3").ToString = "", "null", "'" & .Rows(i)("MAT_GROUP3").ToString & "'")
                            DIVISION_ID = IIf(.Rows(i)("DIVISION_ID").ToString = "DIVISION_ID", "null", "'" & .Rows(i)("DIVISION_ID").ToString & "'")
                            PRICE0 = IIf(.Rows(i)("PRICE0").ToString = "", "null", "'" & .Rows(i)("PRICE0").ToString & "'")
                            PRICE1 = IIf(.Rows(i)("PRICE1").ToString = "", "null", "'" & .Rows(i)("PRICE1").ToString & "'")
                            PRICE2 = IIf(.Rows(i)("PRICE2").ToString = "", "null", "'" & .Rows(i)("PRICE2").ToString & "'")
                            PRICE3 = IIf(.Rows(i)("PRICE3").ToString = "", "null", "'" & .Rows(i)("PRICE3").ToString & "'")
                            PRICE4 = IIf(.Rows(i)("PRICE4").ToString = "", "null", "'" & .Rows(i)("PRICE4").ToString & "'")
                            PRICE5 = IIf(.Rows(i)("PRICE5").ToString = "", "null", "'" & .Rows(i)("PRICE5").ToString & "'")
                            PRICE6 = IIf(.Rows(i)("PRICE6").ToString = "", "null", "'" & .Rows(i)("PRICE6").ToString & "'")
                            PRICE7 = IIf(.Rows(i)("PRICE7").ToString = "", "null", "'" & .Rows(i)("PRICE7").ToString & "'")
                            PRICE8 = IIf(.Rows(i)("PRICE8").ToString = "", "null", "'" & .Rows(i)("PRICE8").ToString & "'")
                            PRICE9 = IIf(.Rows(i)("PRICE9").ToString = "", "null", "'" & .Rows(i)("PRICE9").ToString & "'")
                            PRICE10 = IIf(.Rows(i)("PRICE10").ToString = "", "null", "'" & .Rows(i)("PRICE10").ToString & "'")
                            PRICE11 = IIf(.Rows(i)("PRICE11").ToString = "", "null", "'" & .Rows(i)("PRICE11").ToString & "'")
                            PRICE12 = IIf(.Rows(i)("PRICE12").ToString = "", "null", "'" & .Rows(i)("PRICE12").ToString & "'")
                            TIMEOFSALE = IIf(.Rows(i)("TIMEOFSALE").ToString = "", "null", "'" & .Rows(i)("TIMEOFSALE").ToString & "'")

                            LAST_SALE = IIf(.Rows(i)("LAST_SALE").ToString = "", "null", "" & .Rows(i)("LAST_SALE").ToString & "")
                            If LAST_SALE <> "null" Then
                                LAST_SALE = ConvertDate(LAST_SALE)
                            End If

                            LAST_RECEIVE = IIf(.Rows(i)("LAST_RECEIVE").ToString = "", "null", "" & .Rows(i)("LAST_RECEIVE").ToString & "")
                            If LAST_RECEIVE <> "null" Then
                                LAST_RECEIVE = ConvertDate(LAST_RECEIVE)
                            End If

                            BLOCK = IIf(.Rows(i)("BLOCK").ToString = "", "null", "'" & .Rows(i)("BLOCK").ToString & "'")
                            PRICINGDATE = IIf(.Rows(i)("PRICINGDATE").ToString = "", "null", "" & .Rows(i)("PRICINGDATE").ToString & "")
                            If PRICINGDATE <> "null" Then
                                PRICINGDATE = ConvertDate(PRICINGDATE)
                            End If

                            PRICINGMODBY = IIf(.Rows(i)("PRICINGMODBY").ToString = "", "null", "'" & .Rows(i)("PRICINGMODBY").ToString & "'")
                            LOCATION_ID = IIf(.Rows(i)("LOCATION_ID").ToString = "", "null", "'" & .Rows(i)("LOCATION_ID").ToString & "'")
                            MATCOLOR = IIf(.Rows(i)("MATCOLOR").ToString = "", "null", "'" & .Rows(i)("MATCOLOR").ToString & "'")
                            OBJ_ID = "null" 'IIf(.Rows(i)("OBJ_ID").ToString = "", "null", "'" & .Rows(i)("OBJ_ID").ToString & "'")
                            OBJ_ID_MAT_GROUP3 = "null" 'IIf(.Rows(i)("OBJ_ID_MAT_GROUP3").ToString = "", "null", "'" & .Rows(i)("OBJ_ID_MAT_GROUP3").ToString & "'")
                            OBJ_ID_DIVISION_ID = "null" 'IIf(.Rows(i)("OBJ_ID_DIVISION_ID").ToString = "", "null", "'" & .Rows(i)("OBJ_ID_DIVISION_ID").ToString & "'")
                        End With


                        sql = "INSERT INTO [dbo].[TBMATERIAL]
                           ([MAT_ID]
                           ,[MAT_NAME]
                           ,[MAT_ID2]
                           ,[MAT_NAME2]
                           ,[MAT_NAME3]
                           ,[MAT_BARCODE]
                           ,[QTY]
                           ,[UOM]
                           ,[MOVING_AVG_PRICE]
                           ,[STOCK]
                           ,[STOCK_MIN]
                           ,[STOCK_MAX]
                           ,[STOCK_LOCATION_ID]
                           ,[TAX_CLASS]
                           ,[MAT_GROUP]
                           ,[MAT_GROUP3]
                           ,[DIVISION_ID]
                           ,[PRICE0]
                           ,[PRICE1]
                           ,[PRICE2]
                           ,[PRICE3]
                           ,[PRICE4]
                           ,[PRICE5]
                           ,[PRICE6]
                           ,[PRICE7]
                           ,[PRICE8]
                           ,[PRICE9]
                           ,[PRICE10]
                           ,[PRICE11]
                           ,[PRICE12]
                           ,[TIMEOFSALE]
                           ,[LAST_SALE]
                           ,[LAST_RECEIVE]
                           ,[BLOCK]
                           ,[PRICINGDATE]
                           ,[PRICINGMODBY]
                           ,[LOCATION_ID]
                           ,[MATCOLOR]
                           ,[CREATEDATE]
                           ,[MODDATE]
                           ,[MODBY]
                           ,[OBJ_ID]
                           ,[OBJ_ID_MAT_GROUP3]
                           ,[OBJ_ID_DIVISION_ID])
                            VALUES
                           (" & MAT_ID & "
                           ," & MAT_NAME & "
                           ," & MAT_ID2 & "
                           ," & MAT_NAME2 & "
                           ," & MAT_NAME3 & "
                           ," & MAT_BARCODE & "
                           ," & QTY & "
                           ," & UOM & "
                           ," & MOVING_AVG_PRICE & "
                           ," & STOCK & "
                           ," & STOCK_MIN & "
                           ," & STOCK_MAX & "
                           ," & STOCK_LOCATION_ID & "
                           ," & TAX_CLASS & "
                           ," & MAT_GROUP & "
                           ," & MAT_GROUP3 & "
                           ," & DIVISION_ID & "
                           ," & PRICE0 & "
                           ," & PRICE1 & "
                           ," & PRICE2 & "
                           ," & PRICE3 & "
                           ," & PRICE4 & "
                           ," & PRICE5 & "
                           ," & PRICE6 & "
                           ," & PRICE7 & "
                           ," & PRICE8 & "
                           ," & PRICE9 & "
                           ," & PRICE10 & "
                           ," & PRICE11 & "
                           ," & PRICE12 & "
                           ," & TIMEOFSALE & "
                           ," & LAST_SALE & "
                           ," & LAST_RECEIVE & "
                           ," & BLOCK & "
                           ," & PRICINGDATE & "
                           ," & PRICINGMODBY & "
                           ," & LOCATION_ID & "
                           ," & MATCOLOR & "
                           ,getdate()
                           ,getdate()
                           ,'" & modby & "'
                           ," & OBJ_ID & "
                           ," & OBJ_ID_MAT_GROUP3 & "
                           ," & OBJ_ID_DIVISION_ID & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBBOM_USAGE"
#Region "TBBOM_USAGE"
                    Dim MAT_ID, BASE_QTY, BASE_UOM, COMPONENT, QTY, UOM, BLOCK, DAY_ID, SHIFT_NO, ALT, BOM_USG As String
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        With dt
                            MAT_ID = IIf(.Rows(i)("MAT_ID").ToString = "", "null", "'" & .Rows(i)("MAT_ID").ToString & "'")
                            BASE_QTY = IIf(.Rows(i)("BASE_QTY").ToString = "", "null", "'" & .Rows(i)("BASE_QTY").ToString & "'")
                            BASE_UOM = IIf(.Rows(i)("BASE_UOM").ToString = "", "null", "'" & .Rows(i)("BASE_UOM").ToString & "'")
                            COMPONENT = IIf(.Rows(i)("COMPONENT").ToString = "", "null", "'" & .Rows(i)("COMPONENT").ToString & "'")
                            QTY = IIf(.Rows(i)("QTY").ToString = "", "null", "'" & .Rows(i)("QTY").ToString & "'")
                            UOM = IIf(.Rows(i)("UOM").ToString = "", "null", "'" & .Rows(i)("UOM").ToString & "'")
                            BLOCK = IIf(.Rows(i)("BLOCK").ToString = "", "null", "'" & .Rows(i)("BLOCK").ToString & "'")
                            DAY_ID = IIf(.Rows(i)("DAY_ID").ToString = "", "null", "'" & .Rows(i)("DAY_ID").ToString & "'")
                            SHIFT_NO = IIf(.Rows(i)("SHIFT_NO").ToString = "", "null", "'" & .Rows(i)("SHIFT_NO").ToString & "'")
                            ALT = IIf(.Rows(i)("ALT").ToString = "", "null", "'" & .Rows(i)("ALT").ToString & "'")
                            BOM_USG = IIf(.Rows(i)("BOM_USG").ToString = "", "null", "'" & .Rows(i)("BOM_USG").ToString & "'")
                        End With

                        sql = "INSERT INTO [dbo].[TBBOM_USAGE]
                           ([MAT_ID]
                           ,[BASE_QTY]
                           ,[BASE_UOM]
                           ,[COMPONENT]
                           ,[QTY]
                           ,[UOM]
                           ,[BLOCK]
                           ,[DAY_ID]
                           ,[SHIFT_NO]
                           ,[ALT]
                           ,[BOM_USG]
                           ,[CREATEDATE]
                           ,[MODDATE]
                           ,[MODBY])
                            VALUES
                           (" & MAT_ID & "
                           ," & BASE_QTY & "
                           ," & BASE_UOM & "
                           ," & COMPONENT & "
                           ," & QTY & "
                           ," & UOM & "
                           ," & BLOCK & "
                           ," & DAY_ID & "
                           ," & SHIFT_NO & "
                           ," & ALT & "
                           ," & BOM_USG & "
                           ,getdate()
                           ,getdate()
                           ,'" & modby & "')"


                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "TBMATTERIAL_SITE"
#Region "TBMATTERIAL_SITE"
                    Dim MAT_ID As String
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        MAT_ID = IIf(dt.Rows(i)("MAT_ID").ToString = "", "null", dt.Rows(i)("MAT_ID").ToString)
                        sql = "INSERT INTO [dbo].[TBMATTERIAL_SITE]
                           ([MAT_ID]
                           ,[CREATEDATE]
                           ,[MODDATE]
                           ,[MODBY])
                            VALUES
                           (" & MAT_ID & "
                           ,getdate()
                           ,getdate()
                           ,'" & modby & "')"


                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "TBCONVERSION"
#Region "TBCONVERSION"
                    Dim MAT_ID, BASE_UOM, ALTERNATIVE_UOM, NUMERATOR, DENOMINATOR, NORMAL_SIZE_L, BASE_UOM_BLOCK, DAY_ID, SHIFT_NO As String
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        With dt
                            MAT_ID = IIf(.Rows(i)("MAT_ID").ToString = "", "null", "'" & .Rows(i)("MAT_ID").ToString & "'")
                            BASE_UOM = IIf(.Rows(i)("BASE_UOM").ToString = "", "null", "'" & .Rows(i)("BASE_UOM").ToString & "'")
                            ALTERNATIVE_UOM = IIf(.Rows(i)("ALTERNATIVE_UOM").ToString = "", "null", "'" & .Rows(i)("ALTERNATIVE_UOM").ToString & "'")
                            NUMERATOR = IIf(.Rows(i)("NUMERATOR").ToString = "", "null", "" & .Rows(i)("NUMERATOR").ToString & "")
                            DENOMINATOR = IIf(.Rows(i)("DENOMINATOR").ToString = "", "null", "" & .Rows(i)("DENOMINATOR").ToString & "")
                            NORMAL_SIZE_L = IIf(.Rows(i)("NORMAL_SIZE_L").ToString = "", "null", "'" & .Rows(i)("NORMAL_SIZE_L").ToString & "'")
                            BASE_UOM_BLOCK = IIf(.Rows(i)("BASE_UOM_BLOCK").ToString = "", "null", "'" & .Rows(i)("BASE_UOM_BLOCK").ToString & "'")
                            DAY_ID = IIf(.Rows(i)("DAY_ID").ToString = "", "null", "" & .Rows(i)("DAY_ID").ToString & "")
                            SHIFT_NO = IIf(.Rows(i)("SHIFT_NO").ToString = "", "null", "" & .Rows(i)("SHIFT_NO").ToString & "")
                        End With
                        sql = "INSERT INTO [dbo].[TBCONVERSION]
                           ([MAT_ID]
                           ,[BASE_UOM]
                           ,[ALTERNATIVE_UOM]
                           ,[NUMERATOR]
                           ,[DENOMINATOR]
                           ,[NORMAL_SIZE_L]
                           ,[BASE_UOM_BLOCK]
                           ,[DAY_ID]
                           ,[SHIFT_NO]
                           ,[CREATEDATE]
                           ,[MODDATE]
                           ,[MODBY])
                            VALUES
                           (" & MAT_ID & "
                           ," & BASE_UOM & "
                           ," & ALTERNATIVE_UOM & "
                           ," & NUMERATOR & "
                           ," & DENOMINATOR & "
                           ," & NORMAL_SIZE_L & "
                           ," & BASE_UOM_BLOCK & "
                           ," & DAY_ID & "
                           ," & SHIFT_NO & "
                           ,getdate()
                           ,getdate()
                           ,'" & modby & "')"


                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "LKMAT_GROUP3"
#Region "LKMAT_GROUP3"
                    Dim MAT_GROUP3, GROUP3_NAME As String
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        MAT_GROUP3 = IIf(dt.Rows(i)("MAT_GROUP3").ToString = "", "null", "'" & dt.Rows(i)("MAT_GROUP3").ToString & "'")
                        GROUP3_NAME = IIf(dt.Rows(i)("GROUP3_NAME").ToString = "", "null", "'" & dt.Rows(i)("GROUP3_NAME").ToString & "'")

                        sql = "INSERT INTO [dbo].[LKMAT_GROUP3]
                           ([MAT_GROUP3]
                           ,[GROUP3_NAME]
                           ,[CREATEDATE]
                           ,[MODDATE]
                           ,[MODBY])
                            VALUES
                           (" & MAT_GROUP3 & "
                           ," & GROUP3_NAME & "
                           ,getdate()
                           ,getdate()
                           ,'" & modby & "')"


                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "LKDIVISION"
#Region "LKDIVISION"
                    Dim DIVISION_ID, DIVISION_NAME, CAN_RETURN As String
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1

                        DIVISION_ID = IIf(dt.Rows(i)("DIVISION_ID").ToString = "", "null", "'" & dt.Rows(i)("DIVISION_ID").ToString & "'")
                        DIVISION_NAME = IIf(dt.Rows(i)("DIVISION_NAME").ToString = "", "null", "'" & dt.Rows(i)("DIVISION_NAME").ToString & "'")
                        CAN_RETURN = IIf(dt.Rows(i)("CAN_RETURN").ToString = "", "null", "'" & dt.Rows(i)("CAN_RETURN").ToString & "'")

                        sql = "INSERT INTO [dbo].[LKDIVISION]
                           ([DIVISION_ID]
                           ,[DIVISION_NAME]
                           ,[CAN_RETURN]
                           ,[CREATEDATE]
                           ,[MODDATE]
                           ,[MODBY])
                            VALUES
                           (" & DIVISION_ID & "
                           ," & DIVISION_NAME & "
                           ," & CAN_RETURN & "
                           ,getdate()
                           ,getdate()
                           ,'" & modby & "')"


                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()

                    Next
#End Region

                Case "TBMAT_RECOMMEND"
#Region "TBMAT_RECOMMEND"
                    Dim MAT_ID As String
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        MAT_ID = IIf(dt.Rows(i)("MAT_ID").ToString = "", "null", "'" & dt.Rows(i)("MAT_ID").ToString & "'")

                        sql = "INSERT INTO [dbo].[TBMAT_RECOMMEND]
                           ([MAT_ID])
                            VALUES
                           (" & MAT_ID & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "APP_CONFIG"
#Region "APP_CONFIG"
                    'Dim CONFIG_KEY, CONFIG_DESC, CONFIG_TYPE, MODULE_TYPE, CONFIG_VALUE, ADMIN_ONLY As String
                    'Dim cmd As New SqlCommand
                    'With cmd
                    '    .CommandType = CommandType.Text
                    '    .Connection = trans.Connection
                    '    .Transaction = trans
                    'End With
                    'For i As Integer = 0 To dt.Rows.Count - 1
                    '    CONFIG_KEY = IIf(dt.Rows(i)("CONFIG_KEY").ToString = "", "null", "'" & dt.Rows(i)("CONFIG_KEY").ToString & "'")
                    '    CONFIG_DESC = IIf(dt.Rows(i)("CONFIG_DESC").ToString = "", "null", "'" & dt.Rows(i)("CONFIG_DESC").ToString & "'")
                    '    CONFIG_TYPE = IIf(dt.Rows(i)("CONFIG_TYPE").ToString = "", "null", "'" & dt.Rows(i)("CONFIG_TYPE").ToString & "'")
                    '    MODULE_TYPE = IIf(dt.Rows(i)("MODULE_TYPE").ToString = "", "null", "'" & dt.Rows(i)("MODULE_TYPE").ToString & "'")
                    '    CONFIG_VALUE = IIf(dt.Rows(i)("CONFIG_VALUE").ToString = "", "null", "'" & dt.Rows(i)("CONFIG_VALUE").ToString & "'")
                    '    ADMIN_ONLY = "0" ' IIf(dt.Rows(i)("ADMIN_ONLY").ToString = "", "null", "'" & dt.Rows(i)("ADMIN_ONLY").ToString & "'")

                    '    sql = "INSERT INTO [dbo].[APP_CONFIG]
                    '        ([CONFIG_KEY]
                    '        ,[CONFIG_DESC]
                    '        ,[CONFIG_TYPE]
                    '        ,[MODULE_TYPE]
                    '        ,[CONFIG_VALUE]
                    '        ,[CREATEDATE]
                    '        ,[MODDATE]
                    '        ,[MODBY]
                    '        ,[ADMIN_ONLY])
                    '    VALUES
                    '        (" & CONFIG_KEY & "
                    '        ," & CONFIG_DESC & "
                    '        ," & CONFIG_TYPE & "
                    '        ," & MODULE_TYPE & "
                    '        ," & CONFIG_VALUE & "
                    '        ,getdate()
                    '        ,getdate()
                    '        ,'" & modby & "'
                    '        ," & ADMIN_ONLY & ")"

                    '    cmd.CommandText = sql
                    '    cmd.ExecuteNonQuery()
                    'Next
#End Region

                Case "POS_CONFIG"
#Region "POS_CONFIG"
                    Dim POS_ID, MODULE_TYPE, POS_NO, POS_IP, TERMINAL_ID, RDCODE, ISMAIN, MAXMONEY, EDC1_PORTNAME, EDC2_PORTNAME As String
                    Dim EDC_SPEED, EDC_PARITY, EDC_TIMEOUT, CLEAR_SALE_INFO, SHOW_PUMP_INFO, CASHDRAWERPRINT, AUTOPRINT As String
                    Dim EDC_MODEL, L_MID, L_TID As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        POS_ID = IIf(dt.Rows(i)("POS_ID").ToString = "", "null", "'" & dt.Rows(i)("POS_ID").ToString & "'")
                        MODULE_TYPE = IIf(dt.Rows(i)("MODULE_TYPE").ToString = "", "null", "'" & dt.Rows(i)("MODULE_TYPE").ToString & "'")
                        POS_NO = IIf(dt.Rows(i)("POS_NO").ToString = "", "null", "'" & dt.Rows(i)("POS_NO").ToString & "'")
                        POS_IP = IIf(dt.Rows(i)("POS_IP").ToString = "", "null", "'" & dt.Rows(i)("POS_IP").ToString & "'")
                        TERMINAL_ID = IIf(dt.Rows(i)("TERMINAL_ID").ToString = "", "null", "'" & dt.Rows(i)("TERMINAL_ID").ToString & "'")
                        RDCODE = IIf(dt.Rows(i)("RDCODE").ToString = "", "null", "'" & dt.Rows(i)("RDCODE").ToString & "'")
                        ISMAIN = IIf(dt.Rows(i)("ISMAIN").ToString = "", "null", "'" & dt.Rows(i)("ISMAIN").ToString & "'")
                        MAXMONEY = IIf(dt.Rows(i)("MAXMONEY").ToString = "", "null", "'" & dt.Rows(i)("MAXMONEY").ToString & "'")
                        EDC1_PORTNAME = IIf(dt.Rows(i)("EDC1_PORTNAME").ToString = "", "null", "'" & dt.Rows(i)("EDC1_PORTNAME").ToString & "'")
                        EDC2_PORTNAME = IIf(dt.Rows(i)("EDC2_PORTNAME").ToString = "", "null", "'" & dt.Rows(i)("EDC2_PORTNAME").ToString & "'")
                        EDC_SPEED = IIf(dt.Rows(i)("EDC_SPEED").ToString = "", "null", "'" & dt.Rows(i)("EDC_SPEED").ToString & "'")
                        EDC_PARITY = IIf(dt.Rows(i)("EDC_PARITY").ToString = "", "null", "'" & dt.Rows(i)("EDC_PARITY").ToString & "'")
                        EDC_TIMEOUT = IIf(dt.Rows(i)("EDC_TIMEOUT").ToString = "", "null", "'" & dt.Rows(i)("EDC_TIMEOUT").ToString & "'")
                        CLEAR_SALE_INFO = IIf(dt.Rows(i)("CLEAR_SALE_INFO").ToString = "", "null", "'" & dt.Rows(i)("CLEAR_SALE_INFO").ToString & "'")
                        SHOW_PUMP_INFO = IIf(dt.Rows(i)("SHOW_PUMP_INFO").ToString = "", "null", "'" & dt.Rows(i)("SHOW_PUMP_INFO").ToString & "'")
                        CASHDRAWERPRINT = IIf(dt.Rows(i)("CASHDRAWERPRINT").ToString = "", "null", "'" & dt.Rows(i)("CASHDRAWERPRINT").ToString & "'")
                        AUTOPRINT = IIf(dt.Rows(i)("AUTOPRINT").ToString = "", "null", "'" & dt.Rows(i)("AUTOPRINT").ToString & "'")
                        EDC_MODEL = IIf(dt.Rows(i)("EDC_MODEL").ToString = "", "null", "'" & dt.Rows(i)("EDC_MODEL").ToString & "'")
                        L_MID = IIf(dt.Rows(i)("L_MID").ToString = "", "null", "'" & dt.Rows(i)("L_MID").ToString & "'")
                        L_TID = IIf(dt.Rows(i)("L_TID").ToString = "", "null", "'" & dt.Rows(i)("L_TID").ToString & "'")

                        sql = "INSERT INTO [dbo].[POS_CONFIG]
                               ([POS_ID]
                               ,[MODULE_TYPE]
                               ,[POS_NO]
                               ,[POS_IP]
                               ,[TERMINAL_ID]
                               ,[RDCODE]
                               ,[ISMAIN]
                               ,[MAXMONEY]
                               ,[EDC1_PORTNAME]
                               ,[EDC2_PORTNAME]
                               ,[EDC_SPEED]
                               ,[EDC_PARITY]
                               ,[EDC_TIMEOUT]
                               ,[CLEAR_SALE_INFO]
                               ,[SHOW_PUMP_INFO]
                               ,[CASHDRAWERPRINT]
                               ,[AUTOPRINT]
                               ,[CREATEDATE]
                               ,[MODDATE]
                               ,[MODBY]
                               ,[EDC_MODEL]
                               ,[L_MID]
                               ,[L_TID])
                         VALUES
                               (" & POS_ID & "
                               ," & MODULE_TYPE & "
                               ," & POS_NO & "
                               ," & POS_IP & "
                               ," & TERMINAL_ID & "
                               ," & RDCODE & "
                               ," & ISMAIN & "
                               ," & MAXMONEY & "
                               ," & EDC1_PORTNAME & "
                               ," & EDC2_PORTNAME & "
                               ," & EDC_SPEED & "
                               ," & EDC_PARITY & "
                               ," & EDC_TIMEOUT & "
                               ," & CLEAR_SALE_INFO & "
                               ," & SHOW_PUMP_INFO & "
                               ," & CASHDRAWERPRINT & "
                               ," & AUTOPRINT & "
                               ,getdate()
                               ,getdate()
                               ,'" & modby & "'
                               ," & EDC_MODEL & "
                               ," & L_MID & "
                               ," & L_TID & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next

#End Region

                Case "TBUSER"
#Region "TBUSER"
                    Dim USERNAME, PASSWORD, USERDESC, EXPIRE_DATE, ISUSER, POSITION_ID, ISAUTOCLEAR, USER_ID, F_NAME, L_NAME As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        USERNAME = IIf(dt.Rows(i)("USERNAME").ToString = "", "null", "'" & dt.Rows(i)("USERNAME").ToString & "'")
                        PASSWORD = IIf(dt.Rows(i)("PASSWORD").ToString = "", "null", "'" & base64Encode(dt.Rows(i)("PASSWORD").ToString, dt.Rows(i)("USERNAME").ToString) & "'")
                        USERDESC = IIf(dt.Rows(i)("USERDESC").ToString = "", "null", "'" & dt.Rows(i)("USERDESC").ToString & "'")
                        EXPIRE_DATE = IIf(dt.Rows(i)("EXPIRE_DATE").ToString = "", "null", "" & dt.Rows(i)("EXPIRE_DATE").ToString & "")
                        If EXPIRE_DATE <> "null" Then
                            EXPIRE_DATE = ConvertDate(EXPIRE_DATE)
                        End If

                        ISUSER = IIf(dt.Rows(i)("ISUSER").ToString = "", "null", "'" & dt.Rows(i)("ISUSER").ToString & "'")
                        POSITION_ID = IIf(dt.Rows(i)("POSITION_ID").ToString = "", "null", "'" & dt.Rows(i)("POSITION_ID").ToString & "'")
                        ISAUTOCLEAR = "0" 'IIf(dt.Rows(i)("ISAUTOCLEAR").ToString = "", "null", "'" & dt.Rows(i)("ISAUTOCLEAR").ToString & "'")
                        USER_ID = "29"
                        F_NAME = "null" 'IIf(dt.Rows(i)("F_NAME").ToString = "", "null", "'" & dt.Rows(i)("F_NAME").ToString & "'")
                        L_NAME = "null" 'IIf(dt.Rows(i)("L_NAME").ToString = "", "null", "'" & dt.Rows(i)("L_NAME").ToString & "'")

                        sql = "INSERT INTO [dbo].[TBUSER]
                               ([USERNAME]
                               ,[PASSWORD]
                               ,[USERDESC]
                               ,[EXPIRE_DATE]
                               ,[ISUSER]
                               ,[POSITION_ID]
                               ,[CREATEDATE]
                               ,[MODDATE]
                               ,[MODBY]
                               ,[ISAUTOCLEAR]
                               ,[F_NAME]
                               ,[L_NAME])
                         VALUES
                               (" & USERNAME & "
                               ," & PASSWORD & "
                               ," & USERDESC & "
                               ," & EXPIRE_DATE & "
                               ," & ISUSER & "
                               ," & POSITION_ID & "
                               ,getdate()
                               ,getdate()
                               ,'" & modby & "'
                               ," & ISAUTOCLEAR & "
                               ," & F_NAME & "
                               ," & L_NAME & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TSJOURNAL"
#Region "TSJOURNAL"

                    sql = "If Not EXISTS(SELECT 1 FROM sys.columns "
                    sql &= " WHERE Name = N'IS_VOID_EDCWIFI'"
                    sql &= " And Object_ID = Object_ID(N'dbo.TSJOURNAL'))"
                    sql &= " BEGIN "
                    sql &= " ALTER TABLE TSJOURNAL "
                    sql &= " ADD IS_VOID_EDCWIFI CHAR(1) "
                    sql &= " End "
                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandText = sql
                        .CommandType = CommandType.Text
                        .Transaction = trans
                        .Connection = conn
                        .ExecuteNonQuery()
                    End With


                    sql = "If Not EXISTS(SELECT 1 FROM sys.columns "
                    sql &= " WHERE Name = N'IS_VOID_EDCWIFI'"
                    sql &= " And Object_ID = Object_ID(N'dbo.BKJOURNAL'))"
                    sql &= " BEGIN "
                    sql &= " ALTER TABLE BKJOURNAL "
                    sql &= " ADD IS_VOID_EDCWIFI CHAR(1) "
                    sql &= " End "

                    cmd.CommandText = sql
                    cmd.ExecuteNonQuery()

                    Dim JOURNAL_ID, POS_ID, USERNAME, TAX_NO, DAY_ID, SHIFT_ID, SALE_TYPE, CUS_ID, CUR_VAT As String
                    Dim TOTAL, VATTOTAL, DC, GRANDTOTAL, REFUND, VEHICLE_ID, CAR_TYPE, CARD_NO, CARD_TYPE As String
                    Dim INVOICE_NO, APPROVE_CODE, Signature, PRINT_TIMES, REF_JOURNAL_ID, PRICE_ID As String
                    Dim DEPT, DEPT1, COUNTER, EMP_NUMNER, DISTANCE, ACC_NUMBER, CARD_EXPIRE, TAX_CLASS As String
                    Dim DOCNO, LICENCENO, TAX_INVOICE, MCARDNO, LCARDNO, LCARDDATA, LREPOINT, LTRANS_NO As String
                    Dim LBATCH_NO, LCUSTOMER, LBALANCE, LPOINTTODAY, LREMARK, LPAY, LSTAND_ID As String
                    Dim LREDEEM_TRAN_ID, FLEET_HOST_ID, FLEET_CUST_TAX_ID, FLEET_CUST_BRANCH_NBR As String
                    Dim FLEET_CUST_NAME, FLEET_CUST_ADDRESS, FLEET_CAR_PLATE, IS_VOID_EDC As String
                    Dim AVAILABLE_CREDIT, FG_REF_NO, FLEET_DOCNO, FLEET_CUS_ID, REASON_ID, REASON_DESC As String
                    Dim ORIGINAL_TAX, DOC_TYPE, STATUS_FLAG, TOTAL_BALANCE, ORIGINAL_TAX_ID, FULLTAX_BOOKING_NUM As String
                    Dim DC_BILL_VALUE, DC_BILL_AMOUNT, DC_BILL_TYPE, TRANSACTION_ID, LRESCODE, MethodParam As String
                    Dim RoundSF, EarnByBiz, BILL_PROMOTION_ID, BILL_PROMOTION_PRINT_TIMES As String
                    Dim IS_VOID_EDCWIFI, CREATEDATE As String


                    For i As Integer = 0 To dt.Rows.Count - 1
                        JOURNAL_ID = IIf(dt.Rows(i)("JOURNAL_ID").ToString = "", "null", "'" & dt.Rows(i)("JOURNAL_ID").ToString & "'")
                        POS_ID = IIf(dt.Rows(i)("POS_ID").ToString = "", "null", "'" & dt.Rows(i)("POS_ID").ToString & "'")
                        USERNAME = IIf(dt.Rows(i)("USERNAME").ToString = "", "null", "'" & dt.Rows(i)("USERNAME").ToString & "'")
                        TAX_NO = IIf(dt.Rows(i)("TAX_NO").ToString = "", "null", "'" & dt.Rows(i)("TAX_NO").ToString & "'")
                        DAY_ID = IIf(dt.Rows(i)("DAY_ID").ToString = "", "null", "'" & dt.Rows(i)("DAY_ID").ToString & "'")
                        SHIFT_ID = IIf(dt.Rows(i)("SHIFT_ID").ToString = "", "null", "'" & dt.Rows(i)("SHIFT_ID").ToString & "'")
                        SALE_TYPE = IIf(dt.Rows(i)("SALE_TYPE").ToString = "", "null", "'" & dt.Rows(i)("SALE_TYPE").ToString & "'")
                        CUS_ID = IIf(dt.Rows(i)("CUS_ID").ToString = "", "null", "'" & dt.Rows(i)("CUS_ID").ToString & "'")
                        CUR_VAT = IIf(dt.Rows(i)("CUR_VAT").ToString = "", "null", "'" & dt.Rows(i)("CUR_VAT").ToString & "'")
                        TOTAL = IIf(dt.Rows(i)("TOTAL").ToString = "", "null", "'" & dt.Rows(i)("TOTAL").ToString & "'")

                        VATTOTAL = IIf(dt.Rows(i)("VATTOTAL").ToString = "", "null", "'" & dt.Rows(i)("VATTOTAL").ToString & "'")
                        DC = IIf(dt.Rows(i)("DC").ToString = "", "null", "'" & dt.Rows(i)("DC").ToString & "'")
                        GRANDTOTAL = IIf(dt.Rows(i)("GRANDTOTAL").ToString = "", "null", "'" & dt.Rows(i)("GRANDTOTAL").ToString & "'")
                        REFUND = IIf(dt.Rows(i)("REFUND").ToString = "", "null", "'" & dt.Rows(i)("REFUND").ToString & "'")
                        VEHICLE_ID = IIf(dt.Rows(i)("VEHICLE_ID").ToString = "", "null", "'" & dt.Rows(i)("VEHICLE_ID").ToString & "'")
                        CAR_TYPE = IIf(dt.Rows(i)("CAR_TYPE").ToString = "", "null", "'" & dt.Rows(i)("CAR_TYPE").ToString & "'")
                        CARD_NO = IIf(dt.Rows(i)("CARD_NO").ToString = "", "null", "'" & dt.Rows(i)("CARD_NO").ToString & "'")
                        CARD_TYPE = IIf(dt.Rows(i)("CARD_TYPE").ToString = "", "null", "'" & dt.Rows(i)("CARD_TYPE").ToString & "'")
                        INVOICE_NO = IIf(dt.Rows(i)("INVOICE_NO").ToString = "", "null", "'" & dt.Rows(i)("INVOICE_NO").ToString & "'")
                        APPROVE_CODE = IIf(dt.Rows(i)("APPROVE_CODE").ToString = "", "null", "'" & dt.Rows(i)("APPROVE_CODE").ToString & "'")

                        Signature = IIf(dt.Rows(i)("Signature").ToString = "", "null", "'" & dt.Rows(i)("Signature").ToString & "'")
                        PRINT_TIMES = IIf(dt.Rows(i)("PRINT_TIMES").ToString = "", "null", "'" & dt.Rows(i)("PRINT_TIMES").ToString & "'")
                        REF_JOURNAL_ID = IIf(dt.Rows(i)("REF_JOURNAL_ID").ToString = "", "null", "'" & dt.Rows(i)("REF_JOURNAL_ID").ToString & "'")
                        PRICE_ID = IIf(dt.Rows(i)("PRICE_ID").ToString = "", "null", "'" & dt.Rows(i)("PRICE_ID").ToString & "'")
                        DEPT = IIf(dt.Rows(i)("DEPT").ToString = "", "null", "'" & dt.Rows(i)("DEPT").ToString & "'")
                        DEPT1 = IIf(dt.Rows(i)("DEPT1").ToString = "", "null", "'" & dt.Rows(i)("DEPT1").ToString & "'")
                        COUNTER = IIf(dt.Rows(i)("COUNTER").ToString = "", "null", "'" & dt.Rows(i)("COUNTER").ToString & "'")
                        EMP_NUMNER = IIf(dt.Rows(i)("EMP_NUMNER").ToString = "", "null", "'" & dt.Rows(i)("EMP_NUMNER").ToString & "'")
                        DISTANCE = IIf(dt.Rows(i)("DISTANCE").ToString = "", "null", "'" & dt.Rows(i)("DISTANCE").ToString & "'")
                        ACC_NUMBER = IIf(dt.Rows(i)("ACC_NUMBER").ToString = "", "null", "'" & dt.Rows(i)("ACC_NUMBER").ToString & "'")

                        CARD_EXPIRE = IIf(dt.Rows(i)("CARD_EXPIRE").ToString = "", "null", "'" & dt.Rows(i)("CARD_EXPIRE").ToString & "'")
                        TAX_CLASS = IIf(dt.Rows(i)("TAX_CLASS").ToString = "", "null", "'" & dt.Rows(i)("TAX_CLASS").ToString & "'")
                        DOCNO = IIf(dt.Rows(i)("DOCNO").ToString = "", "null", "'" & dt.Rows(i)("DOCNO").ToString & "'")
                        LICENCENO = IIf(dt.Rows(i)("LICENCENO").ToString = "", "null", "'" & dt.Rows(i)("LICENCENO").ToString & "'")
                        TAX_INVOICE = IIf(dt.Rows(i)("TAX_INVOICE").ToString = "", "null", "'" & dt.Rows(i)("TAX_INVOICE").ToString & "'")
                        MCARDNO = IIf(dt.Rows(i)("MCARDNO").ToString = "", "null", "'" & dt.Rows(i)("MCARDNO").ToString & "'")
                        LCARDNO = IIf(dt.Rows(i)("LCARDNO").ToString = "", "null", "'" & dt.Rows(i)("LCARDNO").ToString & "'")
                        LCARDDATA = IIf(dt.Rows(i)("LCARDDATA").ToString = "", "null", "'" & dt.Rows(i)("LCARDDATA").ToString & "'")
                        LREPOINT = IIf(dt.Rows(i)("LREPOINT").ToString = "", "null", "'" & dt.Rows(i)("LREPOINT").ToString & "'")
                        LTRANS_NO = IIf(dt.Rows(i)("LTRANS_NO").ToString = "", "null", "'" & dt.Rows(i)("LTRANS_NO").ToString & "'")

                        LBATCH_NO = IIf(dt.Rows(i)("LBATCH_NO").ToString = "", "null", "'" & dt.Rows(i)("LBATCH_NO").ToString & "'")
                        LCUSTOMER = IIf(dt.Rows(i)("LCUSTOMER").ToString = "", "null", "'" & dt.Rows(i)("LCUSTOMER").ToString & "'")
                        LBALANCE = IIf(dt.Rows(i)("LBALANCE").ToString = "", "null", "'" & dt.Rows(i)("LBALANCE").ToString & "'")
                        LPOINTTODAY = IIf(dt.Rows(i)("LPOINTTODAY").ToString = "", "null", "'" & dt.Rows(i)("LPOINTTODAY").ToString & "'")
                        LREMARK = IIf(dt.Rows(i)("LREMARK").ToString = "", "null", "'" & dt.Rows(i)("LREMARK").ToString & "'")
                        LPAY = IIf(dt.Rows(i)("LPAY").ToString = "", "null", "'" & dt.Rows(i)("LPAY").ToString & "'")
                        LSTAND_ID = IIf(dt.Rows(i)("LSTAND_ID").ToString = "", "null", "'" & dt.Rows(i)("LSTAND_ID").ToString & "'")
                        LREDEEM_TRAN_ID = IIf(dt.Rows(i)("LREDEEM_TRAN_ID").ToString = "", "null", "'" & dt.Rows(i)("LREDEEM_TRAN_ID").ToString & "'")
                        FLEET_HOST_ID = IIf(dt.Rows(i)("FLEET_HOST_ID").ToString = "", "null", "'" & dt.Rows(i)("FLEET_HOST_ID").ToString & "'")
                        FLEET_CUST_TAX_ID = IIf(dt.Rows(i)("FLEET_CUST_TAX_ID").ToString = "", "null", "'" & dt.Rows(i)("FLEET_CUST_TAX_ID").ToString & "'")


                        FLEET_CUST_BRANCH_NBR = IIf(dt.Rows(i)("FLEET_CUST_BRANCH_NBR").ToString = "", "null", "'" & dt.Rows(i)("FLEET_CUST_BRANCH_NBR").ToString & "'")
                        FLEET_CUST_NAME = IIf(dt.Rows(i)("FLEET_CUST_NAME").ToString = "", "null", "'" & dt.Rows(i)("FLEET_CUST_NAME").ToString & "'")
                        FLEET_CUST_ADDRESS = IIf(dt.Rows(i)("FLEET_CUST_ADDRESS").ToString = "", "null", "'" & dt.Rows(i)("FLEET_CUST_ADDRESS").ToString & "'")
                        FLEET_CAR_PLATE = IIf(dt.Rows(i)("FLEET_CAR_PLATE").ToString = "", "null", "'" & dt.Rows(i)("FLEET_CAR_PLATE").ToString & "'")
                        IS_VOID_EDC = IIf(dt.Rows(i)("IS_VOID_EDC").ToString = "", "null", "'" & dt.Rows(i)("IS_VOID_EDC").ToString & "'")
                        AVAILABLE_CREDIT = IIf(dt.Rows(i)("AVAILABLE_CREDIT").ToString = "", "null", "'" & dt.Rows(i)("AVAILABLE_CREDIT").ToString & "'")
                        FG_REF_NO = IIf(dt.Rows(i)("FG_REF_NO").ToString = "", "null", "'" & dt.Rows(i)("FG_REF_NO").ToString & "'")
                        FLEET_DOCNO = IIf(dt.Rows(i)("FLEET_DOCNO").ToString = "", "null", "'" & dt.Rows(i)("FLEET_DOCNO").ToString & "'")
                        FLEET_CUS_ID = IIf(dt.Rows(i)("FLEET_CUS_ID").ToString = "", "null", "'" & dt.Rows(i)("FLEET_CUS_ID").ToString & "'")
                        REASON_ID = IIf(dt.Rows(i)("REASON_ID").ToString = "", "null", "'" & dt.Rows(i)("REASON_ID").ToString & "'")

                        REASON_DESC = IIf(dt.Rows(i)("REASON_DESC").ToString = "", "null", "'" & dt.Rows(i)("REASON_DESC").ToString & "'")
                        ORIGINAL_TAX = IIf(dt.Rows(i)("ORIGINAL_TAX").ToString = "", "null", "'" & dt.Rows(i)("ORIGINAL_TAX").ToString & "'")
                        DOC_TYPE = IIf(dt.Rows(i)("DOC_TYPE").ToString = "", "null", "'" & dt.Rows(i)("DOC_TYPE").ToString & "'")
                        STATUS_FLAG = IIf(dt.Rows(i)("STATUS_FLAG").ToString = "", "null", "'" & dt.Rows(i)("STATUS_FLAG").ToString & "'")
                        TOTAL_BALANCE = IIf(dt.Rows(i)("TOTAL_BALANCE").ToString = "", "null", "'" & dt.Rows(i)("TOTAL_BALANCE").ToString & "'")
                        ORIGINAL_TAX_ID = IIf(dt.Rows(i)("ORIGINAL_TAX_ID").ToString = "", "null", "'" & dt.Rows(i)("ORIGINAL_TAX_ID").ToString & "'")
                        FULLTAX_BOOKING_NUM = IIf(dt.Rows(i)("FULLTAX_BOOKING_NUM").ToString = "", "null", "'" & dt.Rows(i)("FULLTAX_BOOKING_NUM").ToString & "'")
                        DC_BILL_VALUE = "null" ' IIf(dt.Rows(i)("DC_BILL_VALUE").ToString = "", "null", "'" & dt.Rows(i)("DC_BILL_VALUE").ToString & "'")
                        DC_BILL_AMOUNT = "null" 'IIf(dt.Rows(i)("DC_BILL_AMOUNT").ToString = "", "null", "'" & dt.Rows(i)("DC_BILL_AMOUNT").ToString & "'")
                        DC_BILL_TYPE = "null" 'IIf(dt.Rows(i)("DC_BILL_TYPE").ToString = "", "null", "'" & dt.Rows(i)("DC_BILL_TYPE").ToString & "'")

                        TRANSACTION_ID = "null" 'IIf(dt.Rows(i)("TRANSACTION_ID").ToString = "", "null", "'" & dt.Rows(i)("TRANSACTION_ID").ToString & "'")
                        LRESCODE = IIf(dt.Rows(i)("LRESCODE").ToString = "", "null", "'" & dt.Rows(i)("LRESCODE").ToString & "'")
                        MethodParam = IIf(dt.Rows(i)("MethodParam").ToString = "", "null", "'" & dt.Rows(i)("MethodParam").ToString & "'")
                        RoundSF = IIf(dt.Rows(i)("RoundSF").ToString = "", "null", "'" & dt.Rows(i)("RoundSF").ToString & "'")
                        EarnByBiz = IIf(dt.Rows(i)("EarnByBiz").ToString = "", "null", "'" & dt.Rows(i)("EarnByBiz").ToString & "'")
                        BILL_PROMOTION_ID = "null" 'IIf(dt.Rows(i)("BILL_PROMOTION_ID").ToString = "", "null", "'" & dt.Rows(i)("BILL_PROMOTION_ID").ToString & "'")
                        BILL_PROMOTION_PRINT_TIMES = "0" 'IIf(dt.Rows(i)("BILL_PROMOTION_PRINT_TIMES").ToString = "", "null", "'" & dt.Rows(i)("BILL_PROMOTION_PRINT_TIMES").ToString & "'")

                        Dim columns As DataColumnCollection = dt.Columns
                        If columns.Contains("IS_VOID_EDCWIFI") Then
                            IS_VOID_EDCWIFI = IIf(dt.Rows(i)("IS_VOID_EDCWIFI").ToString = "", "null", "'" & dt.Rows(i)("IS_VOID_EDCWIFI").ToString & "'")
                        Else
                            IS_VOID_EDCWIFI = "null"
                        End If

                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ConvertDate(CREATEDATE)
                        End If



                        sql = "INSERT INTO [dbo].[TSJOURNAL]
                               ([JOURNAL_ID]
                               ,[POS_ID]
                               ,[USERNAME]
                               ,[TAX_NO]
                               ,[DAY_ID]
                               ,[SHIFT_ID]
                               ,[SALE_TYPE]
                               ,[CUS_ID]
                               ,[CUR_VAT]
                               ,[TOTAL]
                               ,[VATTOTAL]
                               ,[DC]
                               ,[GRANDTOTAL]
                               ,[REFUND]
                               ,[VEHICLE_ID]
                               ,[CAR_TYPE]
                               ,[CARD_NO]
                               ,[CARD_TYPE]
                               ,[INVOICE_NO]
                               ,[APPROVE_CODE]
                               ,[SIGNATURE]
                               ,[PRINT_TIMES]
                               ,[REF_JOURNAL_ID]
                               ,[PRICE_ID]
                               ,[CREATEDATE]
                               ,[MODDATE]
                               ,[MODBY]
                               ,[DEPT]
                               ,[DEPT1]
                               ,[COUNTER]
                               ,[EMP_NUMNER]
                               ,[DISTANCE]
                               ,[ACC_NUMBER]
                               ,[CARD_EXPIRE]
                               ,[TAX_CLASS]
                               ,[DOCNO]
                               ,[LICENCENO]
                               ,[TAX_INVOICE]
                               ,[MCARDNO]
                               ,[LCARDNO]
                               ,[LCARDDATA]
                               ,[LREPOINT]
                               ,[LTRANS_NO]
                               ,[LBATCH_NO]
                               ,[LCUSTOMER]
                               ,[LBALANCE]
                               ,[LPOINTTODAY]
                               ,[LREMARK]
                               ,[LPAY]
                               ,[LSTAND_ID]
                               ,[LREDEEM_TRAN_ID]
                               ,[FLEET_HOST_ID]
                               ,[FLEET_CUST_TAX_ID]
                               ,[FLEET_CUST_BRANCH_NBR]
                               ,[FLEET_CUST_NAME]
                               ,[FLEET_CUST_ADDRESS]
                               ,[FLEET_CAR_PLATE]
                               ,[IS_VOID_EDC]
                               ,[AVAILABLE_CREDIT]
                               ,[FG_REF_NO]
                               ,[FLEET_DOCNO]
                               ,[FLEET_CUS_ID]
                               ,[REASON_ID]
                               ,[REASON_DESC]
                               ,[ORIGINAL_TAX]
                               ,[DOC_TYPE]
                               ,[STATUS_FLAG]
                               ,[TOTAL_BALANCE]
                               ,[ORIGINAL_TAX_ID]
                               ,[FULLTAX_BOOKING_NUM]
                               ,[DC_BILL_VALUE]
                               ,[DC_BILL_AMOUNT]
                               ,[DC_BILL_TYPE]
                               ,[TRANSACTION_ID]
                               ,[LRESCODE]
                               ,[MethodParam]
                               ,[RoundSF]
                               ,[EarnByBiz]
                               ,[BILL_PROMOTION_ID]
                               ,[BILL_PROMOTION_PRINT_TIMES],[IS_VOID_EDCWIFI])
                         VALUES
                               (" & JOURNAL_ID & "
                               ," & POS_ID & "
                               ," & USERNAME & "
                               ," & TAX_NO & "
                               ," & DAY_ID & "
                               ," & SHIFT_ID & "
                               ," & SALE_TYPE & "
                               ," & CUS_ID & "
                               ," & CUR_VAT & "
                               ," & TOTAL & "
                               ," & VATTOTAL & "
                               ," & DC & "
                               ," & GRANDTOTAL & "
                               ," & REFUND & "
                               ," & VEHICLE_ID & "
                               ," & CAR_TYPE & "
                               ," & CARD_NO & "
                               ," & CARD_TYPE & "
                               ," & INVOICE_NO & "
                               ," & APPROVE_CODE & "
                               ," & Signature & "
                               ," & PRINT_TIMES & "
                               ," & REF_JOURNAL_ID & "
                               ," & PRICE_ID & "
                               ," & CREATEDATE & "
                               ,getdate()
                               ,'" & modby & "'
                               ," & DEPT & "
                               ," & DEPT1 & "
                               ," & COUNTER & "
                               ," & EMP_NUMNER & "
                               ," & DISTANCE & "
                               ," & ACC_NUMBER & "
                               ," & CARD_EXPIRE & "
                               ," & TAX_CLASS & "
                               ," & DOCNO & "
                               ," & LICENCENO & "
                               ," & TAX_INVOICE & "
                               ," & MCARDNO & "
                               ," & LCARDNO & "
                               ," & LCARDDATA & "
                               ," & LREPOINT & "
                               ," & LTRANS_NO & "
                               ," & LBATCH_NO & "
                               ," & LCUSTOMER & "
                               ," & LBALANCE & "
                               ," & LPOINTTODAY & "
                               ," & LREMARK & "
                               ," & LPAY & "
                               ," & LSTAND_ID & "
                               ," & LREDEEM_TRAN_ID & "
                               ," & FLEET_HOST_ID & "
                               ," & FLEET_CUST_TAX_ID & "
                               ," & FLEET_CUST_BRANCH_NBR & "
                               ," & FLEET_CUST_NAME & "
                               ," & FLEET_CUST_ADDRESS & "
                               ," & FLEET_CAR_PLATE & "
                               ," & IS_VOID_EDC & "
                               ," & AVAILABLE_CREDIT & "
                               ," & FG_REF_NO & "
                               ," & FLEET_DOCNO & "
                               ," & FLEET_CUS_ID & "
                               ," & REASON_ID & "
                               ," & REASON_DESC & "
                               ," & ORIGINAL_TAX & "
                               ," & DOC_TYPE & "
                               ," & STATUS_FLAG & "
                               ," & TOTAL_BALANCE & "
                               ," & ORIGINAL_TAX_ID & "
                               ," & FULLTAX_BOOKING_NUM & "
                               ," & DC_BILL_VALUE & "
                               ," & DC_BILL_AMOUNT & "
                               ," & DC_BILL_TYPE & "
                               ," & TRANSACTION_ID & "
                               ," & LRESCODE & "
                               ," & MethodParam & "
                               ," & RoundSF & "
                               ," & EarnByBiz & "
                               ," & BILL_PROMOTION_ID & "
                               ," & BILL_PROMOTION_PRINT_TIMES & "," & IS_VOID_EDCWIFI & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TSJOURNAL_DETAIL"
#Region "TSJOURNAL_DETAIL"
                    Dim JOURNAL_ID, ITEM_NO, MAT_ID, VOLUME, QTY, PRICE, VALUE, DC_PRICE, DC_VALUE, TRANS_NO,
                    HOSE_ID, PUMP_ID, TANK_ID, IS_OFFLINE, JOURNAL_REF, DC_ITEM_VALUE, DC_ITEM_AMOUNT,
                    DC_ITEM_TYPE, DC_BILL_VALUE, DC_BILL_AMOUNT, DC_BILL_TYPE, VAT_TYPE, VATable,
                    VAT, Total, NetBeforeVat, NetTotal As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        JOURNAL_ID = IIf(dt.Rows(i)("JOURNAL_ID").ToString = "", "null", "'" & dt.Rows(i)("JOURNAL_ID").ToString & "'")
                        ITEM_NO = IIf(dt.Rows(i)("ITEM_NO").ToString = "", "null", "'" & dt.Rows(i)("ITEM_NO").ToString & "'")
                        MAT_ID = IIf(dt.Rows(i)("MAT_ID").ToString = "", "null", "'" & dt.Rows(i)("MAT_ID").ToString & "'")
                        VOLUME = IIf(dt.Rows(i)("VOLUME").ToString = "", "null", "'" & dt.Rows(i)("VOLUME").ToString & "'")
                        QTY = IIf(dt.Rows(i)("QTY").ToString = "", "null", "'" & dt.Rows(i)("QTY").ToString & "'")
                        PRICE = IIf(dt.Rows(i)("PRICE").ToString = "", "null", "'" & dt.Rows(i)("PRICE").ToString & "'")
                        VALUE = IIf(dt.Rows(i)("VALUE").ToString = "", "null", "'" & dt.Rows(i)("VALUE").ToString & "'")
                        DC_PRICE = IIf(dt.Rows(i)("DC_PRICE").ToString = "", "null", "'" & dt.Rows(i)("DC_PRICE").ToString & "'")
                        DC_VALUE = IIf(dt.Rows(i)("DC_VALUE").ToString = "", "null", "'" & dt.Rows(i)("DC_VALUE").ToString & "'")
                        TRANS_NO = IIf(dt.Rows(i)("TRANS_NO").ToString = "", "null", "'" & dt.Rows(i)("TRANS_NO").ToString & "'")

                        HOSE_ID = IIf(dt.Rows(i)("HOSE_ID").ToString = "", "null", "'" & dt.Rows(i)("HOSE_ID").ToString & "'")
                        PUMP_ID = IIf(dt.Rows(i)("PUMP_ID").ToString = "", "null", "'" & dt.Rows(i)("PUMP_ID").ToString & "'")
                        TANK_ID = IIf(dt.Rows(i)("TANK_ID").ToString = "", "null", "'" & dt.Rows(i)("TANK_ID").ToString & "'")
                        IS_OFFLINE = IIf(dt.Rows(i)("IS_OFFLINE").ToString = "", "null", "'" & dt.Rows(i)("IS_OFFLINE").ToString & "'")
                        JOURNAL_REF = "null" 'IIf(dt.Rows(i)("JOURNAL_REF").ToString = "", "null", "'" & dt.Rows(i)("JOURNAL_REF").ToString & "'")
                        DC_ITEM_VALUE = "null" 'IIf(dt.Rows(i)("DC_ITEM_VALUE").ToString = "", "null", "'" & dt.Rows(i)("DC_ITEM_VALUE").ToString & "'")
                        DC_ITEM_AMOUNT = "null" 'IIf(dt.Rows(i)("DC_ITEM_AMOUNT").ToString = "", "null", "'" & dt.Rows(i)("DC_ITEM_AMOUNT").ToString & "'")
                        DC_ITEM_TYPE = "null" 'IIf(dt.Rows(i)("DC_ITEM_TYPE").ToString = "", "null", "'" & dt.Rows(i)("DC_ITEM_TYPE").ToString & "'")
                        DC_BILL_VALUE = "null" 'IIf(dt.Rows(i)("DC_BILL_VALUE").ToString = "", "null", "'" & dt.Rows(i)("DC_BILL_VALUE").ToString & "'")
                        DC_BILL_AMOUNT = "null" 'IIf(dt.Rows(i)("DC_BILL_AMOUNT").ToString = "", "null", "'" & dt.Rows(i)("DC_BILL_AMOUNT").ToString & "'")

                        DC_BILL_TYPE = "null" 'IIf(dt.Rows(i)("DC_BILL_TYPE").ToString = "", "null", "'" & dt.Rows(i)("DC_BILL_TYPE").ToString & "'")
                        VAT_TYPE = "null" 'IIf(dt.Rows(i)("VAT_TYPE").ToString = "", "null", "'" & dt.Rows(i)("VAT_TYPE").ToString & "'")
                        VATable = "null" 'IIf(dt.Rows(i)("VATable").ToString = "", "null", "'" & dt.Rows(i)("VATable").ToString & "'")
                        VAT = "null" 'IIf(dt.Rows(i)("VAT").ToString = "", "null", "'" & dt.Rows(i)("VAT").ToString & "'")
                        Total = "null" 'IIf(dt.Rows(i)("Total").ToString = "", "null", "'" & dt.Rows(i)("Total").ToString & "'")
                        NetBeforeVat = "null" ' IIf(dt.Rows(i)("NetBeforeVat").ToString = "", "null", "'" & dt.Rows(i)("NetBeforeVat").ToString & "'")
                        NetTotal = "null" ' IIf(dt.Rows(i)("NetTotal").ToString = "", "null", "'" & dt.Rows(i)("NetTotal").ToString & "'")

                        sql = "INSERT INTO [dbo].[TSJOURNAL_DETAIL]
                               ([JOURNAL_ID]
                               ,[ITEM_NO]
                               ,[MAT_ID]
                               ,[VOLUME]
                               ,[QTY]
                               ,[PRICE]
                               ,[VALUE]
                               ,[DC_PRICE]
                               ,[DC_VALUE]
                               ,[TRANS_NO]
                               ,[HOSE_ID]
                               ,[PUMP_ID]
                               ,[TANK_ID]
                               ,[IS_OFFLINE]
                               ,[JOURNAL_REF]
                               ,[DC_ITEM_VALUE]
                               ,[DC_ITEM_AMOUNT]
                               ,[DC_ITEM_TYPE]
                               ,[DC_BILL_VALUE]
                               ,[DC_BILL_AMOUNT]
                               ,[DC_BILL_TYPE]
                               ,[VAT_TYPE]
                               ,[VATable]
                               ,[VAT]
                               ,[Total]
                               ,[NetBeforeVat]
                               ,[NetTotal])
                         VALUES
                               (" & JOURNAL_ID & "
                               ," & ITEM_NO & "
                               ," & MAT_ID & "
                               ," & VOLUME & "
                               ," & QTY & "
                               ," & PRICE & "
                               ," & VALUE & "
                               ," & DC_PRICE & "
                               ," & DC_VALUE & "
                               ," & TRANS_NO & "
                               ," & HOSE_ID & "
                               ," & PUMP_ID & "
                               ," & TANK_ID & "
                               ," & IS_OFFLINE & "
                               ," & JOURNAL_REF & "
                               ," & DC_ITEM_VALUE & "
                               ," & DC_ITEM_AMOUNT & "
                               ," & DC_ITEM_TYPE & "
                               ," & DC_BILL_VALUE & "
                               ," & DC_BILL_AMOUNT & "
                               ," & DC_BILL_TYPE & "
                               ," & VAT_TYPE & "
                               ," & VATable & "
                               ," & VAT & "
                               ," & Total & "
                               ," & NetBeforeVat & "
                               ," & NetTotal & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TSJOURNAL_PAYMENT"
#Region "TSJOURNAL_PAYMENT"
                    Dim JOURNAL_ID, ITEM_NO, GROUP_ID, PAYMENT_TYPE, VALUE, DC, CUS_ID, VEHICLE_ID, VOUCHER_NO, CARD_NO, CARD_TYPE,
                        INVOICE_NO, APPROVE_CODE, SIGNATURE, EDCNO, REDEEMVALUE, NII, LTRNPOINT, LBATCH_NO, LTRANS_NO, LSTAND_ID,
                        VOID_DATE, VOID_APPROVE_CODE, PO_NO, VEH_CODE, LCOUPON_QTY, LCOUPON_CODE As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        JOURNAL_ID = IIf(dt.Rows(i)("JOURNAL_ID").ToString = "", "null", "'" & dt.Rows(i)("JOURNAL_ID").ToString & "'")
                        ITEM_NO = IIf(dt.Rows(i)("ITEM_NO").ToString = "", "null", "'" & dt.Rows(i)("ITEM_NO").ToString & "'")
                        GROUP_ID = IIf(dt.Rows(i)("GROUP_ID").ToString = "", "null", "'" & dt.Rows(i)("GROUP_ID").ToString & "'")
                        PAYMENT_TYPE = IIf(dt.Rows(i)("PAYMENT_TYPE").ToString = "", "null", "'" & dt.Rows(i)("PAYMENT_TYPE").ToString & "'")
                        VALUE = IIf(dt.Rows(i)("VALUE").ToString = "", "null", "'" & dt.Rows(i)("VALUE").ToString & "'")
                        DC = IIf(dt.Rows(i)("DC").ToString = "", "null", "'" & dt.Rows(i)("DC").ToString & "'")
                        CUS_ID = IIf(dt.Rows(i)("CUS_ID").ToString = "", "null", "'" & dt.Rows(i)("CUS_ID").ToString & "'")
                        VEHICLE_ID = IIf(dt.Rows(i)("VEHICLE_ID").ToString = "", "null", "'" & dt.Rows(i)("VEHICLE_ID").ToString & "'")
                        VOUCHER_NO = IIf(dt.Rows(i)("VOUCHER_NO").ToString = "", "null", "'" & dt.Rows(i)("VOUCHER_NO").ToString & "'")
                        CARD_NO = IIf(dt.Rows(i)("CARD_NO").ToString = "", "null", "'" & dt.Rows(i)("CARD_NO").ToString & "'")

                        CARD_TYPE = IIf(dt.Rows(i)("CARD_TYPE").ToString = "", "null", "'" & dt.Rows(i)("CARD_TYPE").ToString & "'")
                        INVOICE_NO = IIf(dt.Rows(i)("INVOICE_NO").ToString = "", "null", "'" & dt.Rows(i)("INVOICE_NO").ToString & "'")
                        APPROVE_CODE = IIf(dt.Rows(i)("APPROVE_CODE").ToString = "", "null", "'" & dt.Rows(i)("APPROVE_CODE").ToString & "'")
                        SIGNATURE = IIf(dt.Rows(i)("SIGNATURE").ToString = "", "null", "'" & dt.Rows(i)("SIGNATURE").ToString & "'")
                        EDCNO = IIf(dt.Rows(i)("EDCNO").ToString = "", "null", "'" & dt.Rows(i)("EDCNO").ToString & "'")
                        REDEEMVALUE = IIf(dt.Rows(i)("REDEEMVALUE").ToString = "", "null", "'" & dt.Rows(i)("REDEEMVALUE").ToString & "'")
                        NII = IIf(dt.Rows(i)("NII").ToString = "", "null", "'" & dt.Rows(i)("NII").ToString & "'")
                        LTRNPOINT = IIf(dt.Rows(i)("LTRNPOINT").ToString = "", "null", "'" & dt.Rows(i)("LTRNPOINT").ToString & "'")
                        LBATCH_NO = IIf(dt.Rows(i)("LBATCH_NO").ToString = "", "null", "'" & dt.Rows(i)("LBATCH_NO").ToString & "'")
                        LTRANS_NO = IIf(dt.Rows(i)("LTRANS_NO").ToString = "", "null", "'" & dt.Rows(i)("LTRANS_NO").ToString & "'")

                        LSTAND_ID = IIf(dt.Rows(i)("LSTAND_ID").ToString = "", "null", "'" & dt.Rows(i)("LSTAND_ID").ToString & "'")
                        VOID_DATE = IIf(dt.Rows(i)("VOID_DATE").ToString = "", "null", "" & dt.Rows(i)("VOID_DATE").ToString & "")
                        If VOID_DATE <> "null" Then
                            VOID_DATE = ConvertDate(VOID_DATE)
                        End If
                        VOID_APPROVE_CODE = IIf(dt.Rows(i)("VOID_APPROVE_CODE").ToString = "", "null", "'" & dt.Rows(i)("VOID_APPROVE_CODE").ToString & "'")
                        PO_NO = "null" 'IIf(dt.Rows(i)("PO_NO").ToString = "", "null", "'" & dt.Rows(i)("PO_NO").ToString & "'")
                        VEH_CODE = "null" 'IIf(dt.Rows(i)("VEH_CODE").ToString = "", "null", "'" & dt.Rows(i)("VEH_CODE").ToString & "'")
                        LCOUPON_QTY = IIf(dt.Rows(i)("LCOUPON_QTY").ToString = "", "null", "'" & dt.Rows(i)("LCOUPON_QTY").ToString & "'")
                        LCOUPON_CODE = IIf(dt.Rows(i)("LCOUPON_CODE").ToString = "", "null", "'" & dt.Rows(i)("LCOUPON_CODE").ToString & "'")

                        sql = "INSERT INTO [dbo].[TSJOURNAL_PAYMENT]
                               ([JOURNAL_ID]
                               ,[ITEM_NO]
                               ,[GROUP_ID]
                               ,[PAYMENT_TYPE]
                               ,[VALUE]
                               ,[DC]
                               ,[CUS_ID]
                               ,[VEHICLE_ID]
                               ,[VOUCHER_NO]
                               ,[CARD_NO]
                               ,[CARD_TYPE]
                               ,[INVOICE_NO]
                               ,[APPROVE_CODE]
                               ,[SIGNATURE]
                               ,[EDCNO]
                               ,[REDEEMVALUE]
                               ,[NII]
                               ,[LTRNPOINT]
                               ,[LBATCH_NO]
                               ,[LTRANS_NO]
                               ,[LSTAND_ID]
                               ,[VOID_DATE]
                               ,[VOID_APPROVE_CODE]
                               ,[PO_NO]
                               ,[VEH_CODE]
                               ,[LCOUPON_QTY]
                               ,[LCOUPON_CODE])
                         VALUES
                               (" & JOURNAL_ID & "
                               ," & ITEM_NO & "
                               ," & GROUP_ID & "
                               ," & PAYMENT_TYPE & "
                               ," & VALUE & "
                               ," & DC & "
                               ," & CUS_ID & "
                               ," & VEHICLE_ID & "
                               ," & VOUCHER_NO & "
                               ," & CARD_NO & "
                               ," & CARD_TYPE & "
                               ," & INVOICE_NO & "
                               ," & APPROVE_CODE & "
                               ," & SIGNATURE & "
                               ," & EDCNO & "
                               ," & REDEEMVALUE & "
                               ," & NII & "
                               ," & LTRNPOINT & "
                               ," & LBATCH_NO & "
                               ," & LTRANS_NO & "
                               ," & LSTAND_ID & "
                               ," & VOID_DATE & "
                               ," & VOID_APPROVE_CODE & "
                               ," & PO_NO & "
                               ," & VEH_CODE & "
                               ," & LCOUPON_QTY & "
                               ," & LCOUPON_CODE & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBPERIODS"
#Region "TBPERIODS"
                    Dim PERIOD_ID, POS_ID, USER_OPEN, USER_CLOSE, PERIOD_CREATE_TS, PERIOD_CLOSE_DT, PERIOD_TYPE, PERIOD_STATE,
                    DAY_ID, BUS_DATE, SHIFT_NO, TANK_DIPS_ENTERED, TANK_DROPS_ENTERED, PERIOD_METER_ENTERED, EXPORTED,
                    EXPORT_REQUIRED, WETSTOCK_OUT_OF_VARIANCE, WETSTOCK_APPRROVAL_ID, PRINT_TIMES, PUMP_ALLOW, LOGONMODE,
                    TRANSFER_STATUS, TRANSFER_DATE As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        PERIOD_ID = IIf(dt.Rows(i)("PERIOD_ID").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_ID").ToString & "'")
                        POS_ID = IIf(dt.Rows(i)("POS_ID").ToString = "", "null", "'" & dt.Rows(i)("POS_ID").ToString & "'")
                        USER_OPEN = IIf(dt.Rows(i)("USER_OPEN").ToString = "", "null", "'" & dt.Rows(i)("USER_OPEN").ToString & "'")
                        USER_CLOSE = IIf(dt.Rows(i)("USER_CLOSE").ToString = "", "null", "'" & dt.Rows(i)("USER_CLOSE").ToString & "'")
                        PERIOD_CREATE_TS = IIf(dt.Rows(i)("PERIOD_CREATE_TS").ToString = "", "null", "" & dt.Rows(i)("PERIOD_CREATE_TS").ToString & "")
                        If PERIOD_CREATE_TS <> "null" Then
                            PERIOD_CREATE_TS = ConvertDateTime(PERIOD_CREATE_TS)
                        End If

                        PERIOD_CLOSE_DT = IIf(dt.Rows(i)("PERIOD_CLOSE_DT").ToString = "", "null", "" & dt.Rows(i)("PERIOD_CLOSE_DT").ToString & "")
                        If PERIOD_CLOSE_DT <> "null" Then
                            PERIOD_CLOSE_DT = ConvertDateTime(PERIOD_CLOSE_DT)
                        End If

                        PERIOD_TYPE = IIf(dt.Rows(i)("PERIOD_TYPE").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_TYPE").ToString & "'")
                        PERIOD_STATE = IIf(dt.Rows(i)("PERIOD_STATE").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_STATE").ToString & "'")
                        DAY_ID = IIf(dt.Rows(i)("DAY_ID").ToString = "", "null", "'" & dt.Rows(i)("DAY_ID").ToString & "'")
                        BUS_DATE = IIf(dt.Rows(i)("BUS_DATE").ToString = "", "null", "" & dt.Rows(i)("BUS_DATE").ToString & "")
                        If BUS_DATE <> "null" Then
                            BUS_DATE = ConvertDate(BUS_DATE)
                        End If

                        SHIFT_NO = IIf(dt.Rows(i)("SHIFT_NO").ToString = "", "null", "'" & dt.Rows(i)("SHIFT_NO").ToString & "'")
                        TANK_DIPS_ENTERED = IIf(dt.Rows(i)("TANK_DIPS_ENTERED").ToString = "", "null", "'" & dt.Rows(i)("TANK_DIPS_ENTERED").ToString & "'")
                        TANK_DROPS_ENTERED = IIf(dt.Rows(i)("TANK_DROPS_ENTERED").ToString = "", "null", "'" & dt.Rows(i)("TANK_DROPS_ENTERED").ToString & "'")
                        PERIOD_METER_ENTERED = IIf(dt.Rows(i)("PERIOD_METER_ENTERED").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_METER_ENTERED").ToString & "'")
                        EXPORTED = IIf(dt.Rows(i)("EXPORTED").ToString = "", "null", "'" & dt.Rows(i)("EXPORTED").ToString & "'")
                        EXPORT_REQUIRED = IIf(dt.Rows(i)("EXPORT_REQUIRED").ToString = "", "null", "'" & dt.Rows(i)("EXPORT_REQUIRED").ToString & "'")
                        WETSTOCK_OUT_OF_VARIANCE = IIf(dt.Rows(i)("WETSTOCK_OUT_OF_VARIANCE").ToString = "", "null", "'" & dt.Rows(i)("WETSTOCK_OUT_OF_VARIANCE").ToString & "'")
                        WETSTOCK_APPRROVAL_ID = IIf(dt.Rows(i)("WETSTOCK_APPRROVAL_ID").ToString = "", "null", "'" & dt.Rows(i)("WETSTOCK_APPRROVAL_ID").ToString & "'")
                        PRINT_TIMES = IIf(dt.Rows(i)("PRINT_TIMES").ToString = "", "null", "'" & dt.Rows(i)("PRINT_TIMES").ToString & "'")
                        PUMP_ALLOW = IIf(dt.Rows(i)("PUMP_ALLOW").ToString = "", "null", "'" & dt.Rows(i)("PUMP_ALLOW").ToString & "'")

                        LOGONMODE = IIf(dt.Rows(i)("LOGONMODE").ToString = "", "null", "'" & dt.Rows(i)("LOGONMODE").ToString & "'")
                        TRANSFER_STATUS = IIf(dt.Rows(i)("TRANSFER_STATUS").ToString = "", "null", "'" & dt.Rows(i)("TRANSFER_STATUS").ToString & "'")
                        TRANSFER_DATE = IIf(dt.Rows(i)("TRANSFER_DATE").ToString = "", "null", "" & dt.Rows(i)("TRANSFER_DATE").ToString & "")
                        If TRANSFER_DATE <> "null" Then
                            TRANSFER_DATE = ConvertDate(TRANSFER_DATE)
                        End If


                        sql = "INSERT INTO [dbo].[TBPERIODS]
                               ([PERIOD_ID]
                               ,[POS_ID]
                               ,[USER_OPEN]
                               ,[USER_CLOSE]
                               ,[PERIOD_CREATE_TS]
                               ,[PERIOD_CLOSE_DT]
                               ,[PERIOD_TYPE]
                               ,[PERIOD_STATE]
                               ,[DAY_ID]
                               ,[BUS_DATE]
                               ,[SHIFT_NO]
                               ,[TANK_DIPS_ENTERED]
                               ,[TANK_DROPS_ENTERED]
                               ,[PERIOD_METER_ENTERED]
                               ,[EXPORTED]
                               ,[EXPORT_REQUIRED]
                               ,[WETSTOCK_OUT_OF_VARIANCE]
                               ,[WETSTOCK_APPRROVAL_ID]
                               ,[PRINT_TIMES]
                               ,[PUMP_ALLOW]
                               ,[LOGONMODE]
                               ,[MODBY]
                               ,[TRANSFER_STATUS]
                               ,[TRANSFER_DATE])
                         VALUES
                               (" & PERIOD_ID & "
                               ," & POS_ID & "
                               ," & USER_OPEN & "
                               ," & USER_CLOSE & "
                               ," & PERIOD_CREATE_TS & "
                               ," & PERIOD_CLOSE_DT & "
                               ," & PERIOD_TYPE & "
                               ," & PERIOD_STATE & "
                               ," & DAY_ID & "
                               ," & BUS_DATE & "
                               ," & SHIFT_NO & "
                               ," & TANK_DIPS_ENTERED & "
                               ," & TANK_DROPS_ENTERED & "
                               ," & PERIOD_METER_ENTERED & "
                               ," & EXPORTED & "
                               ," & EXPORT_REQUIRED & "
                               ," & WETSTOCK_OUT_OF_VARIANCE & "
                               ," & WETSTOCK_APPRROVAL_ID & "
                               ," & PRINT_TIMES & "
                               ," & PUMP_ALLOW & "
                               ," & LOGONMODE & "
                               ,'" & modby & "'
                               ," & TRANSFER_STATUS & "
                               ," & TRANSFER_DATE & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBHOSE_HISTORY"
#Region "TBHOSE_HISTORY"
                    Dim HOSE_ID, PERIOD_ID, OPEN_METER_VALUE, CLOSE_METER_VALUE, OPEN_METER_VOLUME, CLOSE_METER_VOLUME,
                        POSTPAY_QUANTITY, POSTPAY_VALUE, POSTPAY_VOLUME, POSTPAY_COST, PREPAY_QUANTITY, PREPAY_VALUE,
                        PREPAY_VOLUME, PREPAY_COST, PREPAY_REFUND_QTY, PREPAY_REFUND_VAL, PREPAY_RFD_LST_QTY, PREPAY_RFD_LST_VAL,
                        PREAUTH_QUANTITY, PREAUTH_VALUE, PREAUTH_VOLUME, PREAUTH_COST, MONITOR_QUANTITY, MONITOR_VALUE,
                        MONITOR_VOLUME, MONITOR_COST, DRIVEOFFS_QUANTITY, DRIVEOFFS_VALUE, DRIVEOFFS_VOLUME, DRIVEOFFS_COST,
                        TEST_DEL_QUANTITY, TEST_DEL_VOLUME, OFFLINE_QUANTITY, OFFLINE_VOLUME, OFFLINE_VALUE, OFFLINE_COST,
                        OPEN_MECH_VOLUME, CLOSE_MECH_VOLUME, OPEN_VOLUME_TURNOVER_CORRECTION, OPEN_MONEY_TURNOVER_CORRECTION,
                        CLOSE_VOLUME_TURNOVER_CORRECTION, CLOSE_MONEY_TURNOVER_CORRECTION, OPEN_VOLUME_TURNOVER_CORRECTION2,
                        CLOSE_VOLUME_TURNOVER_CORRECTION2, PUMP_ID, MAT_ID, TANK_ID As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        HOSE_ID = IIf(dt.Rows(i)("HOSE_ID").ToString = "", "null", "'" & dt.Rows(i)("HOSE_ID").ToString & "'")
                        PERIOD_ID = IIf(dt.Rows(i)("PERIOD_ID").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_ID").ToString & "'")
                        OPEN_METER_VALUE = IIf(dt.Rows(i)("OPEN_METER_VALUE").ToString = "", "null", "'" & dt.Rows(i)("OPEN_METER_VALUE").ToString & "'")
                        CLOSE_METER_VALUE = IIf(dt.Rows(i)("CLOSE_METER_VALUE").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_METER_VALUE").ToString & "'")
                        OPEN_METER_VOLUME = IIf(dt.Rows(i)("OPEN_METER_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_METER_VOLUME").ToString & "'")
                        CLOSE_METER_VOLUME = IIf(dt.Rows(i)("CLOSE_METER_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_METER_VOLUME").ToString & "'")
                        POSTPAY_QUANTITY = IIf(dt.Rows(i)("POSTPAY_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("POSTPAY_QUANTITY").ToString & "'")
                        POSTPAY_VALUE = IIf(dt.Rows(i)("POSTPAY_VALUE").ToString = "", "null", "'" & dt.Rows(i)("POSTPAY_VALUE").ToString & "'")
                        POSTPAY_VOLUME = IIf(dt.Rows(i)("POSTPAY_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("POSTPAY_VOLUME").ToString & "'")
                        POSTPAY_COST = IIf(dt.Rows(i)("POSTPAY_COST").ToString = "", "null", "'" & dt.Rows(i)("POSTPAY_COST").ToString & "'")

                        PREPAY_QUANTITY = IIf(dt.Rows(i)("PREPAY_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_QUANTITY").ToString & "'")
                        PREPAY_VALUE = IIf(dt.Rows(i)("PREPAY_VALUE").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_VALUE").ToString & "'")
                        PREPAY_VOLUME = IIf(dt.Rows(i)("PREPAY_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_VOLUME").ToString & "'")
                        PREPAY_COST = IIf(dt.Rows(i)("PREPAY_COST").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_COST").ToString & "'")
                        PREPAY_REFUND_QTY = IIf(dt.Rows(i)("PREPAY_REFUND_QTY").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_REFUND_QTY").ToString & "'")
                        PREPAY_REFUND_VAL = IIf(dt.Rows(i)("PREPAY_REFUND_VAL").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_REFUND_VAL").ToString & "'")
                        PREPAY_RFD_LST_QTY = IIf(dt.Rows(i)("PREPAY_RFD_LST_QTY").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_RFD_LST_QTY").ToString & "'")
                        PREPAY_RFD_LST_VAL = IIf(dt.Rows(i)("PREPAY_RFD_LST_VAL").ToString = "", "null", "'" & dt.Rows(i)("PREPAY_RFD_LST_VAL").ToString & "'")
                        MONITOR_VOLUME = IIf(dt.Rows(i)("MONITOR_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("MONITOR_VOLUME").ToString & "'")
                        PREAUTH_QUANTITY = IIf(dt.Rows(i)("PREAUTH_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("PREAUTH_QUANTITY").ToString & "'")

                        PREAUTH_VALUE = IIf(dt.Rows(i)("PREAUTH_VALUE").ToString = "", "null", "'" & dt.Rows(i)("PREAUTH_VALUE").ToString & "'")
                        PREAUTH_VOLUME = IIf(dt.Rows(i)("PREAUTH_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("PREAUTH_VOLUME").ToString & "'")
                        PREAUTH_COST = IIf(dt.Rows(i)("PREAUTH_COST").ToString = "", "null", "'" & dt.Rows(i)("PREAUTH_COST").ToString & "'")
                        MONITOR_QUANTITY = IIf(dt.Rows(i)("MONITOR_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("MONITOR_QUANTITY").ToString & "'")
                        MONITOR_VALUE = IIf(dt.Rows(i)("MONITOR_VALUE").ToString = "", "null", "'" & dt.Rows(i)("MONITOR_VALUE").ToString & "'")
                        MONITOR_COST = IIf(dt.Rows(i)("MONITOR_COST").ToString = "", "null", "'" & dt.Rows(i)("MONITOR_COST").ToString & "'")
                        DRIVEOFFS_QUANTITY = IIf(dt.Rows(i)("DRIVEOFFS_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("DRIVEOFFS_QUANTITY").ToString & "'")
                        DRIVEOFFS_VALUE = IIf(dt.Rows(i)("DRIVEOFFS_VALUE").ToString = "", "null", "'" & dt.Rows(i)("DRIVEOFFS_VALUE").ToString & "'")
                        DRIVEOFFS_VOLUME = IIf(dt.Rows(i)("DRIVEOFFS_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("DRIVEOFFS_VOLUME").ToString & "'")
                        DRIVEOFFS_COST = IIf(dt.Rows(i)("DRIVEOFFS_COST").ToString = "", "null", "'" & dt.Rows(i)("DRIVEOFFS_COST").ToString & "'")

                        TEST_DEL_QUANTITY = IIf(dt.Rows(i)("TEST_DEL_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("TEST_DEL_QUANTITY").ToString & "'")
                        TEST_DEL_VOLUME = IIf(dt.Rows(i)("TEST_DEL_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("TEST_DEL_VOLUME").ToString & "'")
                        OFFLINE_QUANTITY = IIf(dt.Rows(i)("OFFLINE_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("OFFLINE_QUANTITY").ToString & "'")
                        OFFLINE_VOLUME = IIf(dt.Rows(i)("OFFLINE_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OFFLINE_VOLUME").ToString & "'")
                        OFFLINE_VALUE = IIf(dt.Rows(i)("OFFLINE_VALUE").ToString = "", "null", "'" & dt.Rows(i)("OFFLINE_VALUE").ToString & "'")
                        OFFLINE_COST = IIf(dt.Rows(i)("OFFLINE_COST").ToString = "", "null", "'" & dt.Rows(i)("OFFLINE_COST").ToString & "'")
                        OPEN_MECH_VOLUME = IIf(dt.Rows(i)("OPEN_MECH_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_MECH_VOLUME").ToString & "'")
                        CLOSE_MECH_VOLUME = IIf(dt.Rows(i)("CLOSE_MECH_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_MECH_VOLUME").ToString & "'")
                        OPEN_VOLUME_TURNOVER_CORRECTION = IIf(dt.Rows(i)("OPEN_VOLUME_TURNOVER_CORRECTION").ToString = "", "null", "'" & dt.Rows(i)("OPEN_VOLUME_TURNOVER_CORRECTION").ToString & "'")
                        OPEN_MONEY_TURNOVER_CORRECTION = IIf(dt.Rows(i)("OPEN_MONEY_TURNOVER_CORRECTION").ToString = "", "null", "'" & dt.Rows(i)("OPEN_MONEY_TURNOVER_CORRECTION").ToString & "'")

                        CLOSE_VOLUME_TURNOVER_CORRECTION = IIf(dt.Rows(i)("CLOSE_VOLUME_TURNOVER_CORRECTION").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_VOLUME_TURNOVER_CORRECTION").ToString & "'")
                        CLOSE_MONEY_TURNOVER_CORRECTION = IIf(dt.Rows(i)("CLOSE_MONEY_TURNOVER_CORRECTION").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_MONEY_TURNOVER_CORRECTION").ToString & "'")
                        OPEN_VOLUME_TURNOVER_CORRECTION2 = IIf(dt.Rows(i)("OPEN_VOLUME_TURNOVER_CORRECTION2").ToString = "", "null", "'" & dt.Rows(i)("OPEN_VOLUME_TURNOVER_CORRECTION2").ToString & "'")
                        CLOSE_VOLUME_TURNOVER_CORRECTION2 = IIf(dt.Rows(i)("CLOSE_VOLUME_TURNOVER_CORRECTION2").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_VOLUME_TURNOVER_CORRECTION2").ToString & "'")

                        Dim get_pump As String = GET_PUMP_ID(HOSE_ID, trans)
                        PUMP_ID = IIf(get_pump = "", "null", "'" & get_pump & "'")

                        Dim get_mat As String = GET_MAT_ID(HOSE_ID, trans)
                        MAT_ID = IIf(get_mat = "", "null", "'" & get_mat & "'")

                        Dim get_tank As String = GET_TANK_ID(HOSE_ID, trans)
                        TANK_ID = IIf(get_tank = "", "null", "'" & get_tank & "'")


                        sql = "INSERT INTO [dbo].[TBHOSE_HISTORY]
                               ([HOSE_ID]
                               ,[PERIOD_ID]
                               ,[OPEN_METER_VALUE]
                               ,[CLOSE_METER_VALUE]
                               ,[OPEN_METER_VOLUME]
                               ,[CLOSE_METER_VOLUME]
                               ,[POSTPAY_QUANTITY]
                               ,[POSTPAY_VALUE]
                               ,[POSTPAY_VOLUME]
                               ,[POSTPAY_COST]
                               ,[PREPAY_QUANTITY]
                               ,[PREPAY_VALUE]
                               ,[PREPAY_VOLUME]
                               ,[PREPAY_COST]
                               ,[PREPAY_REFUND_QTY]
                               ,[PREPAY_REFUND_VAL]
                               ,[PREPAY_RFD_LST_QTY]
                               ,[PREPAY_RFD_LST_VAL]
                               ,[PREAUTH_QUANTITY]
                               ,[PREAUTH_VALUE]
                               ,[PREAUTH_VOLUME]
                               ,[PREAUTH_COST]
                               ,[MONITOR_QUANTITY]
                               ,[MONITOR_VALUE]
                               ,[MONITOR_VOLUME]
                               ,[MONITOR_COST]
                               ,[DRIVEOFFS_QUANTITY]
                               ,[DRIVEOFFS_VALUE]
                               ,[DRIVEOFFS_VOLUME]
                               ,[DRIVEOFFS_COST]
                               ,[TEST_DEL_QUANTITY]
                               ,[TEST_DEL_VOLUME]
                               ,[OFFLINE_QUANTITY]
                               ,[OFFLINE_VOLUME]
                               ,[OFFLINE_VALUE]
                               ,[OFFLINE_COST]
                               ,[OPEN_MECH_VOLUME]
                               ,[CLOSE_MECH_VOLUME]
                               ,[OPEN_VOLUME_TURNOVER_CORRECTION]
                               ,[OPEN_MONEY_TURNOVER_CORRECTION]
                               ,[CLOSE_VOLUME_TURNOVER_CORRECTION]
                               ,[CLOSE_MONEY_TURNOVER_CORRECTION]
                               ,[OPEN_VOLUME_TURNOVER_CORRECTION2]
                               ,[CLOSE_VOLUME_TURNOVER_CORRECTION2]
                               ,[PUMP_ID]
                               ,[MAT_ID]
                               ,[TANK_ID])
                         VALUES
                               (" & HOSE_ID & "
                               ," & PERIOD_ID & "
                               ," & OPEN_METER_VALUE & "
                               ," & CLOSE_METER_VALUE & "
                               ," & OPEN_METER_VOLUME & "
                               ," & CLOSE_METER_VOLUME & "
                               ," & POSTPAY_QUANTITY & "
                               ," & POSTPAY_VALUE & "
                               ," & POSTPAY_VOLUME & "
                               ," & POSTPAY_COST & "
                               ," & PREPAY_QUANTITY & "
                               ," & PREPAY_VALUE & "
                               ," & PREPAY_VOLUME & "
                               ," & PREPAY_COST & "
                               ," & PREPAY_REFUND_QTY & "
                               ," & PREPAY_REFUND_VAL & "
                               ," & PREPAY_RFD_LST_QTY & "
                               ," & PREPAY_RFD_LST_VAL & "
                               ," & PREAUTH_QUANTITY & "
                               ," & PREAUTH_VALUE & "
                               ," & PREAUTH_VOLUME & "
                               ," & PREAUTH_COST & "
                               ," & MONITOR_QUANTITY & "
                               ," & MONITOR_VALUE & "
                               ," & MONITOR_VOLUME & "
                               ," & MONITOR_COST & "
                               ," & DRIVEOFFS_QUANTITY & "
                               ," & DRIVEOFFS_VALUE & " 
                               ," & DRIVEOFFS_VOLUME & "
                               ," & DRIVEOFFS_COST & "
                               ," & TEST_DEL_QUANTITY & "
                               ," & TEST_DEL_VOLUME & "
                               ," & OFFLINE_QUANTITY & "
                               ," & OFFLINE_VOLUME & "
                               ," & OFFLINE_VALUE & "
                               ," & OFFLINE_COST & "
                               ," & OPEN_MECH_VOLUME & "
                               ," & CLOSE_MECH_VOLUME & "
                               ," & OPEN_VOLUME_TURNOVER_CORRECTION & "
                               ," & OPEN_MONEY_TURNOVER_CORRECTION & "
                               ," & CLOSE_VOLUME_TURNOVER_CORRECTION & "
                               ," & CLOSE_MONEY_TURNOVER_CORRECTION & "
                               ," & OPEN_VOLUME_TURNOVER_CORRECTION2 & " 
                               ," & CLOSE_VOLUME_TURNOVER_CORRECTION2 & "
                               ," & PUMP_ID & "
                               ," & MAT_ID & "
                               ," & TANK_ID & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBMATERIAL_HISTORY"
#Region "TBMATERIAL_HISTORY"
                    Dim PERIOD_ID, MAT_ID, MAT_NAME, MAT_ID2, MAT_NAME2, MAT_NAME3, MAT_BARCODE, QTY, UOM, MOVING_AVG_PRICE, STOCK, STOCK_MIN, STOCK_MAX,
                        STOCK_LOCATION_ID, TAX_CLASS, MAT_GROUP, MAT_GROUP3, DIVISION_ID, PRICE0, PRICE1, PRICE2, PRICE3, PRICE4,
                        PRICE5, PRICE6, PRICE7, PRICE8, PRICE9, PRICE10, PRICE11, PRICE12, TIMEOFSALE, LAST_SALE, LAST_RECEIVE, BLOCK, PRICINGDATE, PRICINGMODBY,
                        LOCATION_ID, MATCOLOR, OBJ_ID, OBJ_ID_MAT_GROUP3, OBJ_ID_DIVISION_ID As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        PERIOD_ID = IIf(dt.Rows(i)("PERIOD_ID").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_ID").ToString & "'")
                        MAT_ID = IIf(dt.Rows(i)("MAT_ID").ToString = "", "null", "'" & dt.Rows(i)("MAT_ID").ToString & "'")
                        MAT_NAME = IIf(dt.Rows(i)("MAT_NAME").ToString = "", "null", "'" & dt.Rows(i)("MAT_NAME").ToString & "'")
                        MAT_ID2 = IIf(dt.Rows(i)("MAT_ID2").ToString = "", "null", "'" & dt.Rows(i)("MAT_ID2").ToString & "'")
                        MAT_NAME2 = IIf(dt.Rows(i)("MAT_NAME2").ToString = "", "null", "'" & dt.Rows(i)("MAT_NAME2").ToString & "'")
                        MAT_NAME3 = IIf(dt.Rows(i)("MAT_NAME3").ToString = "", "null", "'" & dt.Rows(i)("MAT_NAME3").ToString & "'")
                        MAT_BARCODE = IIf(dt.Rows(i)("MAT_BARCODE").ToString = "", "null", "'" & dt.Rows(i)("MAT_BARCODE").ToString & "'")
                        QTY = IIf(dt.Rows(i)("QTY").ToString = "", "null", "'" & dt.Rows(i)("QTY").ToString & "'")
                        UOM = IIf(dt.Rows(i)("UOM").ToString = "", "null", "'" & dt.Rows(i)("UOM").ToString & "'")
                        MOVING_AVG_PRICE = IIf(dt.Rows(i)("MOVING_AVG_PRICE").ToString = "", "null", "'" & dt.Rows(i)("MOVING_AVG_PRICE").ToString & "'")

                        STOCK = IIf(dt.Rows(i)("STOCK").ToString = "", "null", "'" & dt.Rows(i)("STOCK").ToString & "'")
                        STOCK_MIN = IIf(dt.Rows(i)("STOCK_MIN").ToString = "", "null", "'" & dt.Rows(i)("STOCK_MIN").ToString & "'")
                        STOCK_MAX = IIf(dt.Rows(i)("STOCK_MAX").ToString = "", "null", "'" & dt.Rows(i)("STOCK_MAX").ToString & "'")
                        STOCK_LOCATION_ID = IIf(dt.Rows(i)("STOCK_LOCATION_ID").ToString = "", "null", "'" & dt.Rows(i)("STOCK_LOCATION_ID").ToString & "'")
                        TAX_CLASS = IIf(dt.Rows(i)("TAX_CLASS").ToString = "", "null", "'" & dt.Rows(i)("TAX_CLASS").ToString & "'")
                        MAT_GROUP = IIf(dt.Rows(i)("MAT_GROUP").ToString = "", "null", "'" & dt.Rows(i)("MAT_GROUP").ToString & "'")
                        MAT_GROUP3 = IIf(dt.Rows(i)("MAT_GROUP3").ToString = "", "null", "'" & dt.Rows(i)("MAT_GROUP3").ToString & "'")
                        DIVISION_ID = IIf(dt.Rows(i)("DIVISION_ID").ToString = "", "null", "'" & dt.Rows(i)("DIVISION_ID").ToString & "'")
                        PRICE0 = IIf(dt.Rows(i)("PRICE0").ToString = "", "null", "'" & dt.Rows(i)("PRICE0").ToString & "'")
                        PRICE1 = IIf(dt.Rows(i)("PRICE1").ToString = "", "null", "'" & dt.Rows(i)("PRICE1").ToString & "'")

                        PRICE2 = IIf(dt.Rows(i)("PRICE2").ToString = "", "null", "'" & dt.Rows(i)("PRICE2").ToString & "'")
                        PRICE3 = IIf(dt.Rows(i)("PRICE3").ToString = "", "null", "'" & dt.Rows(i)("PRICE3").ToString & "'")
                        PRICE4 = IIf(dt.Rows(i)("PRICE4").ToString = "", "null", "'" & dt.Rows(i)("PRICE4").ToString & "'")
                        PRICE5 = IIf(dt.Rows(i)("PRICE5").ToString = "", "null", "'" & dt.Rows(i)("PRICE5").ToString & "'")
                        PRICE6 = IIf(dt.Rows(i)("PRICE6").ToString = "", "null", "'" & dt.Rows(i)("PRICE6").ToString & "'")
                        PRICE7 = IIf(dt.Rows(i)("PRICE7").ToString = "", "null", "'" & dt.Rows(i)("PRICE7").ToString & "'")
                        PRICE8 = IIf(dt.Rows(i)("PRICE8").ToString = "", "null", "'" & dt.Rows(i)("PRICE8").ToString & "'")
                        PRICE9 = IIf(dt.Rows(i)("PRICE9").ToString = "", "null", "'" & dt.Rows(i)("PRICE9").ToString & "'")
                        PRICE10 = IIf(dt.Rows(i)("PRICE10").ToString = "", "null", "'" & dt.Rows(i)("PRICE10").ToString & "'")
                        PRICE11 = IIf(dt.Rows(i)("PRICE11").ToString = "", "null", "'" & dt.Rows(i)("PRICE11").ToString & "'")

                        TIMEOFSALE = IIf(dt.Rows(i)("TIMEOFSALE").ToString = "", "null", "'" & dt.Rows(i)("TIMEOFSALE").ToString & "'")
                        PRICE12 = IIf(dt.Rows(i)("PRICE12").ToString = "", "null", "'" & dt.Rows(i)("PRICE12").ToString & "'")
                        LAST_SALE = IIf(dt.Rows(i)("LAST_SALE").ToString = "", "null", "" & dt.Rows(i)("LAST_SALE").ToString & "")
                        If LAST_SALE <> "null" Then
                            LAST_SALE = ConvertDate(LAST_SALE)
                        End If

                        LAST_RECEIVE = IIf(dt.Rows(i)("LAST_RECEIVE").ToString = "", "null", "" & dt.Rows(i)("LAST_RECEIVE").ToString & "")
                        If LAST_RECEIVE <> "null" Then
                            LAST_RECEIVE = ConvertDate(LAST_RECEIVE)
                        End If

                        BLOCK = IIf(dt.Rows(i)("BLOCK").ToString = "", "null", "'" & dt.Rows(i)("BLOCK").ToString & "'")
                        PRICINGDATE = IIf(dt.Rows(i)("PRICINGDATE").ToString = "", "null", "" & dt.Rows(i)("PRICINGDATE").ToString & "")
                        If PRICINGDATE <> "null" Then
                            PRICINGDATE = ConvertDate(PRICINGDATE)
                        End If

                        PRICINGMODBY = IIf(dt.Rows(i)("PRICINGMODBY").ToString = "", "null", "'" & dt.Rows(i)("PRICINGMODBY").ToString & "'")
                        LOCATION_ID = IIf(dt.Rows(i)("LOCATION_ID").ToString = "", "null", "'" & dt.Rows(i)("LOCATION_ID").ToString & "'")
                        MATCOLOR = IIf(dt.Rows(i)("MATCOLOR").ToString = "", "null", "'" & dt.Rows(i)("MATCOLOR").ToString & "'")
                        OBJ_ID = "null" 'IIf(dt.Rows(i)("OBJ_ID").ToString = "", "null", "'" & dt.Rows(i)("OBJ_ID").ToString & "'")

                        OBJ_ID_MAT_GROUP3 = "null" 'IIf(dt.Rows(i)("OBJ_ID_MAT_GROUP3").ToString = "", "null", "'" & dt.Rows(i)("OBJ_ID_MAT_GROUP3").ToString & "'")
                        OBJ_ID_DIVISION_ID = "null" 'IIf(dt.Rows(i)("OBJ_ID_DIVISION_ID").ToString = "", "null", "'" & dt.Rows(i)("OBJ_ID_DIVISION_ID").ToString & "'")



                        sql = "INSERT INTO [dbo].[TBMATERIAL_HISTORY]
                               ([PERIOD_ID]
                               ,[MAT_ID]
                               ,[MAT_NAME]
                               ,[MAT_ID2]
                               ,[MAT_NAME2]
                               ,[MAT_NAME3]
                               ,[MAT_BARCODE]
                               ,[QTY]
                               ,[UOM]
                               ,[MOVING_AVG_PRICE]
                               ,[STOCK]
                               ,[STOCK_MIN]
                               ,[STOCK_MAX]
                               ,[STOCK_LOCATION_ID]
                               ,[TAX_CLASS]
                               ,[MAT_GROUP]
                               ,[MAT_GROUP3]
                               ,[DIVISION_ID]
                               ,[PRICE0]
                               ,[PRICE1]
                               ,[PRICE2]
                               ,[PRICE3]
                               ,[PRICE4]
                               ,[PRICE5]
                               ,[PRICE6]
                               ,[PRICE7]
                               ,[PRICE8]
                               ,[PRICE9]
                               ,[PRICE10]
                               ,[PRICE11]
                               ,[PRICE12]
                               ,[TIMEOFSALE]
                               ,[LAST_SALE]
                               ,[LAST_RECEIVE]
                               ,[BLOCK]
                               ,[PRICINGDATE]
                               ,[PRICINGMODBY]
                               ,[LOCATION_ID]
                               ,[MATCOLOR]
                               ,[CREATEDATE]
                               ,[MODDATE]
                               ,[MODBY]
                               ,[OBJ_ID]
                               ,[OBJ_ID_MAT_GROUP3]
                               ,[OBJ_ID_DIVISION_ID])
                         VALUES
                               (" & PERIOD_ID & "
                               ," & MAT_ID & "
                               ," & MAT_NAME & "
                               ," & MAT_ID2 & "
                               ," & MAT_NAME2 & "
                               ," & MAT_NAME3 & "
                               ," & MAT_BARCODE & "
                               ," & QTY & "
                               ," & UOM & "
                               ," & MOVING_AVG_PRICE & "
                               ," & STOCK & "
                               ," & STOCK_MIN & "
                               ," & STOCK_MAX & "
                               ," & STOCK_LOCATION_ID & "
                               ," & TAX_CLASS & "
                               ," & MAT_GROUP & "
                               ," & MAT_GROUP3 & "
                               ," & DIVISION_ID & "
                               ," & PRICE0 & "
                               ," & PRICE1 & "
                               ," & PRICE2 & "
                               ," & PRICE3 & "
                               ," & PRICE4 & "
                               ," & PRICE5 & "
                               ," & PRICE6 & "
                               ," & PRICE7 & " 
                               ," & PRICE8 & "
                               ," & PRICE9 & "
                               ," & PRICE10 & "
                               ," & PRICE11 & "
                               ," & PRICE12 & "
                               ," & TIMEOFSALE & "
                               ," & LAST_SALE & "
                               ," & LAST_RECEIVE & "
                               ," & BLOCK & "
                               ," & PRICINGDATE & "
                               ," & PRICINGMODBY & "
                               ," & LOCATION_ID & "
                               ," & MATCOLOR & "
                               ,getdate()
                               ,getdate()
                               ,'" & modby & "'
                               ," & OBJ_ID & "
                               ," & OBJ_ID_MAT_GROUP3 & "
                               ," & OBJ_ID_DIVISION_ID & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBTANK_HISTORY"
#Region "TBTANK_HISTORY"
                    Dim PERIOD_ID, TANK_ID, OPEN_GAUGE_VOLUME, CLOSE_GAUGE_VOLUME, OPEN_THEO_VOLUME, CLOSE_THEO_VOLUME, OPEN_DIP_VOLUME, CLOSE_DIP_VOLUME, HOSE_DEL_QUANTITY,
                        HOSE_DEL_VOLUME, HOSE_DEL_VALUE, HOSE_DEL_COST, TANK_DEL_QUANTITY, TANK_DEL_VOLUME, TANK_DEL_COST, TANK_LOSS_QUANTITY, TANK_LOSS_VOLUME,
                        TANK_TRANSFER_IN_QUANTITY, TANK_TRANSFER_IN_VOLUME, TANK_TRANSFER_OUT_QUANTITY, TANK_TRANSFER_OUT_VOLUME, DIP_FUEL_TEMP, DIP_FUEL_DENSITY,
                        OPEN_DIP_WATER_VOLUME, CLOSE_DIP_WATER_VOLUME, OPEN_GAUGE_TC_VOLUME, CLOSE_GAUGE_TC_VOLUME, OPEN_WATER_VOLUME, CLOSE_WATER_VOLUME,
                        OPEN_FUEL_DENSITY, CLOSE_FUEL_DENSITY, OPEN_FUEL_TEMP, CLOSE_FUEL_TEMP, OPEN_TANK_PROBE_STATUS_ID, CLOSE_TANK_PROBE_STATUS_ID, TANK_READINGS_DT,
                        OPEN_TANK_DELIVERY_STATE_ID, CLOSE_TANK_DELIVERY_STATE_ID, OPEN_PUMP_DELIVERY_STATE, CLOSE_PUMP_DELIVERY_STATE, OPEN_DIP_TYPE_ID,
                        CLOSE_DIP_TYPE_ID, TANK_VARIANCE_REASON_ID, QUOTED_VOLUME, MAT_ID, TANK_NAME, TANK_NUMBER As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        PERIOD_ID = IIf(dt.Rows(i)("PERIOD_ID").ToString = "", "null", "'" & dt.Rows(i)("PERIOD_ID").ToString & "'")
                        TANK_ID = IIf(dt.Rows(i)("TANK_ID").ToString = "", "null", "'" & dt.Rows(i)("TANK_ID").ToString & "'")
                        OPEN_GAUGE_VOLUME = IIf(dt.Rows(i)("OPEN_GAUGE_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_GAUGE_VOLUME").ToString & "'")
                        CLOSE_GAUGE_VOLUME = IIf(dt.Rows(i)("CLOSE_GAUGE_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_GAUGE_VOLUME").ToString & "'")
                        OPEN_THEO_VOLUME = IIf(dt.Rows(i)("OPEN_THEO_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_THEO_VOLUME").ToString & "'")
                        CLOSE_THEO_VOLUME = IIf(dt.Rows(i)("CLOSE_THEO_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_THEO_VOLUME").ToString & "'")
                        OPEN_DIP_VOLUME = IIf(dt.Rows(i)("OPEN_DIP_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_DIP_VOLUME").ToString & "'")
                        CLOSE_DIP_VOLUME = IIf(dt.Rows(i)("CLOSE_DIP_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_DIP_VOLUME").ToString & "'")
                        HOSE_DEL_QUANTITY = IIf(dt.Rows(i)("HOSE_DEL_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("HOSE_DEL_QUANTITY").ToString & "'")
                        HOSE_DEL_VOLUME = IIf(dt.Rows(i)("HOSE_DEL_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("HOSE_DEL_VOLUME").ToString & "'")

                        HOSE_DEL_VALUE = IIf(dt.Rows(i)("HOSE_DEL_VALUE").ToString = "", "null", "'" & dt.Rows(i)("HOSE_DEL_VALUE").ToString & "'")
                        HOSE_DEL_COST = IIf(dt.Rows(i)("HOSE_DEL_COST").ToString = "", "null", "'" & dt.Rows(i)("HOSE_DEL_COST").ToString & "'")
                        TANK_DEL_QUANTITY = IIf(dt.Rows(i)("TANK_DEL_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("TANK_DEL_QUANTITY").ToString & "'")
                        TANK_DEL_VOLUME = IIf(dt.Rows(i)("TANK_DEL_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("TANK_DEL_VOLUME").ToString & "'")
                        TANK_DEL_COST = IIf(dt.Rows(i)("TANK_DEL_COST").ToString = "", "null", "'" & dt.Rows(i)("TANK_DEL_COST").ToString & "'")
                        TANK_LOSS_QUANTITY = IIf(dt.Rows(i)("TANK_LOSS_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("TANK_LOSS_QUANTITY").ToString & "'")
                        TANK_LOSS_VOLUME = IIf(dt.Rows(i)("TANK_LOSS_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("TANK_LOSS_VOLUME").ToString & "'")
                        TANK_TRANSFER_IN_QUANTITY = IIf(dt.Rows(i)("TANK_TRANSFER_IN_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("TANK_TRANSFER_IN_QUANTITY").ToString & "'")
                        TANK_TRANSFER_IN_VOLUME = IIf(dt.Rows(i)("TANK_TRANSFER_IN_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("TANK_TRANSFER_IN_VOLUME").ToString & "'")
                        TANK_TRANSFER_OUT_QUANTITY = IIf(dt.Rows(i)("TANK_TRANSFER_OUT_QUANTITY").ToString = "", "null", "'" & dt.Rows(i)("TANK_TRANSFER_OUT_QUANTITY").ToString & "'")

                        TANK_TRANSFER_OUT_VOLUME = IIf(dt.Rows(i)("TANK_TRANSFER_OUT_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("TANK_TRANSFER_OUT_VOLUME").ToString & "'")
                        DIP_FUEL_TEMP = IIf(dt.Rows(i)("DIP_FUEL_TEMP").ToString = "", "null", "'" & dt.Rows(i)("DIP_FUEL_TEMP").ToString & "'")
                        DIP_FUEL_DENSITY = IIf(dt.Rows(i)("DIP_FUEL_DENSITY").ToString = "", "null", "'" & dt.Rows(i)("DIP_FUEL_DENSITY").ToString & "'")
                        OPEN_DIP_WATER_VOLUME = IIf(dt.Rows(i)("OPEN_DIP_WATER_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_DIP_WATER_VOLUME").ToString & "'")
                        CLOSE_DIP_WATER_VOLUME = IIf(dt.Rows(i)("CLOSE_DIP_WATER_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_DIP_WATER_VOLUME").ToString & "'")
                        OPEN_GAUGE_TC_VOLUME = IIf(dt.Rows(i)("OPEN_GAUGE_TC_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_GAUGE_TC_VOLUME").ToString & "'")
                        CLOSE_GAUGE_TC_VOLUME = IIf(dt.Rows(i)("CLOSE_GAUGE_TC_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_GAUGE_TC_VOLUME").ToString & "'")
                        OPEN_WATER_VOLUME = IIf(dt.Rows(i)("OPEN_WATER_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("OPEN_WATER_VOLUME").ToString & "'")
                        CLOSE_WATER_VOLUME = IIf(dt.Rows(i)("CLOSE_WATER_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_WATER_VOLUME").ToString & "'")
                        OPEN_FUEL_DENSITY = IIf(dt.Rows(i)("OPEN_FUEL_DENSITY").ToString = "", "null", "'" & dt.Rows(i)("OPEN_FUEL_DENSITY").ToString & "'")

                        CLOSE_FUEL_DENSITY = IIf(dt.Rows(i)("CLOSE_FUEL_DENSITY").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_FUEL_DENSITY").ToString & "'")
                        OPEN_FUEL_TEMP = IIf(dt.Rows(i)("OPEN_FUEL_TEMP").ToString = "", "null", "'" & dt.Rows(i)("OPEN_FUEL_TEMP").ToString & "'")
                        CLOSE_FUEL_TEMP = IIf(dt.Rows(i)("CLOSE_FUEL_TEMP").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_FUEL_TEMP").ToString & "'")
                        OPEN_TANK_PROBE_STATUS_ID = IIf(dt.Rows(i)("OPEN_TANK_PROBE_STATUS_ID").ToString = "", "null", "'" & dt.Rows(i)("OPEN_TANK_PROBE_STATUS_ID").ToString & "'")
                        CLOSE_TANK_PROBE_STATUS_ID = IIf(dt.Rows(i)("CLOSE_TANK_PROBE_STATUS_ID").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_TANK_PROBE_STATUS_ID").ToString & "'")
                        TANK_READINGS_DT = IIf(dt.Rows(i)("TANK_READINGS_DT").ToString = "", "null", "" & dt.Rows(i)("TANK_READINGS_DT").ToString & "")
                        If TANK_READINGS_DT <> "null" Then
                            TANK_READINGS_DT = ConvertDate(TANK_READINGS_DT)
                        End If

                        OPEN_TANK_DELIVERY_STATE_ID = IIf(dt.Rows(i)("OPEN_TANK_DELIVERY_STATE_ID").ToString = "", "null", "'" & dt.Rows(i)("OPEN_TANK_DELIVERY_STATE_ID").ToString & "'")
                        CLOSE_TANK_DELIVERY_STATE_ID = IIf(dt.Rows(i)("CLOSE_TANK_DELIVERY_STATE_ID").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_TANK_DELIVERY_STATE_ID").ToString & "'")
                        OPEN_PUMP_DELIVERY_STATE = IIf(dt.Rows(i)("OPEN_PUMP_DELIVERY_STATE").ToString = "", "null", "'" & dt.Rows(i)("OPEN_PUMP_DELIVERY_STATE").ToString & "'")
                        CLOSE_PUMP_DELIVERY_STATE = IIf(dt.Rows(i)("CLOSE_PUMP_DELIVERY_STATE").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_PUMP_DELIVERY_STATE").ToString & "'")

                        OPEN_DIP_TYPE_ID = IIf(dt.Rows(i)("OPEN_DIP_TYPE_ID").ToString = "", "null", "'" & dt.Rows(i)("OPEN_DIP_TYPE_ID").ToString & "'")
                        CLOSE_DIP_TYPE_ID = IIf(dt.Rows(i)("CLOSE_DIP_TYPE_ID").ToString = "", "null", "'" & dt.Rows(i)("CLOSE_DIP_TYPE_ID").ToString & "'")
                        TANK_VARIANCE_REASON_ID = IIf(dt.Rows(i)("TANK_VARIANCE_REASON_ID").ToString = "", "null", "'" & dt.Rows(i)("TANK_VARIANCE_REASON_ID").ToString & "'")
                        QUOTED_VOLUME = "null" 'IIf(dt.Rows(i)("QUOTED_VOLUME").ToString = "", "null", "'" & dt.Rows(i)("QUOTED_VOLUME").ToString & "'")

                        Dim get_mat As String = GET_MAT_ID_TANK(TANK_ID, trans)
                        MAT_ID = IIf(get_mat = "", "null", "'" & get_mat & "'")

                        Dim get_tank As String = GET_TANK_NAME(TANK_ID, trans)
                        TANK_NAME = IIf(get_tank = "", "null", "'" & get_tank & "'")

                        Dim get_tank_num As String = GET_TANK_NUMBER(TANK_ID, trans)
                        TANK_NUMBER = IIf(get_tank_num = "", "null", "'" & get_tank_num & "'")


                        sql = "INSERT INTO [dbo].[TBTANK_HISTORY]
                               ([PERIOD_ID]
                               ,[TANK_ID]
                               ,[OPEN_GAUGE_VOLUME]
                               ,[CLOSE_GAUGE_VOLUME]
                               ,[OPEN_THEO_VOLUME]
                               ,[CLOSE_THEO_VOLUME]
                               ,[OPEN_DIP_VOLUME]
                               ,[CLOSE_DIP_VOLUME]
                               ,[HOSE_DEL_QUANTITY]
                               ,[HOSE_DEL_VOLUME]
                               ,[HOSE_DEL_VALUE]
                               ,[HOSE_DEL_COST]
                               ,[TANK_DEL_QUANTITY]
                               ,[TANK_DEL_VOLUME]
                               ,[TANK_DEL_COST]
                               ,[TANK_LOSS_QUANTITY]
                               ,[TANK_LOSS_VOLUME]
                               ,[TANK_TRANSFER_IN_QUANTITY]
                               ,[TANK_TRANSFER_IN_VOLUME]
                               ,[TANK_TRANSFER_OUT_QUANTITY]
                               ,[TANK_TRANSFER_OUT_VOLUME]
                               ,[DIP_FUEL_TEMP]
                               ,[DIP_FUEL_DENSITY]
                               ,[OPEN_DIP_WATER_VOLUME]
                               ,[CLOSE_DIP_WATER_VOLUME]
                               ,[OPEN_GAUGE_TC_VOLUME]
                               ,[CLOSE_GAUGE_TC_VOLUME]
                               ,[OPEN_WATER_VOLUME]
                               ,[CLOSE_WATER_VOLUME]
                               ,[OPEN_FUEL_DENSITY]
                               ,[CLOSE_FUEL_DENSITY]
                               ,[OPEN_FUEL_TEMP]
                               ,[CLOSE_FUEL_TEMP]
                               ,[OPEN_TANK_PROBE_STATUS_ID]
                               ,[CLOSE_TANK_PROBE_STATUS_ID]
                               ,[TANK_READINGS_DT]
                               ,[OPEN_TANK_DELIVERY_STATE_ID]
                               ,[CLOSE_TANK_DELIVERY_STATE_ID]
                               ,[OPEN_PUMP_DELIVERY_STATE]
                               ,[CLOSE_PUMP_DELIVERY_STATE]
                               ,[OPEN_DIP_TYPE_ID]
                               ,[CLOSE_DIP_TYPE_ID]
                               ,[TANK_VARIANCE_REASON_ID]
                               ,[QUOTED_VOLUME]
                               ,[MAT_ID]
                               ,[TANK_NAME]
                               ,[TANK_NUMBER])
                         VALUES
                               (" & PERIOD_ID & "
                               ," & TANK_ID & "
                               ," & OPEN_GAUGE_VOLUME & "
                               ," & CLOSE_GAUGE_VOLUME & "
                               ," & OPEN_THEO_VOLUME & "
                               ," & CLOSE_THEO_VOLUME & "
                               ," & OPEN_DIP_VOLUME & "
                               ," & CLOSE_DIP_VOLUME & "
                               ," & HOSE_DEL_QUANTITY & "
                               ," & HOSE_DEL_VOLUME & "
                               ," & HOSE_DEL_VALUE & "
                               ," & HOSE_DEL_COST & "
                               ," & TANK_DEL_QUANTITY & "
                               ," & TANK_DEL_VOLUME & "
                               ," & TANK_DEL_COST & "
                               ," & TANK_LOSS_QUANTITY & "
                               ," & TANK_LOSS_VOLUME & "
                               ," & TANK_TRANSFER_IN_QUANTITY & "
                               ," & TANK_TRANSFER_IN_VOLUME & "
                               ," & TANK_TRANSFER_OUT_QUANTITY & "
                               ," & TANK_TRANSFER_OUT_VOLUME & "
                               ," & DIP_FUEL_TEMP & "
                               ," & DIP_FUEL_DENSITY & "
                               ," & OPEN_DIP_WATER_VOLUME & "
                               ," & CLOSE_DIP_WATER_VOLUME & "
                               ," & OPEN_GAUGE_TC_VOLUME & "
                               ," & CLOSE_GAUGE_TC_VOLUME & "
                               ," & OPEN_WATER_VOLUME & "
                               ," & CLOSE_WATER_VOLUME & "
                               ," & OPEN_FUEL_DENSITY & "
                               ," & CLOSE_FUEL_DENSITY & "
                               ," & OPEN_FUEL_TEMP & " 
                               ," & CLOSE_FUEL_TEMP & "
                               ," & OPEN_TANK_PROBE_STATUS_ID & "
                               ," & CLOSE_TANK_PROBE_STATUS_ID & "
                               ," & TANK_READINGS_DT & "
                               ," & OPEN_TANK_DELIVERY_STATE_ID & "
                               ," & CLOSE_TANK_DELIVERY_STATE_ID & "
                               ," & OPEN_PUMP_DELIVERY_STATE & "
                               ," & CLOSE_PUMP_DELIVERY_STATE & "
                               ," & OPEN_DIP_TYPE_ID & "
                               ," & CLOSE_DIP_TYPE_ID & "
                               ," & TANK_VARIANCE_REASON_ID & "
                               ," & QUOTED_VOLUME & "
                               ," & MAT_ID & "
                               ," & TANK_NAME & "
                               ," & TANK_NUMBER & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBPAY_IN"
#Region "TBPAY_IN"
                    Dim REFBILL_NO, TRANSFER_DATE, TRAN_DATE, BUS_DATE, SHIFT_DESCRIPTION, LAST_CLOSE_SHIFT_DT,
                    FILEPATH, FILENAME, TYPE, PAYMENT_TYPE, AMOUNTREC, AMOUNT, AMOUNT_DIFF, REMARK, STATUS_SAP,
                    STATUS, NO_SALE_STATUS, CREATEDATE, CREATEBY, UPDATEDATE, UPDATEBY As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        REFBILL_NO = IIf(dt.Rows(i)("REFBILL_NO").ToString = "", "null", "'" & dt.Rows(i)("REFBILL_NO").ToString & "'")
                        TRANSFER_DATE = IIf(dt.Rows(i)("TRANSFER_DATE").ToString = "", "null", "" & dt.Rows(i)("TRANSFER_DATE").ToString & "")
                        If TRANSFER_DATE <> "null" Then
                            TRANSFER_DATE = ConvertDate(TRANSFER_DATE)
                        End If

                        TRAN_DATE = IIf(dt.Rows(i)("TRAN_DATE").ToString = "", "null", "" & dt.Rows(i)("TRAN_DATE").ToString & "")
                        If TRAN_DATE <> "null" Then
                            TRAN_DATE = ConvertDate(TRAN_DATE)
                        End If

                        BUS_DATE = IIf(dt.Rows(i)("BUS_DATE").ToString = "", "null", "" & dt.Rows(i)("BUS_DATE").ToString & "")
                        If BUS_DATE <> "null" Then
                            BUS_DATE = ConvertDate(BUS_DATE)
                        End If

                        Dim rs_shift_Desc As String = IIf(dt.Rows(i)("SHIFT_DESCRIPTION").ToString = "", "null", "'" & dt.Rows(i)("SHIFT_DESCRIPTION").ToString.Replace("$$", Chr(13)).Replace("&&", Chr(10)) & "'")
                        SHIFT_DESCRIPTION = rs_shift_Desc


                        LAST_CLOSE_SHIFT_DT = IIf(dt.Rows(i)("LAST_CLOSE_SHIFT_DT").ToString = "", "null", "" & dt.Rows(i)("LAST_CLOSE_SHIFT_DT").ToString & "")
                        If LAST_CLOSE_SHIFT_DT <> "null" Then
                            LAST_CLOSE_SHIFT_DT = ConvertDateTime(LAST_CLOSE_SHIFT_DT)
                        End If

                        FILEPATH = IIf(dt.Rows(i)("FILEPATH").ToString = "", "null", "'" & dt.Rows(i)("FILEPATH").ToString & "'")
                        FILENAME = IIf(dt.Rows(i)("FILENAME").ToString = "", "null", "'" & dt.Rows(i)("FILENAME").ToString & "'")
                        TYPE = IIf(dt.Rows(i)("TYPE").ToString = "", "null", "'" & dt.Rows(i)("TYPE").ToString & "'")
                        PAYMENT_TYPE = IIf(dt.Rows(i)("PAYMENT_TYPE").ToString = "", "null", "'" & dt.Rows(i)("PAYMENT_TYPE").ToString & "'")
                        AMOUNTREC = IIf(dt.Rows(i)("AMOUNTREC").ToString = "", "null", "'" & dt.Rows(i)("AMOUNTREC").ToString & "'")
                        AMOUNT = IIf(dt.Rows(i)("AMOUNT").ToString = "", "null", "'" & dt.Rows(i)("AMOUNT").ToString & "'")
                        AMOUNT_DIFF = IIf(dt.Rows(i)("AMOUNT_DIFF").ToString = "", "null", "'" & dt.Rows(i)("AMOUNT_DIFF").ToString & "'")

                        Dim rs_remark As String = IIf(dt.Rows(i)("REMARK").ToString = "", "null", "'" & dt.Rows(i)("REMARK").ToString.Replace("$$", Chr(13)).Replace("&&", Chr(10)) & "'")
                        REMARK = rs_remark
                        STATUS_SAP = IIf(dt.Rows(i)("STATUS_SAP").ToString = "", "null", "'" & dt.Rows(i)("STATUS_SAP").ToString & "'")
                        STATUS = IIf(dt.Rows(i)("STATUS").ToString = "", "null", "'" & dt.Rows(i)("STATUS").ToString & "'")
                        NO_SALE_STATUS = IIf(dt.Rows(i)("NO_SALE_STATUS").ToString = "", "null", "'" & dt.Rows(i)("NO_SALE_STATUS").ToString & "'")
                        CREATEDATE = IIf(dt.Rows(i)("CREATEDATE").ToString = "", "null", "" & dt.Rows(i)("CREATEDATE").ToString & "")
                        If CREATEDATE <> "null" Then
                            CREATEDATE = ConvertDateTime(CREATEDATE)
                        End If
                        CREATEBY = IIf(dt.Rows(i)("CREATEBY").ToString = "", "null", "'" & dt.Rows(i)("CREATEBY").ToString & "'")

                        UPDATEDATE = IIf(dt.Rows(i)("UPDATEDATE").ToString = "", "null", "" & dt.Rows(i)("UPDATEDATE").ToString & "")
                        If UPDATEDATE <> "null" Then
                            UPDATEDATE = ConvertDateTime(UPDATEDATE)
                        End If
                        UPDATEBY = IIf(dt.Rows(i)("UPDATEBY").ToString = "", "null", "'" & dt.Rows(i)("UPDATEBY").ToString & "'")

                        sql = "INSERT INTO [dbo].[TBPAY_IN]
                               ([REFBILL_NO]
                               ,[TRANSFER_DATE]
                               ,[TRAN_DATE]
                               ,[BUS_DATE]
                               ,[SHIFT_DESCRIPTION]
                               ,[LAST_CLOSE_SHIFT_DT]
                               ,[FILEPATH]
                               ,[FILENAME]
                               ,[TYPE]
                               ,[PAYMENT_TYPE]
                               ,[AMOUNTREC]
                               ,[AMOUNT]
                               ,[AMOUNT_DIFF]
                               ,[REMARK]
                               ,[STATUS_SAP]
                               ,[STATUS]
                               ,[NO_SALE_STATUS]
                               ,[CREATEDATE]
                               ,[CREATEBY]
                               ,[UPDATEDATE]
                               ,[UPDATEBY])
                         VALUES
                               (" & REFBILL_NO & "
                               ," & TRANSFER_DATE & "
                               ," & TRAN_DATE & "
                               ," & BUS_DATE & "
                               ," & SHIFT_DESCRIPTION & "
                               ," & LAST_CLOSE_SHIFT_DT & "
                               ," & FILEPATH & "
                               ," & FILENAME & "
                               ," & TYPE & "
                               ," & PAYMENT_TYPE & "
                               ," & AMOUNTREC & "
                               ," & AMOUNT & "
                               ," & AMOUNT_DIFF & "
                               ," & REMARK & "
                               ," & STATUS_SAP & "
                               ," & STATUS & "
                               ," & NO_SALE_STATUS & "
                               ," & CREATEDATE & "
                               ," & CREATEBY & "
                               ," & UPDATEDATE & "
                               ," & UPDATEBY & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

                Case "TBPAYIN_PERIOD_LOG"
#Region "TBPAYIN_PERIOD_LOG"
                    Dim PAYIN_ID, POS_ID, BUS_DATE, SHIFT_START, SHIFT_END, TYPE As String

                    Dim cmd As New SqlCommand
                    With cmd
                        .CommandType = CommandType.Text
                        .Connection = trans.Connection
                        .Transaction = trans
                    End With
                    For i As Integer = 0 To dt.Rows.Count - 1
                        PAYIN_ID = IIf(dt.Rows(i)("PAYIN_ID").ToString = "", "null", "'" & dt.Rows(i)("PAYIN_ID").ToString & "'")
                        POS_ID = IIf(dt.Rows(i)("POS_ID").ToString = "", "null", "" & dt.Rows(i)("POS_ID").ToString & "")

                        BUS_DATE = IIf(dt.Rows(i)("BUS_DATE").ToString = "", "null", "" & dt.Rows(i)("BUS_DATE").ToString & "")
                        If BUS_DATE <> "null" Then
                            BUS_DATE = ConvertDate(BUS_DATE)
                        End If
                        SHIFT_START = IIf(dt.Rows(i)("SHIFT_START").ToString = "", "null", "" & dt.Rows(i)("SHIFT_START").ToString & "")
                        SHIFT_END = IIf(dt.Rows(i)("SHIFT_END").ToString = "", "null", "" & dt.Rows(i)("SHIFT_END").ToString & "")
                        TYPE = IIf(dt.Rows(i)("TYPE").ToString = "", "null", "" & dt.Rows(i)("TYPE").ToString & "")

                        sql = "INSERT INTO [dbo].[TBPAYIN_PERIOD_LOG]
                               ([PAYIN_ID]
                               ,[POS_ID]
                               ,[BUS_DATE]
                               ,[SHIFT_START]
                               ,[SHIFT_END]
                               ,[TYPE])
                         VALUES
                               (" & PAYIN_ID & "
                                ," & POS_ID & "
                               ," & BUS_DATE & "
                               ," & SHIFT_START & "
                               ," & SHIFT_END & "
                               ," & TYPE & ")"

                        cmd.CommandText = sql
                        cmd.ExecuteNonQuery()
                    Next
#End Region

            End Select

            trans.Commit()
            conn.Close()
            Return ""
        Catch ex As Exception
            trans.Rollback()
            Return "พบปัญหาในการนำเข้าข้อมูล : Table " & TableName & ":" & ex.ToString & " sql: " & sql
        End Try
    End Function

    Function ConvertDate(strDate As String) As String
        Dim arr() As String = strDate.Split("/")
        Dim d As String = ""
        Dim m As String = ""
        Dim y As String = ""
        If arr.Length = 3 Then
            d = arr(0)
            m = arr(1)
            y = arr(2).Substring(0, 4)
            If CInt(y) > 2500 Then
                y = CInt(y) - 543
            End If
        End If

        Return "'" & y & "/" & m & "/" & d & "'"
    End Function

    Function ConvertDateTime(strDate As String) As String
        Dim arr_all() As String = strDate.Split(" ")
        Dim arr() As String = arr_all(0).Split("/")
        Dim d As String = ""
        Dim m As String = ""
        Dim y As String = ""
        If arr.Length = 3 Then
            d = arr(0)
            m = arr(1)
            y = arr(2).Substring(0, 4)
            If CInt(y) > 2500 Then
                y = CInt(y) - 543
            End If
        End If

        Dim hh As String = ""
        Dim mm As String = ""
        Dim ss As String = ""
        If arr_all.Length > 1 Then
            Dim arr_time() As String = arr_all(1).Split(":")
            If arr_time.Length = 3 Then
                hh = arr_time(0)
                mm = arr_time(1)
                ss = arr_time(2)
            End If

        End If

        Return "'" & y & "/" & m & "/" & d & " " & hh & ":" & mm & ":" & ss & "'"
    End Function

    Function CheckExistsSP(StoreName As String) As String
        'sp_Initial_LUBE_Stock_Inventory

        Try
            Dim sql As String = "SELECT *  From sysobjects Where id = object_id(N'[dbo].[" & StoreName & "]')  And OBJECTPROPERTY(id, N'IsProcedure') = 1 "
            Dim da As New SqlDataAdapter(sql, ConnStr)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                sql = "DROP PROCEDURE sp_Initial_LUBE_Stock_Inventory"
                Dim conn As New SqlConnection(ConnStr)
                conn.Open()
                Dim cmd As New SqlCommand
                With cmd
                    .CommandText = sql
                    .CommandType = CommandType.Text
                    .Connection = conn
                    .ExecuteNonQuery()
                End With
                conn.Close()
            End If
            Return ""

        Catch ex As Exception
            Return "พบปัญหาในการนำเข้าข้อมูล :" & ex.ToString
        End Try

    End Function


    Function CreateStoreInitialLUBE() As String
        'True = Success , False = Fail
        Try

            Dim path As String = Application.StartupPath & "\" & "Scripts\sp_Initial_LUBE_Stock_Inventory.sql"
            Dim ret As String = RunScriptSQL(path)
            Return ret

        Catch ex As Exception
            Return ex.ToString
        End Try


    End Function




#End Region

#Region "RunScript"
    Function RunScriptSQL(path As String) As String
        Dim lpcstatus_str As String = ""
        lpcstatus_str = Me.ExecScriptFile(path)

        Return lpcstatus_str
    End Function

    Function ExecScriptFile(ByVal pscript_file As String) As String
        Dim lresult_str As String = ""
        Try
            RunCommandCom("Start /min notepad """ & pscript_file & """", "", False)
            Dim fileName As String() = pscript_file.Split("\")
            Dim script = ReadTextFromNotePad(fileName(fileName.Length - 1), 2000)
            RunCommandCom("Taskkill /IM notepad.exe", "", False)

            script = Regex.Replace(script, "/\*(.|\n)*?\*/", "")
            Dim commandStrings As IEnumerable(Of String) = Regex.Split(script, "^\s*GO\s*$|^\s*GO", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
            For Each cmd As String In commandStrings
                If (cmd.Trim() <> "") Then
                    lresult_str = Me.ExecNoneQuery(cmd)
                    If (lresult_str <> "") Then
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            lresult_str = ex.ToString
        End Try

        Return lresult_str
    End Function

    Function ExecNoneQuery(ByVal psql_str As String) As String
        Dim lresult_str As String = ""
        Dim lcomm As SqlClient.SqlCommand = Nothing
        Try

            Dim conn As New SqlConnection(ConnStr)
            conn.Open()
            Dim cmd As New SqlCommand
            With cmd
                .CommandText = psql_str
                .CommandType = CommandType.Text
                .CommandTimeout = 120
                .Connection = conn
                .ExecuteNonQuery()
            End With
            conn.Close()
            lresult_str = ""

        Catch ex As Exception
            lresult_str = ex.ToString
        End Try
        Return lresult_str
    End Function


    Const WM_SETTEXT As Integer = &HC
    Const WM_GETTEXT As Integer = &HD
    Const WM_GETTEXTLENGTH As Integer = &HE
    Function ReadTextFromNotePad(fileName As String, timeOut As Integer) As String
        Dim result As String = "Time Out"
        For time = 0 To timeOut
            Dim hParent As IntPtr = FindWindowEx(IntPtr.Zero, hParent, "Notepad", fileName & " - Notepad")
            If Not hParent.Equals(IntPtr.Zero) Then
                Dim hChild As IntPtr = FindWindowEx(hParent, hChild, "Edit", vbNullString)
                If Not hChild.Equals(IntPtr.Zero) Then
                    Dim txtlen As Integer = SendMessage(hChild, WM_GETTEXTLENGTH, 0, vbNullString)
                    Dim txt As String = Space(txtlen + 1)
                    SendMessage(hChild, WM_GETTEXT, txtlen + 1, txt)
                    Return txt
                Else
                    result = "Child Window Not Found"
                End If
            Else
                result = "Main Window Not Found"
            End If
            System.Threading.Thread.Sleep(1)
        Next
        Return result
    End Function

    Sub RunCommandCom(command As String, arguments As String, permanent As Boolean)
        Dim p As Process = New Process()
        Dim pi As ProcessStartInfo = New ProcessStartInfo()
        pi.Arguments = " " + If(permanent = True, "/K", "/C") + " " + command + " " + arguments
        pi.FileName = "cmd.exe"
        pi.CreateNoWindow = True
        pi.WindowStyle = ProcessWindowStyle.Hidden
        p.StartInfo = pi
        p.Start()
    End Sub
#End Region

#Region "EncrypDecryp"

    Function Decrypt(ByVal s As String, Optional ByVal key As String = "") As String
        Dim rStr As String = ""
        Dim i As Integer
        Dim ChkSum As Byte = 12
        Try
            If key = "" Then key = "oil"
            For i = 0 To s.Length - 1
                rStr &= Chr(Asc(s(i)) - 10)
            Next
            rStr = rStr.Remove(rStr.Length - (key.Length + 1))
            rStr = rStr.Remove(0, 1)
            Return rStr
        Catch ex As Exception
            Return s
        Finally
            rStr = Nothing
            i = Nothing
            ChkSum = Nothing
        End Try
    End Function

    Function base64Encode(ByVal pstr As String, Optional ByVal Password As String = "") As String
        Dim lencData_byte(pstr.Length) As Byte
        Dim lencData As String = ""

        Try
            lencData_byte = System.Text.Encoding.UTF8.GetBytes(pstr)
            lencData = Convert.ToBase64String(lencData_byte)

        Catch ex As Exception
            '
        End Try

        Return lencData

    End Function


#End Region


End Class
