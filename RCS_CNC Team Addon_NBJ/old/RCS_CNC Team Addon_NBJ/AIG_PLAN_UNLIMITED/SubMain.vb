Option Strict Off
Option Explicit On
Module SubMain
	'//  SAP MANAGE UI API 6.5 SDK Sample
	'//****************************************************************************
	'//
	'//  File:      SubMain.bas
	'//
	'//  Copyright (c) SAP MANAGE
	'//
	'// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
	'// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
	'// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
	'// PARTICULAR PURPOSE.
	'//
	'//****************************************************************************
	

	' Declare the global company variable
	Public vCmp As SAPbobsCOM.Company '// To notify applications the printer has changed
	
	Private Const HWND_BROADCAST As Integer = &HFFFFs
	Private Const WM_WININICHANGE As Integer = &H1As
	
    Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Integer
    Public oDBDataSource As SAPbouiCOM.DBDataSource
    Public oDBDataSource2 As SAPbouiCOM.DBDataSource
    Public oDBDataSourceJDT1 As SAPbouiCOM.DBDataSource
    Public oDBDataSourceOHEM As SAPbouiCOM.DBDataSource

    '// declaring a User data source for the "Remarks" Column
    Public oUserDataSource As SAPbouiCOM.UserDataSource
    Public oUserDataSource_03 As SAPbouiCOM.UserDataSource
    Public oUserDataSource_04 As SAPbouiCOM.UserDataSource
    Public oUserDataSource_05 As SAPbouiCOM.UserDataSource
    Public oUserDataSource_LV As SAPbouiCOM.UserDataSource
    Public oUserDataSource_UV As SAPbouiCOM.UserDataSource
    Public oUserDataSource_BBLV As SAPbouiCOM.UserDataSource
    Public oUserDataSourceOHEM As SAPbouiCOM.UserDataSource
    ' Private oForm As SAPbouiCOM.Form
    Public bModal As Boolean = False
    Public strSearchCol As String
    Public intSearchRow As Integer
    Public bolSearchItem As Boolean = False
    Public bolSearchOrder As Boolean = False
    Public bolSearchMA As Boolean = False


	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
    'Private Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
    'Public Graphform As AIG_GraphForm

	Private Sub SetDefaultPrinter(ByVal PrinterName As String, ByVal DriverName As String, ByVal PrinterPort As String)
        'On Error GoTo Err_Proc
        Dim DeviceLine As String
        Try
            'rebuild a valid device line string
            DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort

            'Store the new printer information in the [WINDOWS] section of the WIN.INI file for the DEVICE= item
            WriteProfileString("windows", "Device", DeviceLine)
            'Cause all applications to reload the INI file
            '	Call SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")

        Catch ex As Exception
            'Exit_Proc:  
            'MsgBox("Error in basPrinter (SetDefaultPrinter)" & vbCrLf & Err.Description, MsgBoxStyle.Critical)
            Exit Sub
        End Try


	End Sub

    Public Sub Main(ByVal args() As String)
        'System.Threading.Thread.Sleep(3000) 'For at sikre at andre addon har startet op.

        'System.Diagnostics.Debugger.Launch()

        ' Init the company object (use ProgID first to avoid hardcoded CLSID/version mismatch)
        Try
            vCmp = CType(CreateObject("SAPbobsCOM.Company"), SAPbobsCOM.Company)
        Catch
            vCmp = New SAPbobsCOM.Company
        End Try

        If Not Connect(vCmp) Then
            Dim errCode As Integer = 0
            Dim errMsg As String = String.Empty
            Try
                vCmp.GetLastError(errCode, errMsg)
            Catch
            End Try

            Console.WriteLine($"SAP DI connection failed. Code: {errCode}. Message: {errMsg}")
            Exit Sub
        End If

        If args.Length < 4 Then
            ErrorLog("Main", "Missing arguments")
            Console.WriteLine("Missing arguments")
            Exit Sub
        End If

        'Dim path As String = Windows.Forms.Application.StartupPath() + "\PLAN_UNLIMITED__TEST.txt"
        'Dim fs As FileStream = File.Create(path)
        'Dim info As Byte() = New Text.UTF8Encoding(True).GetBytes(args(0))
        'fs.Write(info, 0, info.Length)
        'fs.Close()

        'Console.WriteLine(args(0))
        'If Not SetConnectionContext(args(0)) = 0 Then
        '    SBO_Application.MessageBox("Failed setting a connection to DI API")
        '    End ' Terminating the Add-On Application
        'End If

        'If Not ConnectToCompany() = 0 Then
        '    SBO_Application.MessageBox("Failed connecting to the company's Database")
        '    End ' Terminating the Add-On Application
        'End If


        '// Creating an object
        Dim oChangeSysForm As ChangeSysForm


        oChangeSysForm = New ChangeSysForm(args(0), args(1), args(2), args(3), args(4))
        'oChangeSysForm = New ChangeSysForm(Boolean.Parse(args(1)), args(2), args(3), args(4), args(5))

        'oChangeSysForm = New ChangeSysForm("BtnPUL", "4434", "2", "-1", "manager") 'TEST
        'oChangeSysForm.PlanUnlimeted = Boolean.Parse(args(0))
        'oChangeSysForm.SalesOrderDocEntry = CLng(args(1))

        'AIG_PLAN_UNLIMITED.exe True 4434 0 manager
        Exit Sub

    End Sub

    Private Function Connect(ByRef vCmp As SAPbobsCOM.Company) As Boolean
        Dim i As Integer
        vCmp.CompanyDB = My.Settings.CompanyDB
        vCmp.DbServerType = My.Settings.DbServerType
        vCmp.LicenseServer = My.Settings.LicenseServer
        vCmp.Server = My.Settings.Server
        If Len(My.Settings.UserName) > 0 And Len(My.Settings.Password) > 0 Then
            vCmp.UserName = My.Settings.UserName
            vCmp.Password = My.Settings.Password
        Else
            vCmp.UserName = My.Settings.SBOUserName
            vCmp.Password = My.Settings.SBOPassword
        End If
        vCmp.UserName = My.Settings.SBOUserName
        vCmp.Password = My.Settings.SBOPassword
        vCmp.UseTrusted = My.Settings.UseTrusted

        i = vCmp.Connect()

        Return vCmp.Connected
    End Function


    'Private Function SetConnectionContext(sConnectionContext As String) As Integer

    '    '// Initialize the Company object
    '    'vCmp = New SAPbobsCOM.Company

    '    '// Set the connection context information to the DI API.
    '    SetConnectionContext = vCmp.SetSboLoginContext(sConnectionContext)

    'End Function

    'Private Function ConnectToCompany() As Integer

    '    '// Establish the connection to the company database.
    '    ConnectToCompany = vCmp.Connect

    'End Function




    'Function VBAWeekNum(ByRef D As Date, ByRef FW As Short) As Short
    '    VBAWeekNum = CShort(VB6.Format(D, "ww", , FW))
    'End Function

    'Function VBAWeekD(ByRef D As Date, ByRef FW As Short) As String

    '    Select Case Weekday(D) 'Begins new nested decision structure
    '        'Will print according to what day of week it is
    '        'WeekDay(Now) can read day, date, time from user computer
    '    Case FirstDayOfWeek.Sunday
    '            VBAWeekD = "Søndag"
    '        Case FirstDayOfWeek.Monday
    '            VBAWeekD = "Mandag"
    '        Case FirstDayOfWeek.Tuesday
    '            VBAWeekD = "Tirsdag"
    '        Case FirstDayOfWeek.Wednesday
    '            VBAWeekD = "Onsdag"
    '        Case FirstDayOfWeek.Thursday
    '            VBAWeekD = "Torsdag"
    '        Case FirstDayOfWeek.Friday
    '            VBAWeekD = "Fredag"
    '        Case FirstDayOfWeek.Saturday
    '            VBAWeekD = "Lørdag"
    '    End Select 'Must end the "inner" Select decision structure
    '    'Return VBAWeekD
    'End Function

    Public Function ConnectAdm() As Integer
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'The folowing lines show how to establish a connection to a company
        ' Pay attention to the "UseTrusted" property when using a workstation not connected to a domain
        ' You must change the connection properties to mach the real ones
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        '      Dim l As Integer

        'vCmp.Disconnect()

        ''vCmp.UseTrusted = True
        'vCmp.Server = SBO_Application.Company.ServerName '"(local)"            'Change _SERVER_ to your real server name
        'vCmp.CompanyDB = SBO_Application.Company.DatabaseName '"SBODemo_DK" 'SelectCompanyDb    'Choose database from list
        'vCmp.UserName = "manager"
        'vCmp.Password = "manager"
        'vCmp.Language = SBO_Application.Language 'BoSuppLangs.ln_Danish

        'l = vCmp.Connect()

        'If l <> 0 Then

        '	vCmp.UseTrusted = True
        '	vCmp.Server = SBO_Application.Company.ServerName '"(local)"            'Change _SERVER_ to your real server name
        '	vCmp.CompanyDB = SBO_Application.Company.DatabaseName '"SBODemo_DK" 'SelectCompanyDb    'Choose database from list
        '	vCmp.UserName = "manager"
        '	vCmp.Password = "manager"
        '	vCmp.Language = SBO_Application.Language 'BoSuppLangs.ln_Danish

        '	l = vCmp.Connect()

        'End If

        'ConnectAdm = l
        Dim Id As String
        If 1 = 2 Then
            Id = SBO_Application.Company.InstallationId
        End If
        '       Id = "0020185678" Or _
        '       Id = "0020193375" Or _
        '       Id = "0020155342" Or _ Athena
        '       Id = "0020212341" Or _
        '       Id = "0020198791" Or _
        '       Id = "0020226106" Or _ Budstikken
        '       Id = "0020162055" or _
        '       Id = "0020198262" Or _ Lars Ole Wiese
        '       Id = "0020233119" Or _ DCT
        '       Id = "0020186854" Then PJC Systemudvikling Aps  
        '       Id = "0020163673" Then BEMA
        '       ID = "0020198791" Gammelbro
        '       ID = "0020247695" Hugo Jørgensen

        'If Id = "0020185678" Or _
        '   Id = "0020193375" Or _
        '   Id = "0020155342" Or _
        '   Id = "0020212341" Or _
        '   Id = "0020198791" Or _
        '   Id = "0020226106" Or _
        '   Id = "0020162055" Or _
        '   Id = "0020198262" Or _
        '   Id = "0020233119" Or _
        '   Id = "0020186854" Or _
        '   Id = "0020198791" Or _
        '   Id = "0020247695" Or _
        '   Id = "0020163673" Then
        '    SBO_Application.StatusBar.SetText(TL("Licens ok") + "!", _
        '                                                              SAPbouiCOM.BoMessageTime.bmt_Short, _
        '                                                              SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        'Else
        '    SBO_Application.MessageBox(TL("Der er ikke licens til denne installation") + "!")
        '    End
        'End If


        Dim l As Integer
        'Dim Id As String
        Try
            vCmp.Disconnect()
        Catch ex As Exception

        End Try
        Dim lResult As Long
        Dim sCookie As String 'Cookie String
        Dim sConStr As String 'ConnectionString


        'Hier Single Sign On-Prozedur:
        '1. Cookie holen
        vCmp = New SAPbobsCOM.Company
        vCmp.Server = SBO_Application.Company.ServerName

        vCmp.language = SBO_Application.Language
        sCookie = vCmp.GetContextCookie
        '2. Context Infos von der Application holen:
        sConStr = SBO_Application.Company.GetConnectionContext(sCookie)
        '3. Connection evtl. beenden:
        If vCmp.Connected = True Then
            vCmp.Disconnect()
        End If
        '4. set Login-Context and connect:
        'lResult = vCmp.SetSboLoginContext(sConStr)
        Dim i As Integer
        '4. set Login-Context and connect:
        lResult = vCmp.SetSboLoginContext(sConStr)

        lResult = vCmp.Connect()

        Do Until lResult = 0 Or i > 5
            i = i + 1
            AIG_Logging.PrintToDebugLog("SubMain.Connect Loop: " + i.ToString)
            ''4. set Login-Context and connect:
            lResult = vCmp.SetSboLoginContext(sConStr)

            lResult = vCmp.Connect()
            If lResult <> 0 Then
                vCmp = Nothing
                System.GC.Collect()
                vCmp = New SAPbobsCOM.Company
                vCmp.Server = SBO_Application.Company.ServerName
                vCmp.language = SBO_Application.Language
                sCookie = vCmp.GetContextCookie
                '2. Context Infos von der Application holen:
                sConStr = SBO_Application.Company.GetConnectionContext(sCookie)
            End If
        Loop


        'If lResult = 0 Then
        '    lResult = vCmp.Connect()
        '    If lResult = 0 Then
        '        'SBO_Application.MessageBox("Single Sign On erfolgreich!")
        '    Else
        '        System.Threading.Thread.Sleep(5000)
        '        lResult = vCmp.Connect()
        '        If lResult = 0 Then
        '            'SBO_Application.MessageBox("Single Sign On erfolgreich!")
        '        Else
        '            System.Threading.Thread.Sleep(5000)
        '            lResult = vCmp.Connect()
        '            If lResult = 0 Then
        '                'SBO_Application.MessageBox("Single Sign On erfolgreich!")
        '            Else
        '                SBO_Application.MessageBox("Connect failed for KOMGÅ")
        '            End If
        '        End If

        '    End If
        'Else
        '    SBO_Application.MessageBox("SingleSignOn fehlgeschlagen")
        'End If


        l = lResult
        ConnectAdm = l


        If l <> 0 Then
            SBO_Application.MessageBox(TL("Failed to connect to company") + "! " + TL("Fejlkode") + " = " + l.ToString())
        Else
            SBO_Application.StatusBar.SetText(TL("Forbundet til company ok") + "!", _
                                                          SAPbouiCOM.BoMessageTime.bmt_Short, _
                                                          SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        End If




    End Function
End Module