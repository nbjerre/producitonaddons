
Imports System.Collections
Imports System.Diagnostics
Imports System.Management

Module AIG_STANDARD

    '//**********************************************
    '// I dette modul kan lægges de standardfunktioner som benyttes igen og igen
    '//**********************************************
    Dim test As Object

    Public NameCol As System.Collections.Specialized.NameValueCollection
    Public rsGetNextRecord As SAPbobsCOM.IRecordset
    Public rsGetNextNum As SAPbobsCOM.IRecordset


    Public Sub killProcessByNameForOwner(ByVal ProcessName As String)

        Dim selectQuery As System.Management.SelectQuery = New SelectQuery("select * from Win32_Process where name like '" + ProcessName + "%'")
        Dim searcher As ManagementObjectSearcher = New ManagementObjectSearcher(selectQuery)
        Dim proc As Management.ManagementObject

        For Each proc In searcher.Get()
            Dim ownerData(2) As String

            proc.InvokeMethod("GetOwner", CType(ownerData, Object()))
            '----------------------------------------
            Dim proName As String = proc("Name").ToString.ToLower
            Dim userName As String = ownerData(0)
            Dim userDomain As String = ownerData(1)
            '----------------------------------------
            If (userName = System.Environment.UserName) Then
                proc.InvokeMethod("Terminate", Nothing)
            End If

        Next

    End Sub


    Public Sub killProcessByNameForAllUsers(ByVal ProcessName As String)
        Dim Process As Array
        Dim p As System.Diagnostics.Process
        Process = System.Diagnostics.Process.GetProcesses
        For Each p In Process
            If p.ProcessName.ToLower = ProcessName.ToLower Then
                p.Kill()
            End If
        Next
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns>Returns DateTimeFormatInfo for formatting string to date</returns>
    ''' <remarks>As only strings can be extracted from SBO use theis format for converting
    ''' strings from SBO to date. 
    ''' Excample:
    '''     Dim DTF As System.Globalization.DateTimeFormatInfo = GetSboDateFormat()
    '''     strFDate = oForm.DataSources.UserDataSources.Item("FDate").Value
    '''     dFDate = DateTime.Parse(strFDate, DTF)
    '''     strFDate = dFDate.ToString("MM-dd-yyyy")
    ''' </remarks>
    Public Function GetSboDateFormat() As System.Globalization.DateTimeFormatInfo
        Dim FormatString As String = ""
        Dim DateSep As String = ""
        Dim sql As String = "SELECT [DateFormat], [DateSep] FROM OADM"
        Dim rs As SAPbobsCOM.Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        rs.DoQuery(sql)

        Dim DTF As System.Globalization.DateTimeFormatInfo
        Dim cinfo As System.Globalization.CultureInfo = GetSBOCultureInfo()

        If Not cinfo Is Nothing Then
            DTF = cinfo.DateTimeFormat
        Else
            DTF = New System.Globalization.DateTimeFormatInfo
        End If

        DateSep = rs.Fields.Item("DateSep").Value

        Select Case rs.Fields.Item("DateFormat").Value
            Case "0"
                FormatString = "dd" + DateSep + "MM" + DateSep + "yy"
            Case "1"
                FormatString = "dd" + DateSep + "MM" + DateSep + "yyyy"
            Case "2"
                FormatString = "MM" + DateSep + "dd" + DateSep + "yy"
            Case "3"
                FormatString = "MM" + DateSep + "dd" + DateSep + "yyyy"
            Case "4"
                FormatString = "yyyy" + DateSep + "MM" + DateSep + "dd"
            Case "5"
                FormatString = "dd" + DateSep + "mmmm" + DateSep + "yyyy"
        End Select

        DTF.ShortDatePattern = FormatString
        DTF.DateSeparator = DateSep

        Return DTF
    End Function

    Public Function GetSBOCultureInfo() As System.Globalization.CultureInfo
        Dim cinfo As System.Globalization.CultureInfo
        Dim sboLanguage As String = ""
        Select Case SBO_Application.Language
            Case "1" ' ln_Hebrew 1 Hebrew (Israel). 
                sboLanguage = "he-IL"
            Case "2" ' ln_Spanish_Ar 2 Spanish (Argentina). 
                sboLanguage = "es-AR"
            Case "3" ' ln_English 3 English (United States). 
                sboLanguage = "en-US"
            Case "5" ' ln_Polish 5 Polish. 
                sboLanguage = "pl-PL"
            Case "6" ' ln_English_Sg 6 English (Singapore). 
                sboLanguage = "en-GB"
            Case "7" ' ln_Spanish_Pa 7 Spanish (Panama). 
                sboLanguage = "es-PA"
            Case "8" ' ln_English_Gb 8 English (United Kingdom). 
                sboLanguage = "en-GB"
            Case "9" ' ln_German 9 German. 
                sboLanguage = "de-DE"
            Case "10" ' ln_Serbian 10 Serbian. 
                sboLanguage = "sr-Latn-CS"
            Case "11" ' ln_Danish 11 Danish (Denmark). 
                sboLanguage = "da-DK"
            Case "12" ' ln_Norwegian 12 Norwegian (Norway). 
                sboLanguage = "nn-NO"
            Case "13" ' ln_Italian 13 Italian. . 
                sboLanguage = "it-IT"
            Case "14" ' ln_Hungarian 14 Hungarian.
                sboLanguage = "hu-HU"
            Case "15" ' ln_Chinese 15 Chinese.
                sboLanguage = "zh-CN"
            Case "16" ' ln_Dutch 16 Dutch (Netherlands). 
                sboLanguage = "nl-NL"
            Case "17" ' ln_Finnish 17 Finnish (Finland). 
                sboLanguage = "fi-FI"
            Case "18" ' ln_Greek 18 Greek. 
                sboLanguage = "el-GR"
            Case "19" ' ln_Portuguese 19 Portuguese. 
                sboLanguage = "pt-PT"
            Case "20" ' ln_Swedish 20 Swedish 
                sboLanguage = "sv-SE"
            Case "21" ' ln_English_Cy 21 English. 
                sboLanguage = "en-GB"
            Case "22" ' ln_French 22 French. 
                sboLanguage = "fr-FR"
            Case "23" ' ln_Spanish 23 Spanish. 
                sboLanguage = "es-ES"
            Case "24" ' ln_Russian 24 Russian. 
                sboLanguage = "ru-RU"
            Case "25" ' ln_Spanish_La 25 Spanish (Latin America). 
                sboLanguage = "es-NI"
            Case "26" ' ln_Czech_Cz 26 Czech. 
                sboLanguage = "cs-CZ"
            Case "27" ' ln_Slovak_Sk 27 Slovak. 
                sboLanguage = "sk-SK"
            Case "28" ' ln_Korean_Kr 28 Korean. 
                sboLanguage = "ko-KR"
            Case "29" ' ln_Portuguese_Br 29 Portuguese (Brazil). 
                sboLanguage = "pt-BR"
            Case "30" ' ln_Japanese_Jp 30 Japanese. 
                sboLanguage = "ja-JP"
            Case "31" ' ln_Turkish_Tr 31 Turkish. 
                sboLanguage = "tr-TR"
            Case "35" ' ln_TrdtnlChinese_Hk 35 Traditional Chinese (Hong Kong). 
                sboLanguage = "zh-HK"
        End Select

        If sboLanguage <> "" Then
            cinfo = New System.Globalization.CultureInfo(sboLanguage)
        Else
            cinfo = Nothing
        End If

        Return cinfo
    End Function

    Public Sub testInterFaceVersion()

        'test = vCmp.Version
        ' test = ""

        ' test = SAPbobsCOM.AdminInfo
    End Sub


    ''' <summary>
    ''' Åbner en fildialog
    ''' </summary>
    ''' <returns>Valgt fil</returns>
    ''' <remarks>Klassen ForegroundWindow skal være tilgængelig.
    ''' Virker ikke i debugmode, da VS vil være ForegroundWindow. 
    ''' Test evt. med en genvejstast.</remarks>
    Public Function AIG_OpenFileDialog(ByVal Filter As String, ByVal InitialDirectory As String) As String
        Dim File As String = ""
        Dim myDialog As Windows.Forms.OpenFileDialog
        Try
            Dim sapForegrund As Windows.Forms.IWin32Window
            sapForegrund = ForegroundWindow.Instance
            myDialog = New Windows.Forms.OpenFileDialog
            myDialog.Filter = Filter
            myDialog.InitialDirectory = InitialDirectory

            If myDialog.ShowDialog(sapForegrund) = DialogResult.OK Then
                File = myDialog.FileName
            Else
                File = ""
            End If
        Catch ex As Exception
            Dim test As String = ex.Message
            ErrorLog("AIG_OpenFileDialog", ex.Message)
        End Try

        Return File
    End Function

    ''' <summary>
    ''' Klassen ForgroundWindow benyttes til filedialog mm. for at 
    ''' sætte SBO vinduet til øverste window
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ForegroundWindow
        'Inherits System.Windows.Forms.IWin32Window
        Implements System.Windows.Forms.IWin32Window
        Private Shared _window As ForegroundWindow = New ForegroundWindow

        Private Sub New()
            MyBase.new()
        End Sub

        Public Shared ReadOnly Property Instance() As IWin32Window
            Get
                Return _window
            End Get
        End Property

        ReadOnly Property IWin32Window_Handle() As IntPtr Implements IWin32Window.Handle
            Get
                Dim iHand As Integer = GetForegroundWindow
                Return iHand
            End Get
        End Property

        Private Declare Function GetForegroundWindow Lib "user32.dll" () As IntPtr

    End Class


    Public Sub OpretForespørgsel(ByVal Navn As String, ByVal SqlStreng As String, ByVal Kategori As Integer)
        Dim oQuery As SAPbobsCOM.IUserQueries = Nothing
        Dim Value As Integer
        Try
            oQuery = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries)

            oQuery.QueryCategory = Kategori
            oQuery.QueryDescription = Navn
            oQuery.Query = SqlStreng

            Value = oQuery.Add()
            Dim intKey As Integer = oQuery.InternalKey
        Catch ex As Exception
            test = ex.Message
        End Try

    End Sub

    Public Sub ReadTranslateFile(ByVal File As String)

        Dim objReader As System.IO.StreamReader = Nothing
        Try


            NameCol = New System.Collections.Specialized.NameValueCollection
            Dim FileArray As String()
            Dim StringArray As String()
            Dim Path, Key As String
            Dim s, q As String
            Dim b As Integer
            Dim a As Integer

            'Dim objwriter As System.IO.StreamWriter = Nothing

            Path = (System.Windows.Forms.Application.StartupPath() + "\" + File) 'WSTranslate.txt")
            objReader = New System.IO.StreamReader(Path, Text.Encoding.UTF7)
            FileArray = objReader.ReadToEnd.Split(vbCrLf)
            'FileArray = My.Resources.AIG_Translate.Translate.Split(vbCrLf)

            For Each s In FileArray
                StringArray = s.Split(vbTab)
                q = ""
                If b = 950 Then
                    b = 950
                End If
                b += 1
                For a = 1 To StringArray.Length - 1

                    If q <> "" Then
                        q = q + vbTab
                    End If
                    q = q + StringArray.GetValue(a)
                Next
                Key = StringArray.GetValue(0)
                Key = Key.Trim(vbCr)
                Key = Key.Trim(vbLf)
                Key = Key.Trim(vbTab)

                NameCol.Add(Key, q)
            Next
        Catch ex As Exception
            test = ex.Message

        Finally
            objReader.Close()
            objReader.Dispose()
        End Try


    End Sub


    Public Sub AIG_ResourceManager()
        Dim FileArray As String()
        'Dim SYSA As System.Array
        Dim StringArray As String()
        'Dim TransString As String
        Dim s, q As String
        Dim a As Integer


        NameCol = New System.Collections.Specialized.NameValueCollection

        ' FileArray = My.Resources.AIG_Translate.Translate.Split(vbCrLf)

        For Each s In FileArray

            StringArray = s.Split(vbTab)

            q = ""
            For a = 1 To StringArray.Length - 1
                If q <> "" Then
                    q = q + vbTab
                End If
                q = q + StringArray.GetValue(a)
            Next

            NameCol.Add(StringArray.GetValue(0), q)

        Next



    End Sub

    Public Function TL(ByVal DK_Tekst As String) As String
        Dim tekst As String
        Dim TxtArray As String()
        Dim Key As String
        tekst = ""
        Dim Language As String

        Try


            Language = 11 'SBO_Application.Language

            Key = DK_Tekst.Replace(vbCrLf, "")
            'Key = Key.Replace("æ", "")
            'Key = Key.Replace("ø", "")
            'Key = Key.Replace("å", "")
            'Key = Key.Replace("Æ", "")
            'Key = Key.Replace("Ø", "")
            'Key = Key.Replace("Å", "")

            'If Language = "11" Then 'SAPbobsCOM.BoSuppLangs.ln_Danish Then
            '    ConvertCode = "105"
            '    Return DK_Tekst
            '    Exit Function
            'End If

            tekst = NameCol.Get(Key)
            If Key = "Hverdag5" Then
                Key = "Hverdag5"
            End If
            If tekst = "" Or tekst Is Nothing Then
                Dim Path As String = (System.Windows.Forms.Application.StartupPath() + "\MissingLink.txt")
                Dim SW As StreamWriter = New StreamWriter(Path, True)
                SW.WriteLine(DK_Tekst)
                SW.Flush()
                SW.Close()
                SW.Dispose()

                Return DK_Tekst
                Exit Function
            End If
            TxtArray = tekst.Split(vbTab)

            Select Case Language
                Case "11" 'DK
                    ConvertCode = "105"
                    tekst = DK_Tekst
                Case "8" 'GB
                    tekst = TxtArray.GetValue(0)
                    If tekst = "" Then
                        tekst = DK_Tekst
                    End If
                Case "9" 'DE
                    tekst = TxtArray.GetValue(1)
                    If tekst = "" Then
                        tekst = TxtArray.GetValue(0)
                    End If
                    If tekst = "" Then
                        tekst = DK_Tekst
                    End If
                Case "22" 'FR
                    tekst = TxtArray.GetValue(2)
                    If tekst = "" Then
                        tekst = TxtArray.GetValue(0)
                    End If
                    If tekst = "" Then
                        tekst = DK_Tekst
                    End If
                Case "12" 'NO
                    tekst = TxtArray.GetValue(3)
                    If tekst = "" Then
                        tekst = TxtArray.GetValue(0)
                    End If
                    If tekst = "" Then
                        tekst = DK_Tekst
                    End If
                Case "20" 'SE
                    ConvertCode = "120"
                    tekst = TxtArray.GetValue(4)
                    If tekst = "" Then
                        tekst = TxtArray.GetValue(0)
                    End If
                    If tekst = "" Then
                        tekst = DK_Tekst
                    End If
                Case "23" 'ES
                    tekst = TxtArray.GetValue(5)
                    If tekst = "" Then
                        tekst = TxtArray.GetValue(0)
                    End If
                    If tekst = "" Then
                        tekst = DK_Tekst
                    End If
            End Select

            'test = system.Text.Encoding.Eq

        Catch ex As Exception
            tekst = ""
        End Try
        Return tekst
    End Function


    Public Function Translate(ByVal DK_Tekst As String) As String
        Dim tekst As String
        Dim TxtArray As String()
        Dim Key As String

        Dim Language As String

        Language = SBO_Application.Language

        Key = DK_Tekst.Replace(vbCrLf, "")
        Key = Key.Replace("æ", "")
        Key = Key.Replace("ø", "")
        Key = Key.Replace("å", "")
        Key = Key.Replace("Æ", "")
        Key = Key.Replace("Ø", "")
        Key = Key.Replace("Å", "")

        If Language = "11" Then 'SAPbobsCOM.BoSuppLangs.ln_Danish Then
            Return DK_Tekst
            Exit Function
        End If

        tekst = NameCol.Get(Key)
        If tekst = "" Then
            Return DK_Tekst
            Exit Function
        End If
        TxtArray = tekst.Split(vbTab)

        Select Case Language
            Case "8" 'GB
                tekst = TxtArray.GetValue(0)
            Case "9" 'DE
                tekst = TxtArray.GetValue(1)
            Case "22" 'FR
                tekst = TxtArray.GetValue(2)
            Case "12" 'NO
                tekst = TxtArray.GetValue(3)
            Case "20" 'SE
                tekst = TxtArray.GetValue(4)
            Case "23" 'SE
                tekst = TxtArray.GetValue(5)
        End Select


        Return tekst
    End Function


    ''' <summary>
    ''' Oversætter dansk tekst til SBO_Application.Language
    ''' </summary>
    ''' <param name="DK_Tekst"></param>
    ''' <returns>Tekst</returns>
    ''' <remarks>Kræver at der findes en oversættelse fra dansk til det pågældende sprog i
    ''' tabellen [@AIG_TRANSLATE], ellers returneres den danske tekst</remarks>
    Public Function Translate_old(ByVal DK_Tekst As String) As String
        Dim Tekst As String = ""
        Dim sql As String = ""
        Dim Sprog As String = ""
        Dim kode As String
        Dim rs As SAPbobsCOM.Recordset
        ' Sprog = SBO_Application.Language.GetName
        Try
            kode = vCmp.language
            Select Case kode
                Case "11"
                    Sprog = "DK"
                Case "9"
                    Sprog = "DE"
                Case "8"
                    Sprog = "GB"
                Case "22"
                    Sprog = "FR"
                Case "12"
                    Sprog = "NO"
                Case "20"
                    Sprog = "SE"
            End Select

            If Sprog = "DK" Then
                Tekst = DK_Tekst
            Else
                sql = "Select U_" + Sprog + " from [@AIG_TRANSLATE] Where U_DK = '" + DK_Tekst + "'"
                rs = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rs.DoQuery(sql)
                Tekst = rs.Fields.Item("U_" + Sprog).Value.ToString
                If Tekst = "" Then
                    Tekst = DK_Tekst
                End If
            End If
        Catch ex As Exception
            test = ex.Message
        End Try

        Return Tekst
    End Function

    ''' <summary>
    ''' Tjek om formen findes som XML
    ''' </summary>
    ''' <param name="FormType"></param>
    ''' <returns></returns>
    ''' <remarks>Bruges før opretning af forms. 
    ''' Hvis formen findes som XML, behøves den ikke at kreeres igen</remarks>
    Public Function IsFormXML(ByVal FormType As String) As Boolean
        Dim sPath As String
        Dim YES As Boolean = False
        sPath = IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString
        If IO.File.Exists(sPath & "\" & FormType & ".xml") Then
            YES = True
        End If
        Return YES
    End Function

    ''' <summary>
    ''' Hvis formen findes som XML kan den kaldes med denne metode
    ''' </summary>
    ''' <param name="FormType"></param>
    ''' <remarks></remarks>
    Public Sub LoadFromXML(ByVal FormType As String)
        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        '// load the content of the XML File
        Dim sPath As String
        sPath = IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString
        oXmlDoc.Load(sPath & "\" & FormType & ".xml")
        '// load the form to the SBO application in one batch
        SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
    End Sub

    ''' <summary>
    ''' Gem en form som XML efter kreation
    ''' </summary>
    ''' <param name="Form"></param>
    ''' <remarks></remarks>
    Public Sub SaveAsXML(ByRef Form As SAPbouiCOM.Form)
        Dim oXmlDoc As Xml.XmlDocument
        Dim sXmlString As String
        Dim FormType As String
        Dim sPath As String
        sPath = IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString
        FormType = Form.TypeEx
        If Not IO.File.Exists(sPath & "\" + FormType + ".xml") Then
            oXmlDoc = New Xml.XmlDocument
            '// get the form as an XML string
            sXmlString = Form.GetAsXML
            '// load the form's XML string to the
            '// XML document object
            oXmlDoc.LoadXml(sXmlString)
            '// save the XML Document
            oXmlDoc.Save(sPath & "\" + FormType + ".xml")
        End If
    End Sub

    ''' <summary>
    ''' Tjek om en tabel er oprettet
    ''' </summary>
    ''' <param name="TableName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DoTableExist(ByVal TableName As String) As Boolean
        Dim UT As SAPbobsCOM.UserTablesMD
        UT = Nothing
        System.GC.Collect()
        UT = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        Return UT.GetByKey(TableName)
    End Function

    Public Sub OpretUserTable(ByVal TableName As String, ByVal Descr As String, Optional ByVal TableType As Integer = 0)
        ' Private Sub TjekUserTables(ByVal TableName As String, ByVal Descr As String)
        Dim UT As SAPbobsCOM.UserTablesMD
        UT = Nothing
        System.GC.Collect()
        UT = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        '//opret Tabel
        UT.TableName = TableName
        UT.TableDescription = Descr
        UT.TableType = TableType
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String = ""
        RetVal = UT.Add()
        If RetVal <> 0 Then
            vCmp.GetLastError(ErrCode, ErrMsg)
            SBO_Application.MessageBox(TL("Fejlede ved oprettelse af UserTable") + ": " & TableName & " " & ErrCode & " " & ErrMsg)
            ErrorLog("OpretUserTable ", ErrMsg)
        Else
            SBO_Application.StatusBar.SetText(TL("Tabel") + ": " + TableName + " " + TL("Oprettet"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
        ' End If
        UT = Nothing
        System.GC.Collect()

    End Sub

    Public Sub OpretUserDefinedField(ByVal Tablename As String, ByVal Name As String, ByVal Description As String, _
                                    ByVal Size As Integer, ByVal Type As String, Optional ByVal SubType As String = "", _
                                    Optional ByVal EditSize As Integer = 0, _
                                    Optional ByVal Mandatory As String = "", Optional ByVal DefaultValue As String = "", _
                                    Optional ByVal LinkedTable As String = "", _
                                    Optional ByVal ValidValues As Array = Nothing)

        Dim UDF As SAPbobsCOM.UserFieldsMD
        UDF = Nothing
        System.GC.Collect()
        UDF = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)


        '// Opret userfield
        UDF.TableName = Tablename
        UDF.Name = Name
        UDF.Description = Description
        UDF.Size = Size

        If EditSize > 0 Then
            UDF.EditSize = EditSize
        End If

        UDF.DefaultValue = DefaultValue
        UDF.LinkedTable = LinkedTable
        Dim i As Integer = 0

        If Not ValidValues Is Nothing Then
            test = ValidValues.Length
            Do Until i = (ValidValues.Length / 2)

                If ValidValues.GetValue(i, 0) = Nothing Then
                    Exit Do
                End If
                If i > 0 Then
                    UDF.ValidValues.Add()
                End If
                UDF.ValidValues.SetCurrentLine(i)
                UDF.ValidValues.Value = ValidValues.GetValue(i, 0)
                UDF.ValidValues.Description = ValidValues.GetValue(i, 1)

                i = i + 1
            Loop
        End If

        Select Case Mandatory
            Case "N"
                UDF.Mandatory = SAPbobsCOM.BoYesNoEnum.tNO
            Case "Y"
                UDF.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES
        End Select

        Select Case Type
            Case "A"
                UDF.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            Case "N"
                UDF.Type = SAPbobsCOM.BoFieldTypes.db_Numeric
            Case "M"
                UDF.Type = SAPbobsCOM.BoFieldTypes.db_Memo
            Case "D"
                UDF.Type = SAPbobsCOM.BoFieldTypes.db_Date
            Case "B"
                UDF.Type = SAPbobsCOM.BoFieldTypes.db_Float
        End Select

        Select Case SubType
            Case "?"
                UDF.SubType = SAPbobsCOM.BoFldSubTypes.st_Address
            Case "#"
                UDF.SubType = SAPbobsCOM.BoFldSubTypes.st_Phone
            Case "T"
                UDF.SubType = SAPbobsCOM.BoFldSubTypes.st_Time
            Case "R"
                UDF.SubType = SAPbobsCOM.BoFldSubTypes.st_Rate
            Case "S"
                UDF.SubType = SAPbobsCOM.BoFldSubTypes.st_Sum
            Case "P"
                UDF.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
            Case "Q"
                UDF.SubType = SAPbobsCOM.BoFldSubTypes.st_Quantity
            Case "%"
                UDF.SubType = SAPbobsCOM.BoFldSubTypes.st_Percentage
            Case "M"
                UDF.SubType = SAPbobsCOM.BoFldSubTypes.st_Measurement
            Case "B"
                UDF.SubType = SAPbobsCOM.BoFldSubTypes.st_Link
            Case "I"
                UDF.SubType = SAPbobsCOM.BoFldSubTypes.st_Image
        End Select

        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String = ""

        RetVal = UDF.Add()
        If RetVal <> 0 Then
            vCmp.GetLastError(ErrCode, ErrMsg)
            If ErrMsg.Contains("duplicate") = False Then
                UDF = Nothing
                System.GC.Collect()
                ErrorLog("OpretUserDefinedField ", ErrMsg)
                Throw New ApplicationException(ErrMsg)
            End If
        End If

        UDF = Nothing
        System.GC.Collect()

    End Sub

    Public Sub ErrorLog(ByVal SubMetode As String, ByVal ErrorMsg As String, Optional ByVal UserName As String = "")
        Dim Path As String
        Dim objwriter As System.IO.StreamWriter = Nothing
        Dim TimeStamp As String
        Dim MessageString As String
        Try
            UserName = vCmp.UserName
        Catch ex As Exception
            UserName = "NoName"
        End Try
        Try

            Path = (System.Windows.Forms.Application.StartupPath() + "\AIG_ERRORLOG-" + UserName + ".log")

            TimeStamp = Now + ":" + Right("00" + Now.Second.ToString, 2) + "," + Now.Millisecond.ToString

            MessageString = TimeStamp + " Sub: " + SubMetode + " Msg: " + ErrorMsg
            objwriter = New System.IO.StreamWriter(Path, True)
            objwriter.Write(MessageString)
            objwriter.WriteLine()
            objwriter.Flush()
        Catch ex As Exception

        Finally
            If Not objwriter Is Nothing Then
                objwriter.Close()
            End If
        End Try
    End Sub


    Public Sub DebugLog(ByVal Msg As String, Optional ByVal UserName As String = "")
        Dim path As String
        Dim objwriter As System.IO.StreamWriter = Nothing
        Dim TimeStamp As String
        Dim MessageString As String
        Try
            UserName = vCmp.UserName
        Catch ex As Exception
            UserName = "NoName"
        End Try

        Try

            'Dim file As IO.File
            path = (System.Windows.Forms.Application.StartupPath() + "\AIG_DEBUG_NO.log")
            If IO.File.Exists(path) Then
                Exit Sub
            End If

            path = (System.Windows.Forms.Application.StartupPath() + "\AIG_DEBUGLOG-" + UserName + ".log")

            TimeStamp = Now + ":" + Right("00" + Now.Second.ToString, 2) + "," + Now.Millisecond.ToString
            MessageString = TimeStamp + " Msg: " + Msg
            objwriter = New System.IO.StreamWriter(path, True)
            objwriter.Write(MessageString)
            objwriter.WriteLine()
            objwriter.Flush()

        Catch ex As Exception
        Finally
            If Not objwriter Is Nothing Then
                objwriter.Close()
            End If
        End Try
    End Sub


    Function GetMinFrom2000(ByVal D As Date, ByVal T As Long) As Long
        Dim Min As Long

        Min = DateDiff(DateInterval.Minute, D, Convert.ToDateTime("2000-01-01 00:00:00"), FirstDayOfWeek.Monday, FirstWeekOfYear.System)



        Return Min
    End Function


    Function GetWeekDayFromDate(ByVal D As Date) As String
        'Finder ugedag ud fra dato
        Dim WD As String = TL("Søndag")

        Select Case D.DayOfWeek
            Case DayOfWeek.Sunday
                WD = TL("Søndag")
            Case DayOfWeek.Monday
                WD = TL("Mandag")
            Case DayOfWeek.Tuesday
                WD = TL("Tirsdag")
            Case DayOfWeek.Wednesday
                WD = TL("Onsdag")
            Case DayOfWeek.Thursday
                WD = TL("Torsdag")
            Case DayOfWeek.Friday
                WD = TL("Fredag")
            Case DayOfWeek.Saturday
                WD = TL("Lørdag")
        End Select
        Return WD
    End Function

    Function GetLocalNextNum(ByVal LastNum As String, ByVal AddTo As Integer) As String
        '// Denne function kan bruges ved opbygning af SQL strenge til Batch overførsel
        Dim strNum As String
        Dim intNum As Integer
        Dim filler As String = "00000000"

        intNum = Convert.ToInt32(LastNum)
        intNum = intNum + AddTo
        strNum = intNum.ToString
        strNum = Mid(filler, 1, (8 - strNum.Length)) + strNum

        Return strNum

    End Function
    Public Function GetNextCode02(Optional ByVal IncrementBy As Int32 = 1) As Int32

        Dim nextCode As Int32 = 0
        Dim newMaxCode As Int32 = 0
        Dim sql As String = ""

        Try
            sql = "declare @RCSNumber table(next int, new int) "
            sql += vbCrLf + "update [@AIG_KLADDE_07] set U_N1 = U_N1 + " + IncrementBy.ToString + " "
            sql += vbCrLf + "output deleted.U_N1, inserted.U_N1 into @RCSNumber"
            sql += vbCrLf + "where Name = 'AIG_KLADDE_02'"
            sql += vbCrLf + "select * from @RCSNumber"
            If rsGetNextRecord Is Nothing Then
                rsGetNextRecord = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
            End If
            rsGetNextRecord.DoQuery(sql)
            nextCode = rsGetNextRecord.Fields.Item("next").Value
            newMaxCode = rsGetNextRecord.Fields.Item("new").Value
            If newMaxCode > 99999998 Then
                sql = "Worksheet record numbering has reached its limit!"
                sql += vbCrLf + "Please contact Your support provider."
                SBO_Application.MessageBox(sql)
            End If
        Catch ex As Exception
            ErrLog("GetNextRecord", ex)
            nextCode = -1
        End Try
        Return nextCode
    End Function
    'Function GetNextNum(ByVal Table As String, ByVal AddTo As Integer) As String

    '    Dim strNum, sql As String
    '    Dim intNum As Integer
    '    Dim filler As String = "00000000"
    '    '// Få sidste buntNr
    '    Dim rsGetNextNum As SAPbobsCOM.IRecordset
    '    rsGetNextNum = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    sql = "Select top 1 isnull(Code, '0') As Code from " + Table + " ORDER BY Code DESC"
    '    rsGetNextNum.DoQuery(sql)
    '    strNum = rsGetNextNum.Fields.Item("Code").Value
    '    If strNum = "" Then
    '        strNum = "0"
    '    End If
    '    intNum = Convert.ToInt32(strNum)
    '    intNum = intNum + AddTo
    '    strNum = intNum.ToString
    '    strNum = Mid(filler, 1, (8 - strNum.Length)) + strNum

    '    Return strNum

    'End Function

    Function GetNextNum(ByVal Table As String, ByVal AddTo As Integer) As String

        Dim strNum As String
        Dim sql As String
        Dim intNum As Integer
        Dim filler As String = "00000000"
        '// Få sidste buntNr
        'AIG_STANDARD.GetNextNum("[@AIG_KLADDE_02]", 1)
        If Table = "[@AIG_KLADDE_02]" Then
            intNum = GetNextCode02(AddTo)
        Else
            If rsGetNextNum Is Nothing Then
                rsGetNextNum = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
            End If

            'Dim rsGetNextNum As SAPbobsCOM.IRecordset
            'rsGetNextNum = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            sql = "Select top 1 isnull(Code, '0') As Code from " + Table + " ORDER BY Code DESC"
            rsGetNextNum.DoQuery(sql)
            If rsGetNextNum.RecordCount = 0 Then
                strNum = "0"
            Else
                strNum = rsGetNextNum.Fields.Item("Code").Value
                If strNum = "" Then
                    strNum = "0"
                End If
            End If

            intNum = Convert.ToInt32(strNum)
            intNum = intNum + AddTo
        End If

        strNum = intNum.ToString
        ' strNum = Mid(filler, 1, (8 - strNum.Length)) + strNum
        strNum = Right(filler + strNum, 8)
        Return strNum

    End Function


    Public Function FormOpen(ByVal formUID As String) As Boolean
        Dim i As Integer
        i = 0
        Do While i < SBO_Application.Forms.Count
            If formUID = SBO_Application.Forms.Item(i).UniqueID Then
                SBO_Application.Forms.Item(formUID).Select()
                FormOpen = True
                Exit Function
            End If
            i = i + 1
        Loop
        FormOpen = False
    End Function

    Public Function AIG_FindFormUID(ByVal formUID As String) As Boolean
        Dim i As Integer

        i = 0
        Do While i < SBO_Application.Forms.Count
            If formUID = SBO_Application.Forms.Item(i).UniqueID Then
                AIG_FindFormUID = True
                Exit Function
            End If
            i = i + 1
        Loop
        AIG_FindFormUID = False
    End Function

End Module
