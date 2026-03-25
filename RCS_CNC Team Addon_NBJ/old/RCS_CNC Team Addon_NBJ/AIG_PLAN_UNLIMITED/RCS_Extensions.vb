Module RCS_Extensions

#Region "SAP recordset extensions"

    <System.Runtime.CompilerServices.Extension()>
    Public Function GetDecimalValueAsDBString(ByRef rs As SAPbobsCOM.Recordset, ByVal FieldName As String) As String
        Dim result As String = ""
        Dim Value As Decimal = CType(rs.Fields.Item(FieldName).Value, Decimal)
        Dim culture As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-US")
        result = Value.ToString("G", culture)
        Return result
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function GetDoubleValueAsDBString(ByRef rs As SAPbobsCOM.Recordset, ByVal FieldName As String) As String
        Dim result As String = ""
        Dim Value As Double = CType(rs.Fields.Item(FieldName).Value, Double)
        Dim culture As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-US")
        result = Value.ToString("G", culture)
        Return result
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function GetDateValueAsDBString(ByRef rs As SAPbobsCOM.Recordset, ByVal FieldName As String) As String
        Dim result As String = ""
        Dim Value As Date = CType(rs.Fields.Item(FieldName).Value, Date)
        Dim culture As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-US")
        result = Value.ToString("YYYY-mm-DD")
        Return result
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function ValueAsString(ByRef rs As SAPbobsCOM.Recordset, ByVal FieldName As String) As String
        Dim result As String = CType(rs.Fields.Item(FieldName).Value, String)
        Return result
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function ValueAsDecimal(ByRef rs As SAPbobsCOM.Recordset, ByVal FieldName As String) As Decimal
        'Dim result As Decimal = CType(rs.Fields.Item(FieldName).Value, Decimal)
        Dim result As Decimal = 0
        If Decimal.TryParse(rs.Fields.Item(FieldName).Value, result) Then

        End If
        Return result
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function ValueAsDouble(ByRef rs As SAPbobsCOM.Recordset, ByVal FieldName As String) As Double
        Dim result As Double = 0
        'result As Double = CType(rs.Fields.Item(FieldName).Value, Double)
        If Double.TryParse(rs.Fields.Item(FieldName).Value, result) Then

        End If
        Return result
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function ValueAsInteger(ByRef rs As SAPbobsCOM.Recordset, ByVal FieldName As String) As Integer
        'Dim result As Integer = CType(rs.Fields.Item(FieldName).Value, Integer)
        Dim result As Integer = 0
        If Integer.TryParse(rs.Fields.Item(FieldName).Value, result) Then

        End If
        Return result
    End Function

    <System.Runtime.CompilerServices.Extension()>
    Public Function ValueAsDate(ByRef rs As SAPbobsCOM.Recordset, ByVal FieldName As String) As Date
        Dim result As Date = CType(rs.Fields.Item(FieldName).Value, Date)
        Return result
    End Function

#End Region

    ''' <summary>
    ''' Releases comobject and runs garbage collector and returns Nothing
    ''' </summary>
    ''' <param name="o"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function ReleaseComObject(ByRef o As Object) As Object
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
            o = Nothing
        Catch ex As Exception

        End Try
        Return o
    End Function

    ''' <summary>
    ''' Releases SAPbobsCOM.Recordset and runs garbage collector and returns Nothing
    ''' </summary>
    ''' <param name="o"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function ReleaseComObject(ByRef o As SAPbobsCOM.Recordset) As Object
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
            o = Nothing
        Catch ex As Exception

        End Try
        Return o
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToDBString(ByRef d As Date) As String
        Dim result As String = d.ToString("yyyy-MM-dd")
        Return result
    End Function

    ''' <summary>
    ''' Returns date as sting in format "yyyyMMdd hh:MM:ss.fff"
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToTimeStampString(ByRef d As Date) As String
        Dim result As String = d.ToString("yyyyMMdd hh:mm:ss.fff")
        Return result
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToDBString(ByRef d As Decimal) As String
        Dim culture As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-US")
        Dim result As String = d.ToString("G", culture)
        Return result
    End Function

    ''' <summary>
    ''' Returns string formated with dot as decimal point
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToDBString(ByRef d As Double) As String
        Dim culture As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-US")
        Dim result As String = d.ToString("G", culture)
        Return result
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToDBString(ByRef s As String) As String
        Dim result As String = s.Replace("'", "''")
        Return result
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToStringFormatted(ByRef Number As Double, ByVal GroupSeperator As String, _
                             ByVal DecimalSeperator As String, ByVal DecimalDigits As Integer) As String
        Dim Text As String
        If DecimalSeperator = "" Then
            DecimalSeperator = ","
        End If
        Dim NumberInfo As System.Globalization.NumberFormatInfo = New System.Globalization.NumberFormatInfo
        NumberInfo.NumberGroupSeparator = GroupSeperator
        NumberInfo.NumberDecimalSeparator = DecimalSeperator
        NumberInfo.NumberDecimalDigits = DecimalDigits
        Text = Number.ToString("N", NumberInfo)
        Return Text
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToStringFormatted(ByRef Number As Decimal, ByVal GroupSeperator As String, _
                             ByVal DecimalSeperator As String, ByVal DecimalDigits As Integer) As String
        Dim Text As String
        If DecimalSeperator = "" Then
            DecimalSeperator = ","
        End If
        Dim NumberInfo As System.Globalization.NumberFormatInfo = New System.Globalization.NumberFormatInfo
        NumberInfo.NumberGroupSeparator = GroupSeperator
        NumberInfo.NumberDecimalSeparator = DecimalSeperator
        NumberInfo.NumberDecimalDigits = DecimalDigits
        Text = Number.ToString("N", NumberInfo)
        Return Text
    End Function


    <System.Runtime.CompilerServices.Extension> _
    Public Function ToDateTime(s As String, Optional format As String = "yyyyMMdd", Optional cultureString As String = "en-US") As DateTime
        Try
            Dim r As Date = DateTime.ParseExact(s:=s, format:=format, provider:=Globalization.CultureInfo.GetCultureInfo(cultureString))
            Return r
        Catch ex As Exception

        End Try
    End Function




End Module
