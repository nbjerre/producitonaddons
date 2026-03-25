Module CodeGenerator

    ''' <summary>
    ''' Genererer Kode til at oprette DB Felter
    ''' </summary>
    ''' <remarks>Koden kommer til at ligge som en text fil med tabellens navn</remarks>
    Public Sub CodeGenerator()

        Dim TableName As String
        Dim TableArray(20) As String
        Dim Code As String
        Dim objwriter As System.IO.StreamWriter = Nothing

        '// Tilføj de tabeller der ønskes oprettet i array'et
        TableArray.SetValue("OHEM", 0)
        'TableArray.SetValue("AIG_OSYS", 1)
        'TableArray.SetValue("AIG_PPR", 2)
        'TableArray.SetValue("AIG_RPD", 3)
        'TableArray.SetValue("AIG_RUP", 4)
        'TableArray.SetValue("AIG_SDR", 5)
        'TableArray.SetValue("AIG_TOOLS_DOC", 6)
        'TableArray.SetValue("AIG_TOOLS_PK", 7)
        'TableArray.SetValue("AIG_TOOLS_SIG", 8)
        'TableArray.SetValue("AIG_TOOLS_TG", 9)

        rs = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim AliasID, Descr, SizeID, TypeID, EditType, EditSize, NotNull, Dflt, RTable, path, sql As String
        Dim i As Integer = 0
        Do Until i >= TableArray.Length
            TableName = TableArray.GetValue(i)
            If TableName <> "" Then
                Code = ""


                sql = "select * from CUFD where (TableID = '@" + TableName + "')"
                rs.DoQuery(sql)
                Do Until rs.EoF

                    AliasID = rs.Fields.Item("AliasID").Value
                    Descr = rs.Fields.Item("Descr").Value
                    SizeID = rs.Fields.Item("SizeID").Value.ToString
                    TypeID = rs.Fields.Item("TypeID").Value
                    EditType = rs.Fields.Item("EditType").Value
                    EditSize = rs.Fields.Item("EditSize").Value.ToString
                    NotNull = rs.Fields.Item("NotNull").Value
                    Dflt = rs.Fields.Item("Dflt").Value
                    RTable = rs.Fields.Item("RTable").Value


                    Code = Code + "Try" + vbCrLf
                    Code = Code + " AIG_STANDARD.OpretUserDefinedField(""" + TableName + ""","""
                    Code = Code + AliasID + ""","""
                    Code = Code + Descr + ""","
                    Code = Code + SizeID + ","""
                    Code = Code + TypeID + ""","""
                    Code = Code + EditType + ""","
                    Code = Code + EditSize + ","""
                    Code = Code + NotNull + ""","""
                    Code = Code + Dflt + ""","""
                    Code = Code + RTable + """)" + vbCrLf
                    Code = Code + " Catch ex As Exception" + vbCrLf
                    Code = Code + " Fejl = True" + vbCrLf
                    Code = Code + " End Try" + vbCrLf + vbCrLf


                    rs.MoveNext()
                Loop

                path = (System.Windows.Forms.Application.StartupPath() + "\" + TableName + ".txt")
                objwriter = New System.IO.StreamWriter(path, True)
                objwriter.Write(Code)
                objwriter.WriteLine()
                objwriter.Flush()

            End If

            i = i + 1
        Loop


    End Sub

End Module
