Module AIG_KGJT_PlanUnlimited

    Public Class Response
        Public Property Success As Boolean = True
        Public Property Message As String = String.Empty
        Public Property ArticleId As String = String.Empty
    End Class

    Private Class ProduktionOrderLine
        Public DocEntry As Integer
        Public ObjType As Integer
        Public LineNum As Integer
        Public PlannedQty As Double
        Public IssuedQty As Double
        Public RestQty As Double
        Public U_AIG_ARBT As Double
        Public StartDate As Date
        Public EndDate As Date
        Public VisResCode As String
        Public ResGrpCod As Integer
        Public WorkingDates As New System.Collections.Generic.List(Of Date)
        Public CapacityPerDay As Double
    End Class


    Public HuskProdDateForfald As Date
    Public HuskProdDateForfaldNy As Date
    Public HuskProdDateStart As Date
    Public HuskProdDateStartNy As Date
    Public HuskProdDateEnd As Date
    Public HuskProdDateEndNy As Date
    Public HuskProdDateStartForFaldNy As Date


    Private Sub WriteInConsole(ByVal text As String, Optional ByVal showMsg As Boolean = False)
        If Not String.IsNullOrEmpty(text) Then
            'showMsg = True 'Changeto True to show all Messages
            If My.Settings.ShowAllMessages Or showMsg Then
                Console.WriteLine(text)
            End If
        End If
    End Sub



#Region "Plan Order (Planlæg Linje i Kundeordre)"

    'Public Function PlanAllProductionOrdersPerLineOrig(ByVal planUnlimeted As Boolean, ByVal salesOrderDocEntry As Long, ByVal lineNum As Integer) As Response

    '    Dim response As Response = New Response

    '    Dim test As Object = ""
    '    Dim test2 As String

    '    Dim oSO As SAPbobsCOM.Documents
    '    oSO = vCmp.GetBusinessObject(BoObjectTypes.oOrders)
    '    Dim rs As SAPbobsCOM.Recordset
    '    Dim rs1 As SAPbobsCOM.Recordset
    '    Dim rs2 As SAPbobsCOM.Recordset
    '    rs = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    rs1 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    rs2 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    rs3 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    rs4 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '    Try
    '        'oForm = SBO_Application.Forms.Item(FormUID)
    '        'oForm = SBO_Application.Forms.ActiveForm
    '        'oForm.Update()

    '        'Dim salesOrderDocEntry As Long
    '        'Dim xDocParams As Xml.XmlDocument
    '        'xDocParams = New Xml.XmlDocument
    '        '    If String.IsNullOrEmpty(SalesOrderItemcode) Then
    '        '    'SBO_Application.StatusBar.SetText("Vælg en linje i ordren til produktionsordre", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
    '        '   Return
    '        '    End If
    '        'Dim oMatrix As SAPbouiCOM.Matrix
    '        'oMatrix = oForm.Items.Item("38").Specific

    '        'Dim lineNum As Integer = -1
    '        'Dim i As Integer = 0
    '        'Dim oEdit As SAPbouiCOM.EditText
    '        ' Dim oCell As SAPbouiCOM.Cell

    '        'Do While i < oMatrix.RowCount
    '        '    i = i + 1
    '        '    If oMatrix.IsRowSelected(i) = True Then
    '        '        'oCell = oMatrix.Columns.Item("257").Cells.Item(i).Specific
    '        '        'test = oCell.GetType
    '        '        oEdit = oMatrix.Columns.Item("110").Cells.Item(i).Specific
    '        '        If Not String.IsNullOrEmpty(oEdit.Value) Then
    '        '            Integer.TryParse(oEdit.Value, lineNum)
    '        '        End If
    '        '        Exit Do
    '        '    End If
    '        'Loop

    '        'If lineNum < 0 Then
    '        '    SBO_Application.StatusBar.SetText("Ingen linje valgt", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
    '        '    Return
    '        'End If

    '        'xDocParams.LoadXml(oForm.BusinessObject.Key)

    '        Dim oProd As SAPbobsCOM.ProductionOrders
    '        oProd = vCmp.GetBusinessObject(BoObjectTypes.oProductionOrders)

    '        Dim iResult As Integer = 0
    '        Dim Msg As String = ""

    '        'Dim xmlNode As System.Xml.XmlNode = xDocParams.SelectSingleNode("DocumentParams/DocEntry")

    '        Dim HuskFatherStartDate As Date
    '        Dim MoveDate As Integer = 0
    '        'oForm.


    '        Dim HuskChieldEndDate As Date
    '        'Dim sb As New Text.StringBuilder()
    '        'sb.Append("select DocEntry from ordr  ")
    '        'sb.Append("Where DocNum = ").Append(salesOrderDocEntry))
    '        'Sql = sb.ToString
    '        'rs.DoQuery(Sql)

    '        'If Integer.TryParse(xmlNode.InnerText, salesOrderDocEntry) Then
    '        If (oSO.GetByKey(salesOrderDocEntry)) Then
    '            Dim sql As String
    '            Dim sb As New Text.StringBuilder()
    '            Dim Cnt As Integer = 0
    '            Dim PlanOk As Boolean = False

    '            sb.Append("select distinct project, U_RCS_AFS, VisOrder from rdr1  ").AppendLine()
    '            sb.Append("Where Docentry = ").Append(salesOrderDocEntry).Append(" and LineNum = ").Append(lineNum).AppendLine()
    '            sb.Append("And RDR1.U_RCS_ONSTO = 'Y' order by Visorder")
    '            sql = sb.ToString
    '            rs3.DoQuery(sql)
    '            Dim ProdDocEntry As Integer
    '            Dim ProdDocnum As String = ""
    '            Dim ProdItemCode As String = ""
    '            Dim HuskNyStartDate As Date = "01.01.9999"
    '            Dim HuskNyEndDate As Date = "01.01.9999"
    '            Dim NextIsParallel As Boolean = False
    '            Dim HuskFather As String = ""
    '            Dim HuskLowestNyStartDate As Date = "01.01.9999"


    '            Do Until rs3.EoF
    '                HuskLowestNyStartDate = "01.01.9999"
    '                Cnt = 0
    '                Dim testdate = rs3.Fields.Item("U_RCS_AFS").Value


    '                sb.Clear()
    '                sb.Append("select ItemCode, DocEntry, Docnum from OWOR  ").AppendLine()
    '                sb.Append("Where Project='").Append(rs3.Fields.Item("Project").Value).Append("'").AppendLine()
    '                sb.Append("And U_RCS_BVO = ").Append(rs3.Fields.Item("Visorder").Value).AppendLine()
    '                sb.Append("and status <> 'C'").AppendLine()
    '                sb.Append("and OWOR.OriginAbs = ").Append(salesOrderDocEntry).AppendLine()
    '                sb.Append("order by DocEntry ")
    '                sql = sb.ToString()
    '                rs1.DoQuery(sql)

    '                Do Until rs1.EoF

    '                    ProdDocnum = rs1.ValueAsString("DocNum")
    '                    ProdDocEntry = rs1.ValueAsInteger("DocEntry")
    '                    ProdItemCode = rs1.ValueAsString("ItemCode")

    '                    sql = "Delete from ""@RCS_CAP_COMMIT"" where U_DocEntry = " + ProdDocEntry.ToString + ""
    '                    rs2.DoQuery(sql)
    '                    rs1.MoveNext()
    '                Loop
    '                rs3.MoveNext()
    '            Loop
    '            '//*********************
    '            rs3.MoveFirst()
    '            rs1.MoveFirst()

    '            Do Until rs3.EoF
    '                Dim testdate = rs3.Fields.Item("U_RCS_AFS").Value
    '                Dim AfsDate As Date = rs3.Fields.Item("U_RCS_AFS").Value
    '                PlanOk = False
    '                Do Until PlanOk = True
    '                    PlanOk = True
    '                    HuskLowestNyStartDate = "01.01.9999"
    '                    Cnt = 0

    '                    sb.Clear()
    '                    sb.Append("select ItemCode, DocEntry, Docnum from OWOR  ").AppendLine()
    '                    sb.Append("Where Project='").Append(rs3.Fields.Item("Project").Value).Append("'").AppendLine()
    '                    sb.Append("And U_RCS_BVO = ").Append(rs3.Fields.Item("Visorder").Value).AppendLine()
    '                    sb.Append("and status <> 'C'").AppendLine()
    '                    sb.Append("and OWOR.OriginAbs = ").Append(salesOrderDocEntry).AppendLine()
    '                    sb.Append("order by DocEntry ")
    '                    sql = sb.ToString()
    '                    rs1.DoQuery(sql)

    '                    Do Until rs1.EoF
    '                        ProdDocnum = rs1.ValueAsString("DocNum")
    '                        ProdDocEntry = rs1.ValueAsInteger("DocEntry")
    '                        ProdItemCode = rs1.ValueAsString("ItemCode")
    '                        sql = "Delete from ""@RCS_CAP_COMMIT"" where U_DocEntry = " + ProdDocEntry.ToString + ""
    '                        rs2.DoQuery(sql)

    '                        If Cnt = 0 Then
    '                            oProd = vCmp.GetBusinessObject(BoObjectTypes.oProductionOrders)

    '                            oProd.GetByKey(ProdDocEntry)
    '                            oProd.StartDate = Date.Now

    '                            iResult = oProd.Update()

    '                            oProd.GetByKey(ProdDocEntry)

    '                            oProd.DueDate = AfsDate
    '                            'oProd.RoutingDateCalculation = ResourceAllocationEnum.raEndDateBackwards
    '                            Dim cnt2 As Integer = 0

    '                            Do Until cnt2 > oProd.Stages.Count - 1
    '                                oProd.Stages.SetCurrentLine(cnt2)
    '                                oProd.Stages.StartDate = Date.Now
    '                                oProd.Stages.EndDate = AfsDate
    '                                cnt2 = cnt2 + 1
    '                            Loop

    '                            iResult = oProd.Update()

    '                            oProd.GetByKey(ProdDocEntry)
    '                            If oProd.DueDate <> AfsDate Then
    '                                Dim testttt As String
    '                                testttt = "test"
    '                            End If

    '                            'oProd.SaveXML(ProdDocEntry.ToString + ".xml")
    '                            oProd.ReleaseComObject
    '                            If iResult < 0 Then
    '                                'vCmp.GetLastError(iResult, Msg)
    '                                Msg = "Produkttionsordre " + ProdDocnum + " error:  " + Msg
    '                                ErrorLog("PlanAllProductionOrders", Msg)
    '                                'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
    '                                response.Message += Msg
    '                                WriteInConsole(Msg)
    '                            Else

    '                                'Msg = "Produkttionsordre " + ProdDocnum + " er opdateret med forfaldsdato"
    '                                'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
    '                                'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
    '                            End If
    '                        Else
    '                            If InStr(HuskFather, ProdItemCode) > 0 Then

    '                                oProd = vCmp.GetBusinessObject(BoObjectTypes.oProductionOrders)
    '                                oProd.GetByKey(ProdDocEntry)
    '                                oProd.StartDate = Date.Now

    '                                iResult = oProd.Update()

    '                                oProd.GetByKey(ProdDocEntry)
    '                                oProd.DueDate = HuskNyEndDate
    '                                Dim cnt2 As Integer = 0

    '                                Do Until cnt2 > oProd.Stages.Count - 1
    '                                    oProd.Stages.SetCurrentLine(cnt2)
    '                                    oProd.Stages.StartDate = Date.Now
    '                                    oProd.Stages.EndDate = HuskNyEndDate

    '                                    cnt2 = cnt2 + 1
    '                                Loop

    '                                iResult = oProd.Update()
    '                                '   oProd.SaveXML(ProdDocEntry.ToString + ".xml")
    '                                oProd.ReleaseComObject
    '                                If iResult < 0 Then
    '                                    vCmp.GetLastError(iResult, Msg)
    '                                    Msg = "Produkttionsordre " + ProdDocnum + " error: " + Msg
    '                                    ErrorLog("PlanAllProductionOrdersPERlINE", Msg)
    '                                    'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
    '                                    response.Message += Msg
    '                                    WriteInConsole(Msg)
    '                                Else
    '                                    'Msg = "Produkttionsordre er planlagt"
    '                                    'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
    '                                    '  SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
    '                                End If
    '                            Else

    '                                If HuskLowestNyStartDate < HuskNyStartDate Then
    '                                    oProd = vCmp.GetBusinessObject(BoObjectTypes.oProductionOrders)
    '                                    oProd.GetByKey(ProdDocEntry)
    '                                    oProd.StartDate = Date.Now

    '                                    iResult = oProd.Update()

    '                                    oProd.GetByKey(ProdDocEntry)
    '                                    oProd.DueDate = HuskLowestNyStartDate
    '                                    Dim cnt2 As Integer = 0

    '                                    Do Until cnt2 > oProd.Stages.Count - 1
    '                                        oProd.Stages.SetCurrentLine(cnt2)
    '                                        oProd.Stages.StartDate = Date.Now
    '                                        oProd.Stages.EndDate = HuskLowestNyStartDate

    '                                        cnt2 = cnt2 + 1
    '                                    Loop

    '                                    iResult = oProd.Update()
    '                                    ' oProd.SaveXML(ProdDocEntry.ToString + ".xml")
    '                                    oProd.ReleaseComObject
    '                                    If iResult < 0 Then
    '                                        vCmp.GetLastError(iResult, Msg)
    '                                        Msg = "Produkttionsordre " + ProdDocnum + " error: " + Msg
    '                                        ErrorLog("PlanAllProductionOrderspERlINE", Msg)
    '                                        'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
    '                                        response.Message += Msg
    '                                        WriteInConsole(Msg)
    '                                    Else
    '                                        '     Msg = "Produkttionsordre er planlagt"
    '                                        'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
    '                                        '    SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
    '                                    End If

    '                                Else
    '                                    oProd = vCmp.GetBusinessObject(BoObjectTypes.oProductionOrders)
    '                                    oProd.GetByKey(ProdDocEntry)
    '                                    oProd.StartDate = Date.Now

    '                                    iResult = oProd.Update()

    '                                    oProd.GetByKey(ProdDocEntry)
    '                                    oProd.DueDate = HuskNyStartDate
    '                                    Dim cnt2 As Integer = 0

    '                                    Do Until cnt2 > oProd.Stages.Count - 1
    '                                        oProd.Stages.SetCurrentLine(cnt2)
    '                                        oProd.Stages.StartDate = Date.Now
    '                                        oProd.Stages.EndDate = HuskNyStartDate

    '                                        cnt2 = cnt2 + 1
    '                                    Loop

    '                                    iResult = oProd.Update()
    '                                    '  oProd.SaveXML(ProdDocEntry.ToString + ".xml")
    '                                    oProd.ReleaseComObject
    '                                    If iResult < 0 Then
    '                                        vCmp.GetLastError(iResult, Msg)
    '                                        Msg = "Produkttionsordre " + ProdDocnum + " error: " + Msg
    '                                        ErrorLog("PlanAllProductionOrders", Msg)
    '                                        'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
    '                                        response.Message += Msg
    '                                        WriteInConsole(Msg)
    '                                    Else
    '                                        '   Msg = "Produkttionsordre er planlagt"
    '                                        'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
    '                                        '  SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
    '                                    End If
    '                                End If


    '                            End If

    '                        End If

    '                        Cnt = Cnt + 1

    '                        UpdateProductionOrderLineStartEndDate(ProdDocEntry, ProdDocnum, planUnlimeted, salesOrderDocEntry, rs3.Fields.Item("Visorder").Value)

    '                        sql = "select top 1 owor.Itemcode as Father, wor1.StartDate from WOR1 inner join OWOR on OWOR.DocEntry = WOR1.DocEntry "
    '                        sql = sql + "Where WOR1.DocEntry ='" + ProdDocEntry.ToString() + "' order by Wor1.linenum  "
    '                        rs2.DoQuery(sql)

    '                        Dim HuskNyStartDateFromPlan = rs2.Fields.Item("StartDate").Value

    '                        If HuskNyStartDateFromPlan > HuskNyStartDate Then
    '                            '   PlanOk = False
    '                            '  Exit Do

    '                        End If

    '                        HuskNyStartDate = HuskNyStartDateFromPlan



    '                        If HuskLowestNyStartDate > HuskNyStartDate Then
    '                            HuskLowestNyStartDate = HuskNyStartDate
    '                        End If
    '                        'check if have fahter
    '                        HuskFather = ""
    '                        '  sql = "Select Code from itt1 where Type = 4 And father = (Select top 1 Code from oitt where code = (Select top 1 Father from itt1 where code = '" + rs2.Fields.Item("Father").Value + "'))"

    '                        sql = "select wor1.Itemcode from wor1 where itemType = 4 And Wor1.docentry = (Select top 1 wor1.DocEntry from wor1 inner join owor On owor.docentry = wor1.docentry where itemType = 4 And owor.originAbs = " + salesOrderDocEntry.ToString() + "  And owor.U_RCS_BVO = " + rs3.Fields.Item("Visorder").Value.ToString() + " And wor1.itemcode = '" + rs2.Fields.Item("Father").Value.ToString() + "')"

    '                        rs2.DoQuery(sql)
    '                        If rs2.RecordCount > 0 Then


    '                            NextIsParallel = True
    '                            Do Until rs2.EoF
    '                                HuskFather = HuskFather + "§" + rs2.Fields.Item("ItemCode").Value + ""
    '                                rs2.MoveNext()
    '                            Loop


    '                            sql = "select DueDate from OWOR "
    '                            sql = sql + "Where DocEntry ='" + ProdDocEntry.ToString() + "'"
    '                            rs2.DoQuery(sql)
    '                            HuskNyEndDate = rs2.Fields.Item("DueDate").Value

    '                            If HuskNyEndDate > HuskFatherStartDate Then
    '                                PlanOk = False

    '                                MoveDate = DateDiff(DateInterval.Day, HuskFatherStartDate, HuskNyEndDate)
    '                                AfsDate = AfsDate.AddDays(MoveDate)

    '                                sql = "Update RDR1 set U_RCS_AFS = '" + AfsDate.ToString("yyyyMMdd") + "'"
    '                                sql = sql + " Where VisOrder = " + rs3.Fields.Item("Visorder").Value.ToString + " and DocEntry = " + salesOrderDocEntry.ToString
    '                                rs4.DoQuery(sql)

    '                                'SBO_Application.SetStatusBarMessage("Prøver med afsendelsesdato: " + AfsDate.ToString("dd-MM-yyyy"))
    '                                response.Message = "Prøver med afsendelsesdato: " + AfsDate.ToString("dd-MM-yyyy")
    '                                WriteInConsole("Prøver med afsendelsesdato: " + AfsDate.ToString("dd-MM-yyyy"))
    '                                Exit Do
    '                            End If
    '                        Else
    '                            HuskNyEndDate = HuskLowestNyStartDate
    '                            HuskFatherStartDate = HuskNyEndDate

    '                            ' sql = "select top 1 owor.Itemcode as Father, owor.dueDate from WOR1 inner join OWOR on OWOR.DocEntry = WOR1.DocEntry "
    '                            'sql = sql + "Where WOR1.DocEntry ='" + ProdDocEntry.ToString() + "' order by  Wor1.linenum desc "
    '                            'rs2.DoQuery(sql)

    '                            'HuskChieldEndDate = rs2.Fields.Item("DueDate").Value

    '                            'HuskFatherStartDate = HuskChieldEndDate
    '                        End If

    '                        rs1.MoveNext()

    '                    Loop



    '                Loop
    '                rs3.MoveNext()
    '            Loop
    '            '//*********************

    '            If iResult < 0 Then
    '                vCmp.GetLastError(iResult, Msg)
    '                ' Msg = "Produkttionsordre " + ProdDocnum + " error: " + Msg
    '                ErrorLog("PlanAllProductionOrderspERlINE", Msg)
    '                'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
    '                WriteInConsole(Msg)
    '            Else
    '                'Msg = "Produkttionsordrer er planlagt tjek meddelser!"
    '                'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
    '                'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
    '                WriteInConsole("Produkttionsordrer er planlagt", True)
    '            End If

    '            Return response
    '        End If

    '        'End If

    '        response.Message += $"Salgsordre {salesOrderDocEntry} er ikke fundet"
    '        response.Success = False
    '        Return response

    '    Catch ex As Exception
    '        ErrLog("PlanAllProductionOrdersPerLine", ex)

    '        response.Success = False
    '        response.Message += $"{vbCrLf}PlanAllProductionOrdersPerLine, message: {ex.Message}  exception: {ex.ToString}"
    '        Return response
    '    Finally
    '        oSO.ReleaseComObject
    '        rs.ReleaseComObject
    '        rs1.ReleaseComObject
    '        rs3.ReleaseComObject
    '    End Try

    'End Function

    Public Function PlanAllProductionOrdersPerLine(ByVal planUnlimeted As Boolean, ByVal salesOrderDocEntry As Long, ByVal lineNum As Integer) As Response

        Dim response As Response = New Response

        Dim test As Object = ""
        Dim test2 As String

        Dim oSO As Documents = vCmp.GetBusinessObject(BoObjectTypes.oOrders)
        Dim oProd As ProductionOrders = vCmp.GetBusinessObject(BoObjectTypes.oProductionOrders)
        Dim rs As SAPbobsCOM.Recordset
        Dim rs1 As SAPbobsCOM.Recordset
        Dim rs2 As SAPbobsCOM.Recordset
        rs = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs1 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs2 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs3 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs4 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            Dim iResult As Integer = 0
            Dim Msg As String = ""
            Dim HuskFatherStartDate As Date
            Dim MoveDate As Integer = 0

            If oSO.GetByKey(salesOrderDocEntry) Then
                Dim sql As String
                Dim sb As New Text.StringBuilder()
                Dim Cnt As Integer = 0
                Dim PlanOk As Boolean = False

                sb.Append("select distinct project, U_RCS_AFS, VisOrder from rdr1  ").AppendLine()
                sb.Append("Where Docentry = ").Append(salesOrderDocEntry).Append(" and LineNum = ").Append(lineNum).AppendLine()
                sb.Append("And RDR1.U_RCS_ONSTO = 'Y' order by Visorder")
                sql = sb.ToString
                rs3.DoQuery(sql)
                Dim ProdDocEntry As Integer
                Dim ProdDocnum As String = ""
                Dim ProdItemCode As String = ""
                Dim HuskNyStartDate As Date = "01.01.9999"
                Dim HuskNyEndDate As Date = "01.01.9999"
                Dim NextIsParallel As Boolean = False
                Dim HuskFather As String = ""
                Dim HuskLowestNyStartDate As Date = "01.01.9999"
                Dim DueDate As Date = "01.01.9999"

                Do Until rs3.EoF
                    Dim testdate = rs3.Fields.Item("U_RCS_AFS").Value
                    Dim AfsDate As Date = rs3.Fields.Item("U_RCS_AFS").Value
                    PlanOk = False
                    Do Until PlanOk = True
                        PlanOk = True
                        HuskLowestNyStartDate = "01.01.9999"
                        Cnt = 0

                        sb.Clear()
                        sb.Append("select ItemCode, DocEntry, Docnum from OWOR  ").AppendLine()
                        sb.Append("Where Project='").Append(rs3.Fields.Item("Project").Value).Append("'").AppendLine()
                        sb.Append("And U_RCS_BVO = ").Append(rs3.Fields.Item("Visorder").Value).AppendLine()
                        sb.Append("and status <> 'C'").AppendLine()
                        sb.Append("and OWOR.OriginAbs = ").Append(salesOrderDocEntry).AppendLine()
                        sb.Append("order by DocEntry ")
                        sql = sb.ToString()
                        rs1.DoQuery(sql)

                        Do Until rs1.EoF
                            ProdDocnum = rs1.ValueAsString("DocNum")
                            ProdDocEntry = rs1.ValueAsInteger("DocEntry")
                            ProdItemCode = rs1.ValueAsString("ItemCode")
                            sql = "Delete from ""@RCS_CAP_COMMIT"" where U_DocEntry = " + ProdDocEntry.ToString + ""
                            rs2.DoQuery(sql)

                            If Cnt = 0 Then
                                DueDate = AfsDate
                            Else
                                If InStr(HuskFather, ProdItemCode) > 0 Then
                                    DueDate = HuskNyEndDate
                                Else
                                    If HuskLowestNyStartDate < HuskNyStartDate Then
                                        DueDate = HuskLowestNyStartDate
                                    Else
                                        DueDate = HuskNyStartDate
                                    End If
                                End If
                            End If


                            'Console.WriteLine("ProdDocEntry:" + ProdDocEntry.ToString)
                            'Console.WriteLine("planUnlimeted:" + planUnlimeted.ToString)
                            'Console.WriteLine("salesOrderDocEntry:" + salesOrderDocEntry.ToString)
                            oProd.GetByKey(ProdDocEntry)
                            oProd.StartDate = Date.Now
                            oProd.DueDate = DueDate

                            Dim cnt2 As Integer = 0
                            'Console.WriteLine("oProd.Stages.Count:" + oProd.Stages.Count.ToString)
                            Do Until cnt2 > oProd.Stages.Count - 1
                                oProd.Stages.SetCurrentLine(cnt2)
                                oProd.Stages.StartDate = Date.Now
                                oProd.Stages.EndDate = DueDate
                                cnt2 = cnt2 + 1
                            Loop

                            iResult = oProd.Update()
                            If iResult < 0 Then
                                Msg = "Produkttionsordre med DocNum: " + ProdDocnum + " fejlet:  " + Msg
                                response.Message += $"{vbCrLf}Produkttionsordre {ProdDocnum} error: {Msg}"
                                ErrorLog("PlanAllProductionOrders", Msg)
                                'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                WriteInConsole(Msg)
                            Else
                                Msg = "Produkttionsordre med DocNum: " + ProdDocnum + " er planlagt"
                                'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
                                'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                                WriteInConsole(Msg)
                            End If



                            Cnt = Cnt + 1

                            'UpdateProductionOrderLineStartEndDate(ProdDocEntry, ProdDocnum, planUnlimeted, salesOrderDocEntry, rs3.Fields.Item("Visorder").Value)
                            UpdateProductionOrderLineStartEndDateNew(ProdDocEntry, ProdDocnum, planUnlimeted, salesOrderDocEntry, rs3.Fields.Item("Visorder").Value)

                            sql = "select top 1 owor.Itemcode as Father, wor1.StartDate from WOR1 inner join OWOR on OWOR.DocEntry = WOR1.DocEntry "
                            sql = sql + "Where WOR1.DocEntry ='" + ProdDocEntry.ToString() + "' order by Wor1.linenum  "
                            rs2.DoQuery(sql)

                            Dim HuskNyStartDateFromPlan = rs2.Fields.Item("StartDate").Value

                            'If HuskNyStartDateFromPlan > HuskNyStartDate Then
                            '    '   PlanOk = False
                            '    '  Exit Do

                            'End If

                            HuskNyStartDate = HuskNyStartDateFromPlan



                            If HuskLowestNyStartDate > HuskNyStartDate Then
                                HuskLowestNyStartDate = HuskNyStartDate
                            End If
                            'check if have fahter
                            HuskFather = ""

                            sql = "select wor1.Itemcode from wor1 where itemType = 4 And Wor1.docentry = (Select top 1 wor1.DocEntry from wor1 inner join owor On owor.docentry = wor1.docentry where itemType = 4 And owor.originAbs = " + salesOrderDocEntry.ToString() + "  And owor.U_RCS_BVO = " + rs3.Fields.Item("Visorder").Value.ToString() + " And wor1.itemcode = '" + rs2.Fields.Item("Father").Value.ToString() + "')"

                            rs2.DoQuery(sql)
                            If rs2.RecordCount > 0 Then


                                NextIsParallel = True
                                Do Until rs2.EoF
                                    HuskFather = HuskFather + "§" + rs2.Fields.Item("ItemCode").Value + ""
                                    rs2.MoveNext()
                                Loop


                                sql = "select DueDate from OWOR "
                                sql = sql + "Where DocEntry ='" + ProdDocEntry.ToString() + "'"
                                rs2.DoQuery(sql)
                                HuskNyEndDate = rs2.Fields.Item("DueDate").Value

                                If HuskNyEndDate > HuskFatherStartDate Then
                                    PlanOk = False

                                    MoveDate = DateDiff(DateInterval.Day, HuskFatherStartDate, HuskNyEndDate)
                                    AfsDate = AfsDate.AddDays(MoveDate)

                                    sql = "Update RDR1 set U_RCS_AFS = '" + AfsDate.ToString("yyyyMMdd") + "'"
                                    sql = sql + " Where VisOrder = " + rs3.Fields.Item("Visorder").Value.ToString + " and DocEntry = " + salesOrderDocEntry.ToString
                                    rs4.DoQuery(sql)

                                    'SBO_Application.SetStatusBarMessage("Prøver med afsendelsesdato: " + AfsDate.ToString("dd-MM-yyyy"))
                                    response.Message = "Prøver med afsendelsesdato: " + AfsDate.ToString("dd-MM-yyyy")
                                    WriteInConsole("Prøver med afsendelsesdato: " + AfsDate.ToString("dd-MM-yyyy"))
                                    Exit Do
                                End If
                            Else
                                HuskNyEndDate = HuskLowestNyStartDate
                                HuskFatherStartDate = HuskNyEndDate
                            End If

                            rs1.MoveNext()

                        Loop



                    Loop
                    rs3.MoveNext()
                Loop
                '//*********************

                If iResult < 0 Then
                    vCmp.GetLastError(iResult, Msg)
                    ' Msg = "Produkttionsordre " + ProdDocnum + " error: " + Msg
                    ErrorLog("PlanAllProductionOrderspERlINE", Msg)
                    'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    WriteInConsole(Msg)
                Else
                    'Msg = "Produkttionsordrer er planlagt tjek meddelser!"
                    'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
                    'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    WriteInConsole("Produkttionsordrer er planlagt", True)
                End If

                Return response
            End If

            response.Message += $"Salgsordre {salesOrderDocEntry} er ikke fundet"
            response.Success = False
            Return response

        Catch ex As Exception
            ErrLog("PlanAllProductionOrdersPerLine", ex)

            response.Success = False
            response.Message += $"{vbCrLf}PlanAllProductionOrdersPerLine, message: {ex.Message}  exception: {ex.ToString}"
            Return response
        Finally
            oSO.ReleaseComObject
            oProd.ReleaseComObject
            rs.ReleaseComObject
            rs1.ReleaseComObject
            rs3.ReleaseComObject
        End Try

    End Function

#End Region


#Region "Plan Order (Planlæg Ubegransæt og Planlæg Alle i Kundeordre)"
    Public Function PlanAllProductionOrders(ByVal planUnlimeted As Boolean, ByVal salesOrderDocEntry As Long) As Response

        Dim response As Response = New Response

        Dim test As Object = ""
        Dim oSO As Documents = vCmp.GetBusinessObject(BoObjectTypes.oOrders)
        Dim oProd As ProductionOrders = vCmp.GetBusinessObject(BoObjectTypes.oProductionOrders)
        Dim rs As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim rs1 As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim rs2 As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim rs3 As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim rs4 As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)

        Try
            Dim iResult As Integer = 0
            Dim Msg As String = ""

            If oSO.GetByKey(salesOrderDocEntry) Then
                Dim sql As String
                Dim sb As New Text.StringBuilder()
                Dim Cnt As Integer = 0
                Dim PlanOk As Boolean = False
                Dim HuskFatherStartDate As Date
                Dim MoveDate As Integer = 0
                Dim ProdDocEntry As Integer
                Dim ProdDocnum As String = ""
                Dim ProdItemCode As String = ""
                Dim HuskNyStartDate As Date = "01.01.9999"
                Dim HuskNyEndDate As Date = "01.01.9999"
                Dim NextIsParallel As Boolean = False
                Dim HuskFather As String = ""
                Dim HuskLowestNyStartDate As Date = "01.01.9999"
                Dim DueDate As Date = "01.01.9999"

                sb.Append("select distinct project, U_RCS_AFS, VisOrder, ShipDate from rdr1  ")
                sb.Append("Where Docentry = ").Append(salesOrderDocEntry).Append(" And RDR1.U_RCS_ONSTO = 'Y' order by ShipDate, Visorder")
                sql = sb.ToString
                rs3.DoQuery(sql)

                Do Until rs3.EoF
                    'Dim testdate = rs3.Fields.Item("U_RCS_AFS").Value
                    Dim AfsDate As Date = rs3.Fields.Item("U_RCS_AFS").Value
                    PlanOk = False
                    Do Until PlanOk = True
                        PlanOk = True
                        HuskLowestNyStartDate = "01.01.9999"
                        Cnt = 0

                        sb.Clear()
                        sb.Append("select ItemCode, DocEntry, Docnum from OWOR  ").AppendLine()
                        sb.Append("Where Project='").Append(rs3.Fields.Item("Project").Value).Append("'").AppendLine()
                        sb.Append("And U_RCS_BVO = ").Append(rs3.Fields.Item("Visorder").Value).AppendLine()
                        sb.Append("and status <> 'C'").AppendLine()
                        sb.Append("and OWOR.OriginAbs = ").Append(salesOrderDocEntry).AppendLine()
                        sb.Append("order by DocEntry ")
                        sql = sb.ToString()
                        rs1.DoQuery(sql)

                        Do Until rs1.EoF
                            ProdDocnum = rs1.ValueAsString("DocNum")
                            ProdDocEntry = rs1.ValueAsInteger("DocEntry")
                            ProdItemCode = rs1.ValueAsString("ItemCode")
                            sql = "Delete from ""@RCS_CAP_COMMIT"" where U_DocEntry = " + ProdDocEntry.ToString + ""
                            rs2.DoQuery(sql)

                            If Cnt = 0 Then
                                DueDate = AfsDate
                            Else
                                If InStr(HuskFather, ProdItemCode) > 0 Then
                                    DueDate = HuskNyEndDate
                                Else
                                    If HuskLowestNyStartDate < HuskNyStartDate Then
                                        DueDate = HuskLowestNyStartDate
                                    Else
                                        DueDate = HuskNyStartDate
                                    End If
                                End If
                            End If


                            'Console.WriteLine("ProdDocEntry:" + ProdDocEntry.ToString)
                            'Console.WriteLine("planUnlimeted:" + planUnlimeted.ToString)
                            'Console.WriteLine("salesOrderDocEntry:" + salesOrderDocEntry.ToString)
                            oProd.GetByKey(ProdDocEntry)
                            oProd.StartDate = Date.Now
                            oProd.DueDate = DueDate

                            Dim cnt2 As Integer = 0
                            'Console.WriteLine("oProd.Stages.Count:" + oProd.Stages.Count.ToString)
                            Do Until cnt2 > oProd.Stages.Count - 1
                                oProd.Stages.SetCurrentLine(cnt2)
                                oProd.Stages.StartDate = Date.Now
                                oProd.Stages.EndDate = DueDate
                                cnt2 = cnt2 + 1
                            Loop

                            iResult = oProd.Update()
                            If iResult < 0 Then
                                Msg = "Produkttionsordre med DocNum: " + ProdDocnum + " fejlet:  " + Msg
                                response.Message += $"{vbCrLf}Produkttionsordre med DocNum: {ProdDocnum} fejlet: {Msg}"
                                ErrorLog("PlanAllProductionOrders", Msg)
                                'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                WriteInConsole(Msg)
                            Else
                                Msg = "Produkttionsordre med DocNum: " + ProdDocnum + " er opdateret med forfaldsdato"
                                'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
                                'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                                WriteInConsole(Msg)
                            End If

                            Cnt = Cnt + 1

                            'Dim planUnlimeted As Boolean = False

                            'If pValItemUID = "BtnPU" Then
                            '    planUnlimeted = True
                            'End If

                            UpdateProductionOrderLineStartEndDateNew(ProdDocEntry, ProdDocnum, planUnlimeted, salesOrderDocEntry, rs3.Fields.Item("Visorder").Value)

                            sql = "select top 1 owor.Itemcode as Father, wor1.StartDate from WOR1 inner join OWOR on OWOR.DocEntry = WOR1.DocEntry "
                            sql = sql + "Where WOR1.DocEntry ='" + ProdDocEntry.ToString() + "' order by Wor1.linenum  "
                            rs2.DoQuery(sql)

                            Dim HuskNyStartDateFromPlan = rs2.Fields.Item("StartDate").Value

                            HuskNyStartDate = HuskNyStartDateFromPlan

                            If HuskLowestNyStartDate > HuskNyStartDate Then
                                HuskLowestNyStartDate = HuskNyStartDate
                            End If
                            'check if have fahter
                            HuskFather = ""

                            sql = "select wor1.Itemcode from wor1 where itemType = 4 And Wor1.docentry = (Select top 1 wor1.DocEntry from wor1 inner join owor On owor.docentry = wor1.docentry where itemType = 4 And owor.originAbs = " + salesOrderDocEntry.ToString() + "  And owor.U_RCS_BVO = " + rs3.Fields.Item("Visorder").Value.ToString() + " And wor1.itemcode = '" + rs2.Fields.Item("Father").Value.ToString() + "')"

                            rs2.DoQuery(sql)
                            If rs2.RecordCount > 0 Then

                                NextIsParallel = True
                                Do Until rs2.EoF
                                    HuskFather = HuskFather + "§" + rs2.Fields.Item("ItemCode").Value + ""
                                    rs2.MoveNext()
                                Loop

                                sql = "select DueDate from OWOR "
                                sql = sql + "Where DocEntry ='" + ProdDocEntry.ToString() + "'"
                                rs2.DoQuery(sql)
                                HuskNyEndDate = rs2.Fields.Item("DueDate").Value

                                If HuskNyEndDate > HuskFatherStartDate Then
                                    PlanOk = False

                                    MoveDate = DateDiff(DateInterval.Day, HuskFatherStartDate, HuskNyEndDate)
                                    AfsDate = AfsDate.AddDays(MoveDate)

                                    sql = "Update RDR1 set U_RCS_AFS = '" + AfsDate.ToString("yyyyMMdd") + "'"
                                    sql = sql + " Where VisOrder = " + rs3.Fields.Item("Visorder").Value.ToString + " and DocEntry = " + salesOrderDocEntry.ToString
                                    rs4.DoQuery(sql)

                                    'SBO_Application.SetStatusBarMessage("Prøver med afsendelsesdato: " + AfsDate.ToString("dd-MM-yyyy"))
                                    WriteInConsole("Prøver med afsendelsesdato: " + AfsDate.ToString("dd-MM-yyyy"))
                                    Exit Do
                                End If
                            Else
                                HuskNyEndDate = HuskLowestNyStartDate
                                HuskFatherStartDate = HuskNyEndDate
                            End If

                            rs1.MoveNext()
                            '*************************** REMOVE *************************
                            'Return response 'For TEST only
                            '*************************** REMOVE *************************
                        Loop

                    Loop

                    rs3.MoveNext()
                Loop
                '//*********************

                Return response
            End If

            response.Message += $"Salgsordre med DocEntry: {salesOrderDocEntry} er ikke fundet"
            response.Success = False
            Return response
        Catch ex As Exception
            ErrLog("PlanAllProductionOrders", ex)

            response.Success = False
            response.Message += $"{vbCrLf}PlanAllProductionOrders, message: {ex.Message}  exception: {ex.ToString}"
            Return response
        Finally
            oSO.ReleaseComObject
            oProd.ReleaseComObject
            rs.ReleaseComObject
            rs1.ReleaseComObject
            rs2.ReleaseComObject
            rs3.ReleaseComObject
            rs4.ReleaseComObject
        End Try
    End Function
#End Region


#Region "Plan Production Order (Planlæg Rute datoer og Planlæg Ubegrænset i Produktionsordre)"
    Public Function UpdateProductionOrderLineStartEndDateUnLimited(ByVal planUnlimeted As Boolean, ByVal prodDocNum As Integer) As Response

        Dim response As Response = New Response

        Try
            Dim rs As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
            rs.DoQuery($"select distinct DocEntry, OriginAbs, U_RCS_BVO from owor Where DocNum = {prodDocNum.ToString}")

            UpdateProductionOrderLineStartEndDateNew(rs.Fields.Item("DocEntry").Value, prodDocNum.ToString, planUnlimeted, rs.Fields.Item("OriginAbs").Value, rs.Fields.Item("U_RCS_BVO").Value)
            Return response

        Catch ex As Exception
            ErrLog("UpdateProductionOrderLineStartEndDateUnLimited", ex)

            response.Success = False
            response.Message += $"UpdateProductionOrderLineStartEndDateUnLimited, message: {ex.Message}  exception: {ex.ToString}"
            Return response
        End Try

    End Function
#End Region


    'Public Sub UpdateProductionOrderLineStartEndDateOrig(ByVal ProdDocEntry As Integer, ByVal ProdDocNum As String, ByVal PlanUnlimeted As Boolean, salesOrderDocEntry As Long, VisOrder As Integer)

    '    Dim test As Object = ""
    '    Dim MaxPlanDays As Integer = Settings.Default.MaxPlanDays

    '    Dim rs As SAPbobsCOM.Recordset
    '    Dim rs2 As SAPbobsCOM.Recordset
    '    Dim rs3 As SAPbobsCOM.Recordset
    '    Dim rs4 As SAPbobsCOM.Recordset
    '    Dim rs5 As SAPbobsCOM.Recordset
    '    Dim rs6 As SAPbobsCOM.Recordset

    '    rs = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    'rs1 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    rs2 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    rs3 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    rs4 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    rs5 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    rs6 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '    Dim oProd As SAPbobsCOM.ProductionOrders
    '    oProd = vCmp.GetBusinessObject(BoObjectTypes.oProductionOrders)
    '    Dim sb As New System.Text.StringBuilder
    '    Try


    '        Dim iResult As Integer = 0
    '        Dim Msg As String = ""


    '        Dim sql As String
    '        Dim PlanOk As Boolean = False



    '        Do Until PlanOk = True
    '            PlanOk = True
    '            sql = "select distinct DocEntry from owor  "
    '            sql = sql + "Where DocNum = " + ProdDocNum + ""


    '            rs.DoQuery(sql)

    '            If rs.EoF Then
    '                Exit Try
    '            End If

    '            Dim PlannedQtyTotalHour As Double

    '            sql = "Select sum(case when PlannedQty < 1 then 1 else PlannedQty end) as PlannedQtyTotal from WOR1 inner join orsc on wor1.itemcode = orsc.rescode "
    '            sql = sql + "Where (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and ItemType = 290 and DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
    '            rs2.DoQuery(sql)


    '            sql = "Delete from ""@RCS_CAP_COMMIT"" where U_DocEntry = " + rs.Fields.Item("DocEntry").Value.ToString + ""
    '            rs3.DoQuery(sql)

    '            PlannedQtyTotalHour = rs2.Fields.Item("PlannedQtyTotal").Value / 60

    '            Dim DelayTotal As Double
    '            'Dim WeekEndTotal As Double
    '            Dim CntDay As Double = 0
    '            Dim Delay As Double = 0
    '            Dim Day As String = ""
    '            Dim WeekendDays As Integer = 0

    '            sql = "Select sum(1) as DelayTotal from WOR1 inner join orsc on wor1.itemcode = orsc.rescode "
    '            sql = sql + "Where (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and  wor1.ItemType = 290 and wor1.DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
    '            rs2.DoQuery(sql)
    '            DelayTotal = DelayTotal + rs2.Fields.Item("DelayTotal").Value

    '            Dim ResTotal As Double

    '            sql = "Select count(DocEntry) as ResTotal from WOR1  inner join orsc on wor1.itemcode = orsc.rescode "
    '            sql = sql + "Where (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and ItemType = 290 and DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
    '            rs2.DoQuery(sql)

    '            ResTotal = rs2.Fields.Item("ResTotal").Value



    '            '  sql = "Select WOR1.linenum, WOR1.StageId, wor4.SeqNum-1 as SeqNum ,WOR1.PlannedQty, WOR1.VisOrder, WOR1.U_RCS_Issued, WOR1.Itemcode, 0 as U_RCS_ND from WOR1 inner join orsc on wor1.itemcode = orsc.rescode "
    '            '  sql = sql + " inner join wor4 on wor4.stageid = wor1.stageid and wor1.docentry = wor4.docentry "
    '            ' sql = sql + "Where wor1.ItemType = 290 and wor1.DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
    '            ' sql = sql + " order by VisOrder"

    '            'sql = "Select  ORSC.ResGrpCod, wor1.StageId , wor4.SeqNum-1 As SeqNum ,sum(WOR1.PlannedQty+OWOR.PlannedQty*isnull(U_RCS_ESKM,0)) as PQty, sum(WOR1.U_RCS_Issued), sum(isnull(U_RCS_ED,0)) as Nd, orsc.QryGroup7, orsc.QryGroup6, orsc.QryGroup8, isnull(U_RCS_PU,'') as U_RCS_PU,  isnull(U_RCS_DEL,'') as U_RCS_DEL  from WOR1 "
    '            'sql = sql + "inner Join OWOR on OWOR.DocEntry =  wor1.docentry inner Join wor4 on wor4.stageid = wor1.stageid And wor1.docentry = wor4.docentry inner join orsc on wor1.itemcode = orsc.rescode left join ocrd on SubString(wor1.itemcode,3,2) = OCRD.U_RCS_CONO "
    '            'sql = sql + "Where  (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and  wor1.ItemType = 290 And wor1.DocEntry  =" + rs.Fields.Item("DocEntry").Value.ToString + " "
    '            'sql = sql + "Group by SeqNum , wor1.StageId, orsc.QryGroup7, orsc.QryGroup6, orsc.QryGroup8, isnull(U_RCS_PU,''),  isnull(U_RCS_DEL,''),  ORSC.ResGrpCod "
    '            'sql = sql + "order by SeqNum "

    '            ' BR comment out 20220321
    '            'sql = "Select  ORSC.ResGrpCod, wor1.StageId , wor4.SeqNum-1 As SeqNum ,sum(WOR1.PlannedQty+OWOR.PlannedQty*isnull(U_RCS_ESKM,0)) as PQty, sum(isnull(WOR1.U_RCS_Issued,0)) U_RCS_Issued, sum(isnull(U_RCS_ED,0)) as Nd, orsc.QryGroup7, orsc.QryGroup14, orsc.QryGroup6,  max(Case When orsc.QryGroup8 = 'Y' then 1 else 0 end) as QryGroup8, isnull(U_RCS_PU,'') as U_RCS_PU,  isnull(U_RCS_DEL,'') as U_RCS_DEL  from WOR1 "
    '            sql = "Select  ORSC.ResGrpCod, wor1.StageId , wor4.SeqNum-1 As SeqNum ,sum(WOR1.PlannedQty+OWOR.PlannedQty*isnull(U_RCS_ESKM,0)-WOR1.IssuedQty) as PQty, sum(isnull(WOR1.U_RCS_Issued,0)) U_RCS_Issued, sum(isnull(U_RCS_ED,0)) as Nd, orsc.QryGroup7, orsc.QryGroup14, orsc.QryGroup6,  max(Case When orsc.QryGroup8 = 'Y' then 1 else 0 end) as QryGroup8, isnull(U_RCS_PU,'') as U_RCS_PU,  isnull(U_RCS_DEL,'') as U_RCS_DEL  from WOR1 "
    '            sql = sql + "inner Join OWOR on OWOR.DocEntry =  wor1.docentry inner Join wor4 on wor4.stageid = wor1.stageid And wor1.docentry = wor4.docentry inner join orsc on wor1.itemcode = orsc.rescode left join ocrd on SubString(wor1.itemcode,3,2) = OCRD.U_RCS_CONO "
    '            sql = sql + "Where  (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and  wor1.ItemType = 290 And wor1.DocEntry  =" + rs.Fields.Item("DocEntry").Value.ToString + " "
    '            sql = sql + "Group by SeqNum , wor1.StageId, orsc.QryGroup7, orsc.QryGroup14, orsc.QryGroup6,  isnull(U_RCS_PU,''),  isnull(U_RCS_DEL,''),  ORSC.ResGrpCod "
    '            sql = sql + "order by SeqNum "

    '            '  Select Case ORSC.ResGrpCod, wor1.StageId , wor4.SeqNum-1 As SeqNum ,sum(WOR1.PlannedQty+OWOR.PlannedQty*isnull(U_RCS_ESKM,0)) As PQty, 
    '            'sum(isnull(WOR1.U_RCS_Issued, 0)) As U_RCS_Issued, sum(isnull(U_RCS_ED,0)) As Nd, orsc.QryGroup7, orsc.QryGroup14, orsc.QryGroup6, max(Case When orsc.QryGroup8 = 'Y' then 1 else 0 end) as QryGroup8,
    '            'isnull(U_RCS_PU,'') as U_RCS_PU,  isnull(U_RCS_DEL,'') as U_RCS_DEL, wor4.StartDate, wor4.endDate  from WOR1 inner Join OWOR on OWOR.DocEntry =  wor1.docentry 
    '            'inner Join wor4 on wor4.stageid = wor1.stageid And wor1.docentry = wor4.docentry inner join orsc on wor1.itemcode = orsc.rescode 
    '            'Left Join ocrd On SubString(wor1.itemcode, 3, 2) = OCRD.U_RCS_CONO Where  (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or
    '            'orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and  wor1.ItemType = 290 And wor1.DocEntry  =12111 Group by SeqNum , wor1.StageId,
    '            'orsc.QryGroup7, orsc.QryGroup14, orsc.QryGroup6, isnull(U_RCS_PU,''),  isnull(U_RCS_DEL,''),  ORSC.ResGrpCod, wor4.StartDate, wor4.endDate order by SeqNum 

    '            rs2.DoQuery(sql)
    '            'Dim DateCnt As Double
    '            Dim DateCntPlanned As Double
    '            Dim DateCntPlannedPerDate As Double
    '            'Dim CapPerDate As Double
    '            Dim PlannedQty As Double

    '            Dim HuskProdDateStartFromStart As Date
    '            Dim HuskProdDateEndLatest As Date

    '            If oProd.GetByKey(rs.Fields.Item("DocEntry").Value) Then

    '                HuskProdDateStartFromStart = oProd.StartDate
    '                HuskProdDateStartNy = oProd.StartDate
    '                HuskProdDateEndNy = oProd.DueDate
    '                HuskProdDateEndLatest = oProd.DueDate
    '                '  HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)




    '                'DateCnt = DateCnt - WeekendDays

    '                'DateCnt = DateCnt - (DelayTotal - 1)

    '                '  CapPerDate = DateCnt / (PlannedQtyTotalHour)
    '                Dim Cnt As Integer = rs2.RecordCount - 1
    '                Dim CntRecHusk As Integer = rs2.RecordCount - 1
    '                Dim HuskResGrpcode As Integer = 0
    '                Dim SameDate As Boolean = False
    '                Dim HuskQtySameDate As Double = 0
    '                Dim HuskSameDateOneShot As Boolean = False
    '                Dim HPD As Double = 8
    '                Dim HuskStartSlutDato As String = ""
    '                Dim HuskStartSlutDatoUB As String = ""
    '                'Dim testday

    '                rs2.MoveLast()
    '                Do Until Cnt < 0
    '                    Dim PQty As Double = rs2.Fields.Item("PQty").Value
    '                    Dim temp As String = rs2.Fields.Item("StageId").Value

    '                    If HuskResGrpcode = rs2.Fields.Item("ResGrpCod").Value Then

    '                        If oProd.Stages.EndDate = oProd.Stages.StartDate Then
    '                            SameDate = True
    '                            HuskSameDateOneShot = True
    '                        Else
    '                            SameDate = False



    '                        End If
    '                    Else
    '                        SameDate = False

    '                        'rs 20210927
    '                        ' If HuskSameDateOneShot = False Then

    '                        Dim testDate As Date = oProd.Stages.StartDate
    '                        If HuskProdDateEndNy = oProd.Stages.StartDate Then
    '                            HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                            Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
    '                            If Day = "Sat" Then
    '                                HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                            End If
    '                            If Day = "Sun" Then
    '                                HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
    '                            End If
    '                        End If

    '                        ' End If
    '                        'HuskSameDateOneShot = False
    '                        ' If HuskSameDateOneShot = True Then
    '                        'HuskSameDateOneShot = False
    '                        'If rs2.Fields.Item("QryGroup7").Value <> "Y" Then
    '                        '    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                        '    If Day = "Sat" Then
    '                        '        HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                        '    End If
    '                        '    If Day = "Sun" Then
    '                        '        HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
    '                        '    End If
    '                        'End If

    '                        '  End If


    '                    End If




    '                    '   HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-rs2.Fields.Item("Nd").Value)
    '                    Dim CntHusk As Integer = rs2.Fields.Item("Nd").Value
    '                    If rs2.Fields.Item("Nd").Value > 0 And 1 = 2 Then
    '                        WeekendDays = 0
    '                        CntHusk = rs2.Fields.Item("Nd").Value
    '                        CntDay = CntHusk
    '                        Dim CntDay2a As Double = rs2.Fields.Item("Nd").Value
    '                        Dim CntDayHuska As Double = rs2.Fields.Item("Nd").Value
    '                        If CntDay > 1 Then
    '                            Do Until (CntDay + WeekendDays <= 0)

    '                                Day = HuskProdDateEndNy.AddDays(CntDayHuska - CntDay2a - 1).DayOfWeek.ToString().Substring(0, 3)
    '                                If Day = "Sat" Or Day = "Sun" Then

    '                                    WeekendDays = WeekendDays + 1

    '                                End If
    '                                CntDay = CntDay - 1
    '                                CntDay2a = CntDay2a + 1

    '                                'you have to check day equal to "sat" Or "sun".
    '                            Loop
    '                        End If

    '                        HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-(CntHusk) - WeekendDays)


    '                    End If



    '                    ' Dim testVisOrder As String = rs2.Fields.Item("VisOrder").Value
    '                    '  oProd.Lines.SetCurrentLine(rs2.Fields.Item("VisOrder").Value)
    '                    Dim stdate = oProd.Stages.StartDate
    '                    Dim endate = oProd.Stages.EndDate
    '                    '   oProd.Update()

    '                    Dim testseqNum = rs2.Fields.Item("SeqNum").Value
    '                    oProd.Stages.SetCurrentLine(rs2.Fields.Item("SeqNum").Value)


    '                    'testday = oProd.Stages.EndDate
    '                    If HuskResGrpcode = rs2.Fields.Item("ResGrpCod").Value Then
    '                        sql = "Select isnull(U_RCS_NQD,'N') as NQD from ORSB "
    '                        sql = sql + "Where ResGrpCod = " + HuskResGrpcode.ToString()

    '                        rs4.DoQuery(sql)
    '                        Dim NQD As String
    '                        NQD = rs4.Fields.Item("NQD").Value

    '                        If NQD = "N" Then 'Ikke samme dato
    '                            If SameDate = True Then
    '                                HuskProdDateEndNy = HuskProdDateEndLatest
    '                            End If
    '                        End If



    '                    Else
    '                        HuskQtySameDate = 0
    '                        sql = "Select isnull(U_RCS_NQD,'N') as NQD from ORSB "
    '                        sql = sql + "Where ResGrpCod = " + HuskResGrpcode.ToString()

    '                        rs4.DoQuery(sql)
    '                        Dim NQD As String
    '                        NQD = rs4.Fields.Item("NQD").Value

    '                        If NQD = "Y" Then 'Leverandør
    '                            HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)

    '                        End If


    '                        HuskResGrpcode = rs2.Fields.Item("ResGrpCod").Value
    '                    End If

    '                    sql = "Select isnull(U_RCS_HD,8) as HPD from ORSB "
    '                    sql = sql + "Where ResGrpCod = " + HuskResGrpcode.ToString()

    '                    rs4.DoQuery(sql)

    '                    HPD = rs4.Fields.Item("HPD").Value
    '                    If HPD < 1 Then ' BR
    '                        HPD = 8
    '                    End If

    '                    PlannedQty = rs2.Fields.Item("PQty").Value / 60
    '                    If PlannedQty < 1 Then
    '                        PlannedQty = 1
    '                    End If


    '                    If rs2.Fields.Item("QryGroup7").Value = "Y" Then

    '                        'DateCntPlannedPerDate = (PlannedQty / 8) + 1 ' Antal dage
    '                        DateCntPlannedPerDate = (PlannedQty / 8) '+ 1 ' Antal dage

    '                    ElseIf rs2.Fields.Item("QryGroup6").Value = "Y" Then
    '                        DateCntPlannedPerDate = (PlannedQty / 8) + 1 ' Antal dage
    '                    Else
    '                        DateCntPlannedPerDate = PlannedQty / HPD ' Antal dage
    '                    End If


    '                    DateCntPlannedPerDate = Math.Ceiling(DateCntPlannedPerDate)

    '                    DateCntPlannedPerDate = DateCntPlannedPerDate + rs2.Fields.Item("Nd").Value
    '                    CntDay = DateCntPlannedPerDate







    '                    DateCntPlanned = DateCntPlannedPerDate

    '                    DateCntPlanned = Math.Ceiling(DateCntPlanned)
    '                    If rs2.Fields.Item("Nd").Value = 0 Then
    '                        If DateCntPlanned <= 1 Then
    '                            'rs 20210927
    '                            ' DateCntPlanned = 2
    '                            DateCntPlanned = 1
    '                        End If
    '                    Else
    '                        If DateCntPlanned <= 1 Then
    '                            DateCntPlanned = 1
    '                        End If
    '                    End If



    '                    Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)



    '                    If Day = "Sat" Then
    '                        HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                    End If
    '                    If Day = "Sun" Then
    '                        HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
    '                    End If

    '                    ' check for delevery
    '                    If rs2.Fields.Item("QryGroup6").Value = "Y" Then
    '                        Dim DayNum As String

    '                        Day = HuskProdDateEndNy.DayOfWeek.ToString()

    '                        Dim startdayvalue As String = "0"
    '                        Select Case Day.ToLower
    '                            Case "monday"
    '                                startdayvalue = "1"
    '                            Case "tuesday"
    '                                startdayvalue = "2"
    '                            Case "wednesday"
    '                                startdayvalue = "3"
    '                            Case "thursday"
    '                                startdayvalue = "4"
    '                            Case "friday"
    '                                startdayvalue = "5"
    '                            Case "saturday"
    '                                startdayvalue = "6"
    '                            Case "sunday"
    '                                startdayvalue = "7"
    '                        End Select



    '                        DayNum = startdayvalue

    '                        Dim DelDays As String = rs2.Fields.Item("U_RCS_DEL").Value


    '                        Dim CntDel As Integer = 0
    '                        Dim status As Boolean = False
    '                        Do Until CntDel > 1

    '                            Do Until Int(DayNum) <= 0
    '                                If InStr(DelDays, Trim(DayNum)) > 0 Then
    '                                    ' all good
    '                                    status = True
    '                                    Exit Do
    '                                Else
    '                                    DayNum = Str(Int(DayNum) - 1)
    '                                    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                                    Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
    '                                    If Day = "Sat" Then
    '                                        HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                                    End If
    '                                    If Day = "Sun" Then
    '                                        HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
    '                                    End If
    '                                End If

    '                            Loop
    '                            If status = True Then
    '                                Exit Do
    '                            End If
    '                            HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                            Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
    '                            If Day = "Sat" Then
    '                                HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                            End If
    '                            If Day = "Sun" Then
    '                                HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
    '                            End If
    '                            CntDel = CntDel + 1
    '                            DayNum = "5"
    '                        Loop

    '                        If Cnt < CntRecHusk Then
    '                            oProd.Stages.SetCurrentLine(rs2.Fields.Item("SeqNum").Value + 1)
    '                            oProd.Stages.StartDate = HuskProdDateEndNy
    '                            oProd.Stages.SetCurrentLine(rs2.Fields.Item("SeqNum").Value)
    '                        End If

    '                        If rs2.Fields.Item("QryGroup8").Value = "1" Then

    '                            oProd.Stages.StartDate = HuskProdDateEndNy.AddDays(-7).AddDays(CntHusk)
    '                            HuskProdDateEndNy = oProd.Stages.StartDate
    '                            DateCntPlanned = 1 + 5
    '                        End If

    '                    End If

    '                    CntDay = DateCntPlanned

    '                    Dim testdate4 As Date = oProd.Stages.EndDate
    '                    'rs 20220207
    '                    oProd.Stages.EndDate = HuskProdDateEndNy

    '                    HuskProdDateEndLatest = HuskProdDateEndNy
    '                    '   CntDay = DateCntPlanned
    '                    WeekendDays = 0
    '                    Dim CntDay2 As Double = CntDay
    '                    Dim CntDayHusk As Double = CntDay
    '                    If CntDay > 1 Then
    '                        Do Until ((CntDay - 1 + WeekendDays) <= 0)

    '                            Day = HuskProdDateEndNy.AddDays(CntDayHusk - CntDay2 - 1).DayOfWeek.ToString().Substring(0, 3)
    '                            If Day = "Sat" Or Day = "Sun" Then

    '                                WeekendDays = WeekendDays + 1

    '                            End If
    '                            CntDay = CntDay - 1
    '                            CntDay2 = CntDay2 + 1

    '                            'you have to check day equal to "sat" Or "sun".
    '                        Loop
    '                    End If


    '                    Day = oProd.Stages.EndDate.DayOfWeek.ToString().Substring(0, 3)


    '                    Day = oProd.Stages.EndDate.AddDays(-(DateCntPlanned + WeekendDays - 1)).DayOfWeek.ToString().Substring(0, 3)

    '                    If Day = "Sat" Then
    '                        oProd.Stages.StartDate = oProd.Stages.EndDate.AddDays(-(DateCntPlanned + WeekendDays - 1) - 1)
    '                        HuskProdDateEndNy = oProd.Stages.StartDate
    '                    End If
    '                    If Day = "Sun" Then
    '                        oProd.Stages.StartDate = oProd.Stages.EndDate.AddDays(-(DateCntPlanned + WeekendDays - 1) - 2)
    '                        HuskProdDateEndNy = oProd.Stages.StartDate
    '                    End If
    '                    If Day <> "Sun" And Day <> "Sat" Then
    '                        oProd.Stages.StartDate = oProd.Stages.EndDate.AddDays(-(DateCntPlanned + WeekendDays - 1))
    '                        HuskProdDateEndNy = oProd.Stages.StartDate
    '                    End If



    '                    ' check for for pickup
    '                    If rs2.Fields.Item("QryGroup6").Value = "Y" And rs2.Fields.Item("QryGroup8").Value = "1" Then
    '                        Dim DayNum As String

    '                        Day = HuskProdDateEndNy.DayOfWeek.ToString()

    '                        Dim startdayvalue As String = "0"
    '                        Select Case Day.ToLower
    '                            Case "monday"
    '                                startdayvalue = "1"
    '                            Case "tuesday"
    '                                startdayvalue = "2"
    '                            Case "wednesday"
    '                                startdayvalue = "3"
    '                            Case "thursday"
    '                                startdayvalue = "4"
    '                            Case "friday"
    '                                startdayvalue = "5"
    '                            Case "saturday"
    '                                startdayvalue = "6"
    '                            Case "sunday"
    '                                startdayvalue = "7"
    '                        End Select



    '                        DayNum = startdayvalue

    '                        Dim DelDays As String = rs2.Fields.Item("U_RCS_PU").Value


    '                        Dim CntDel As Integer = 0
    '                        Dim status As Boolean = False
    '                        Do Until CntDel > 1

    '                            Do Until Int(DayNum) <= 0
    '                                If InStr(DelDays, Trim(DayNum)) > 0 Then
    '                                    ' all good
    '                                    status = True
    '                                    Exit Do
    '                                Else
    '                                    DayNum = Str(Int(DayNum) - 1)
    '                                    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                                    Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
    '                                    If Day = "Sat" Then
    '                                        HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                                    End If
    '                                    If Day = "Sun" Then
    '                                        HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
    '                                    End If
    '                                End If

    '                            Loop
    '                            If status = True Then
    '                                Exit Do
    '                            End If
    '                            '  HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                            'Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
    '                            'If Day = "Sat" Then
    '                            '    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                            'End If
    '                            'If Day = "Sun" Then
    '                            '    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
    '                            'End If
    '                            CntDel = CntDel + 1
    '                            DayNum = "5"
    '                        Loop

    '                        oProd.Stages.StartDate = HuskProdDateEndNy
    '                    End If

    '                    'rs 20210927 
    '                    Dim HuskNewEndDate As Date
    '                    If rs2.Fields.Item("QryGroup7").Value = "Y" Or 1 = 1 Then
    '                        HuskNewEndDate = oProd.Stages.EndDate
    '                        If oProd.Stages.StartDate >= oProd.Stages.EndDate Then
    '                            HuskNewEndDate = oProd.Stages.StartDate
    '                        End If

    '                    Else
    '                        HuskNewEndDate = oProd.Stages.EndDate.AddDays(-1)
    '                        If oProd.Stages.StartDate >= oProd.Stages.EndDate.AddDays(-1) Then
    '                            HuskNewEndDate = oProd.Stages.StartDate
    '                        End If

    '                    End If


    '                    ' Dim HuskNewEndDate As Date = oProd.Stages.EndDate.AddDays(-1)
    '                    'If oProd.Stages.StartDate >= oProd.Stages.EndDate.AddDays(-1) Then
    '                    'HuskNewEndDate = oProd.Stages.StartDate
    '                    'End If

    '                    sql = "Select SeqNum-1 as pos From wor4 Where DocEntry = " + oProd.AbsoluteEntry.ToString + " and StageID = " + oProd.Stages.StageID.ToString
    '                    rs.DoQuery(sql)
    '                    Dim pos As String = rs.Fields.Item("pos").Value


    '                    sql = "select sum(ORCJ.Capacity) as Cap From ORCJ "
    '                    sql = sql + "Join ORSC On ORSC.ResCode = ORCJ.ResCode "
    '                    sql = sql + "Where ORSC.ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
    '                    sql = sql + "And  ORCJ.CapDate BETWEEN '" + oProd.Stages.StartDate.ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' "
    '                    sql = sql + "And ORCJ.CapType = 'I' "
    '                    sql = sql + " And ORCJ.WhsCode = '01' "
    '                    rs.DoQuery(sql)



    '                    sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry "
    '                    sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
    '                    sql = sql + "And  U_Date BETWEEN '" + oProd.Stages.StartDate.ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"


    '                    rs4.DoQuery(sql)


    '                    Dim test1, test2, test3, testdate1, testdate2
    '                    test1 = rs2.Fields.Item("PQty").Value
    '                    test2 = rs.Fields.Item("Cap").Value
    '                    test3 = rs4.Fields.Item("ComCap").Value




    '                    '  pos = (oProd.Stages.StageID - 1).ToString

    '                    testdate1 = oProd.Stages.StartDate.ToString("MM-dd-yyyy")
    '                    testdate2 = HuskNewEndDate.ToString("MM-dd-yyyy")

    '                    Dim HuskPosEndDate As String = HuskNewEndDate.ToString("dd-MM-yyyy")
    '                    HuskStartSlutDatoUB = HuskStartSlutDatoUB + " Pos: " + pos.ToString + " Dato: " + oProd.Stages.StartDate.ToString("dd-MM-yyyy") + ":" + HuskNewEndDate.ToString("dd-MM-yyyy")


    '                    '  If 1 = 1 And rs2.Fields.Item("PQty").Value > (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value - HuskQtySameDate) And PlanUnlimeted = False Then
    '                    'rs 2021-12-11 If 1 = 1 And PlanUnlimeted = False Or rs2.Fields.Item("PQty").Value > (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value - HuskQtySameDate) And PlanUnlimeted = False Then
    '                    'If rs2.Fields.Item("PQty").Value > (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value - HuskQtySameDate) Then
    '                    If 1 = 1 Then

    '                        Dim CntDay3 As Integer = 0
    '                        Dim Cnt2 As Integer = 0
    '                        Dim sql1 As String = ""
    '                        Dim Cnt4 As Integer = 0
    '                        Dim BookState As Boolean = False
    '                        Dim CanStartBook As Boolean = False
    '                        Dim CapPerDayTotal As Integer = 0
    '                        Dim CapPerDayTotalFirstRun As Integer = 0

    '                        Dim firtsRun As Boolean = True

    '                        sql = "select sum(ORCJ.Capacity) as Cap From ORCJ "
    '                        sql = sql + "Join ORSC On ORSC.ResCode = ORCJ.ResCode "
    '                        sql = sql + "Where ORSC.ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
    '                        sql = sql + "And  ORCJ.CapDate BETWEEN '" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' "
    '                        sql = sql + "And ORCJ.CapType = 'I' "
    '                        sql = sql + " And ORCJ.WhsCode = '01' "
    '                        rs5.DoQuery(sql)
    '                        test2 = rs5.Fields.Item("Cap").Value

    '                        sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
    '                        sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
    '                        sql = sql + "And  U_Date BETWEEN '" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"

    '                        rs6.DoQuery(sql)


    '                        test3 = rs6.Fields.Item("ComCap").Value

    '                        '   Do Until (rs2.Fields.Item("PQty").Value <= (rs5.Fields.Item("Cap").Value - rs6.Fields.Item("ComCap").Value - HuskQtySameDate)) And (rs2.Fields.Item("PQty").Value <= CapPerDayTotal)
    '                        '  Do Until (rs2.Fields.Item("PQty").Value <= (rs5.Fields.Item("Cap").Value - rs6.Fields.Item("ComCap").Value - HuskQtySameDate)) Or (rs2.Fields.Item("PQty").Value <= CapPerDayTotal)
    '                        Do Until (rs2.Fields.Item("PQty").Value <= CapPerDayTotal)

    '                            ' SBO_Application.SetStatusBarMessage("Ikke nok resource til artikel nr. " + oProd.ItemNo + " pos: " + oProd.Stages.Name + " Planlagt: " + rs2.Fields.Item("PQty").Value.ToString + " muligt forbrug: " + (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value).ToString + "I perioden " + oProd.Stages.StartDate.AddDays(CntDay3).ToString() + " " + HuskNewEndDate.ToString())


    '                            sql = "select sum(ORCJ.Capacity) as Cap From ORCJ "
    '                            sql = sql + "Join ORSC On ORSC.ResCode = ORCJ.ResCode "
    '                            sql = sql + "Where ORSC.ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
    '                            sql = sql + "And  ORCJ.CapDate BETWEEN '" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' "
    '                            sql = sql + "And ORCJ.CapType = 'I' "
    '                            sql = sql + " And ORCJ.WhsCode = '01' "
    '                            rs5.DoQuery(sql)
    '                            test2 = rs5.Fields.Item("Cap").Value

    '                            sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
    '                            sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
    '                            sql = sql + "And  U_Date BETWEEN '" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"

    '                            rs6.DoQuery(sql)


    '                            test3 = rs6.Fields.Item("ComCap").Value








    '                            sql1 = "select sum(ORCJ.Capacity) as Cap From ORCJ "
    '                            sql1 = sql1 + "Join ORSC On ORSC.ResCode = ORCJ.ResCode "
    '                            sql1 = sql1 + "Where ORSC.ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
    '                            sql1 = sql1 + "And  ORCJ.CapDate BETWEEN '" + HuskNewEndDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' "
    '                            sql1 = sql1 + "And ORCJ.CapType = 'I' "
    '                            sql1 = sql1 + " And ORCJ.WhsCode = '01' "
    '                            rs.DoQuery(sql1)
    '                            test2 = rs.Fields.Item("Cap").Value


    '                            sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
    '                            sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
    '                            sql = sql + "And  U_Date BETWEEN '" + HuskNewEndDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"

    '                            rs4.DoQuery(sql)

    '                            test1 = rs2.Fields.Item("PQty").Value
    '                            test2 = rs.Fields.Item("Cap").Value
    '                            test3 = rs4.Fields.Item("ComCap").Value

    '                            'if capacity > comitted capacity
    '                            If (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value) > 0 Then

    '                                ' if free capacity > max capacity per day
    '                                If (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value) > (HPD * 60) Then

    '                                    CapPerDayTotal = CapPerDayTotal + (HPD * 60)

    '                                Else
    '                                    If PlanUnlimeted = True And rs.Fields.Item("Cap").Value > 0 Then
    '                                        CapPerDayTotal = CapPerDayTotal + (HPD * 60)
    '                                    Else
    '                                        CapPerDayTotal = CapPerDayTotal + (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value)

    '                                    End If


    '                                End If

    '                                '  CntDay3 = CntDay3 - Math.Ceiling((rs2.Fields.Item("PQty").Value - (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value)) / (HPD * 60))
    '                                CntDay3 = CntDay3 - 1
    '                            Else
    '                                If PlanUnlimeted = True And rs.Fields.Item("Cap").Value > 0 Then
    '                                    CapPerDayTotal = CapPerDayTotal + (HPD * 60)

    '                                End If

    '                                CntDay3 = CntDay3 - 1
    '                            End If














    '                            HuskProdDateEndNy = HuskNewEndDate.AddDays(CntDay3 + 1) ' oProd.Stages.StartDate.AddDays(CntDay3)

    '                            Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
    '                            If Day = "Sat" Then
    '                                HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                            End If
    '                            If Day = "Sun" Then
    '                                HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
    '                            End If
    '                            '  rs.MoveNext()

    '                            ' check for for pickup
    '                            If rs2.Fields.Item("QryGroup6").Value = "Y" And rs2.Fields.Item("QryGroup8").Value = "1" Then
    '                                Dim DayNum As String

    '                                Day = HuskProdDateEndNy.DayOfWeek.ToString()

    '                                Dim startdayvalue As String = "0"
    '                                Select Case Day.ToLower
    '                                    Case "monday"
    '                                        startdayvalue = "1"
    '                                    Case "tuesday"
    '                                        startdayvalue = "2"
    '                                    Case "wednesday"
    '                                        startdayvalue = "3"
    '                                    Case "thursday"
    '                                        startdayvalue = "4"
    '                                    Case "friday"
    '                                        startdayvalue = "5"
    '                                    Case "saturday"
    '                                        startdayvalue = "6"
    '                                    Case "sunday"
    '                                        startdayvalue = "7"
    '                                End Select



    '                                DayNum = startdayvalue

    '                                Dim DelDays As String = rs2.Fields.Item("U_RCS_PU").Value


    '                                Dim CntDel As Integer = 0
    '                                Dim status As Boolean = False
    '                                Do Until CntDel > 1

    '                                    Do Until Int(DayNum) <= 0
    '                                        If InStr(DelDays, Trim(DayNum)) > 0 Then
    '                                            ' all good
    '                                            status = True
    '                                            Exit Do
    '                                        Else
    '                                            DayNum = Str(Int(DayNum) - 1)
    '                                            HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                                            Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
    '                                            If Day = "Sat" Then
    '                                                HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                                            End If
    '                                            If Day = "Sun" Then
    '                                                HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
    '                                            End If
    '                                        End If

    '                                    Loop
    '                                    If status = True Then
    '                                        Exit Do
    '                                    End If
    '                                    '  HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                                    'Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
    '                                    'If Day = "Sat" Then
    '                                    '    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
    '                                    'End If
    '                                    'If Day = "Sun" Then
    '                                    '    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
    '                                    'End If
    '                                    CntDel = CntDel + 1
    '                                    DayNum = "5"
    '                                Loop

    '                                ' oProd.Stages.StartDate = HuskProdDateEndNy
    '                            End If




    '                            'Check for passing startdate
    '                            'rs 2021-1´2-11 If DateDiff(DateInterval.Day, oProd.StartDate, HuskProdDateEndNy) < 0 Then
    '                            If (DateDiff(DateInterval.Day, Date.Now, HuskProdDateEndNy) < 0) Or (DateDiff(DateInterval.Day, oProd.StartDate, HuskProdDateEndNy) < 0) Then


    '                                ' Then start from begin with new end date moved 1 day
    '                                Dim newDueDate As Date = oProd.DueDate.AddDays(1)


    '                                Day = newDueDate.DayOfWeek.ToString().Substring(0, 3)
    '                                If Day = "Sat" Then
    '                                    newDueDate = newDueDate.AddDays(2)
    '                                End If
    '                                If Day = "Sun" Then
    '                                    newDueDate = newDueDate.AddDays(1)
    '                                End If


    '                                oProd.StartDate = Date.Now
    '                                oProd.DueDate = newDueDate
    '                                Dim cnt3 As Integer = 0

    '                                Do Until cnt3 > oProd.Stages.Count - 1
    '                                    oProd.Stages.SetCurrentLine(cnt3)
    '                                    oProd.Stages.StartDate = Date.Now
    '                                    oProd.Stages.EndDate = newDueDate

    '                                    cnt3 = cnt3 + 1
    '                                Loop


    '                                iResult = oProd.Update()
    '                                If iResult < 0 Then
    '                                    vCmp.GetLastError(iResult, Msg)
    '                                    ErrorLog("CreateCopyOfBillOfMaterial", Msg)
    '                                    'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
    '                                    WriteInConsole("Produktionsordre for artikel " + oProd.ItemNo.ToString + " fejlet: " + Msg)
    '                                Else
    '                                    Msg = "Produktionsordre for artikel " + oProd.ItemNo.ToString + " er opdateret med ny dato. Stopper planlæggning ved pos: " + pos.ToString + " med slutdato: " + HuskNewEndDate.ToString("dd-MM-yyyy")
    '                                    'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
    '                                    WriteInConsole(Msg)
    '                                End If


    '                                sql = "Update RDR1 set U_RCS_AFS = '" + oProd.DueDate.ToString("yyyyMMdd") + "'"
    '                                sql = sql + " Where VisOrder = " + VisOrder.ToString + " and DocEntry = " + salesOrderDocEntry.ToString
    '                                rs4.DoQuery(sql)

    '                                'SBO_Application.SetStatusBarMessage(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " Prøver med forfaldsdato: " + oProd.DueDate.ToString("dd-MM-yyyy"))
    '                                'SBO_Application.SetStatusBarMessage(HuskStartSlutDato)

    '                                'SBO_Application.StatusBar.SetText("Planlægning fejler. Artikelnr.: " + oProd.ItemNo.ToString, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
    '                                PlanOk = False


    '                                Exit Do
    '                            End If

    '                            'Check for passing startdate
    '                            'rs 2021-12-11 If DateDiff(DateInterval.Day, oProd.StartDate, HuskProdDateEndNy) < 0 Then
    '                            If (DateDiff(DateInterval.Day, Date.Now, HuskProdDateEndNy) < 0) Or (DateDiff(DateInterval.Day, oProd.StartDate, HuskProdDateEndNy) < 0) Then

    '                                SBO_Application.SetStatusBarMessage(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " Har passeret startdato stopper planlægning! " + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' mellem '" + HuskNewEndDate.ToString("MM-dd-yyyy"))
    '                                Cnt = 0
    '                                Exit Do
    '                            End If
    '                            'Check for max 60 days
    '                            If Cnt2 > MaxPlanDays Then
    '                                SBO_Application.SetStatusBarMessage(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + "Har passeret 30 dage stopper planlægning!" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' mellem '" + HuskNewEndDate.ToString("MM-dd-yyyy"))

    '                                Dim Tekst = "Skal bruge " + test1.ToString + " Der er kapacitet ialt " + test2.ToString + " Der er brugt " + test3.ToString

    '                                sql = "Select [@RCS_CAP_COMMIT].U_Date ,owor.docNum, round(Sum(U_Capacity)/60,2) as comcap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry "
    '                                sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
    '                                sql = sql + "And  U_Date BETWEEN '" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString
    '                                sql = sql + " group by [@RCS_CAP_COMMIT].U_Date,owor.DocNum "
    '                                sql = sql + " order by U_Date"
    '                                rs4.DoQuery(sql)

    '                                test1 = rs4.Fields.Item("U_Date").Value
    '                                test2 = rs4.Fields.Item("docNum").Value
    '                                test3 = rs4.Fields.Item("ComCap").Value

    '                                Do Until rs4.EoF
    '                                    test1 = rs4.Fields.Item("U_Date").Value
    '                                    test2 = rs4.Fields.Item("docNum").Value
    '                                    test3 = rs4.Fields.Item("ComCap").Value
    '                                    Tekst = Tekst + vbCrLf + test1.ToString + " " + test2.ToString + " " + test3.ToString
    '                                    rs4.MoveNext()
    '                                Loop
    '                                'SBO_Application.MessageBox(Tekst)
    '                                WriteInConsole(Tekst)
    '                                Cnt = 0
    '                                Exit Do
    '                            End If
    '                            Cnt2 = Cnt2 + 1
    '                        Loop
    '                        '  SBO_Application.SetStatusBarMessage("Ikke nok resource til artikel nr. " + oProd.ItemNo + " pos: " + oProd.Stages.Name + " Planlagt: " + rs2.Fields.Item("PQty").Value.ToString + " muligt forbrug: " + (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value).ToString + "I perioden " + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + " " + HuskNewEndDate.ToString("MM-dd-yyyy"))

    '                        'RS20220207  
    '                        oProd.Stages.StartDate = HuskProdDateEndNy
    '                        HuskStartSlutDato = HuskStartSlutDato + " Pos: " + pos.ToString + " Dato: " + HuskProdDateEndNy.ToString("dd-MM-yyyy") + ":" + HuskPosEndDate

    '                        'SBO_Application.SetStatusBarMessage("Ikke nok resource til artikel nr. " + oProd.ItemNo + " pos: " + oProd.Stages.Name + " Planlagt: " + rs2.Fields.Item("PQty").Value.ToString + " muligt forbrug: " + (test2 - test3).ToString)

    '                    End If
    '                    HuskQtySameDate = HuskQtySameDate + rs2.Fields.Item("PQty").Value

    '                    'SBO_Application.SetStatusBarMessage("Fundet resource til artikel nr. " + oProd.ItemNo + " Pos : " + oProd.Stages.Name + " Start: " + oProd.Stages.StartDate.ToString + " Slut: " + HuskNewEndDate.ToString, BoMessageTime.bmt_Long, False)
    '                    Dim HuskProdDateEndNy2 As Date
    '                    HuskProdDateEndNy2 = HuskProdDateEndNy
    '                    HuskProdDateEndNy2 = HuskProdDateEndNy2.AddDays(1)
    '                    Day = HuskProdDateEndNy2.DayOfWeek.ToString().Substring(0, 3)
    '                    If Day = "Sat" Then
    '                        HuskProdDateEndNy2 = HuskProdDateEndNy2.AddDays(-1)
    '                    End If
    '                    If Day = "Sun" Then
    '                        HuskProdDateEndNy2 = HuskProdDateEndNy2.AddDays(-2)
    '                    End If

    '                    'rs 20210927    
    '                    If rs2.Fields.Item("QryGroup7").Value = "Y" Then
    '                        ' oProd.Stages.StartDate = HuskProdDateEndNy2
    '                        '  oProd.Stages.StartDate = oProd.Stages.StartDate.AddDays(1)
    '                    End If

    '                    'oProd.Stages.StartDate = HuskProdDateEndNy2

    '                    oProd.Stages.EndDate = HuskNewEndDate

    '                    Cnt = Cnt - 1
    '                    If Cnt >= 0 Then
    '                        rs2.MovePrevious()

    '                    End If

    '                    testdate1 = oProd.Stages.StartDate.ToString("MM-dd-yyyy")
    '                    testdate2 = HuskNewEndDate.ToString("MM-dd-yyyy")

    '                    If PlanOk = False Then
    '                        Exit Do
    '                    End If
    '                Loop


    '                If PlanOk = True Then

    '                    Dim sd As Date = oProd.StartDate
    '                    Dim ssd As Date = oProd.Stages.StartDate

    '                    'If PlanUnlimeted = True And oProd.StartDate > oProd.Stages.StartDate Then
    '                    '    Dim Cnt2 As Integer = 0
    '                    '    Do Until Cnt2 > oProd.Stages.Count - 1
    '                    '        oProd.Stages.SetCurrentLine(Cnt2)
    '                    '        oProd.Stages.StartDate = Date.Now
    '                    '        oProd.Stages.EndDate = oProd.DueDate

    '                    '        Cnt2 = Cnt2 + 1
    '                    '    Loop

    '                    'End If
    '                    Dim test1a As Date = oProd.StartDate
    '                    Dim test2a As Date = oProd.Stages.StartDate

    '                    If oProd.StartDate > oProd.Stages.StartDate And PlanUnlimeted = False Then

    '                        Dim Missday

    '                        Missday = DateDiff(DateInterval.Day, oProd.StartDate, oProd.Stages.StartDate)

    '                        Msg = "Produktionsordre " + oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " er ikke opdateret! er forbi start dato!  "
    '                        'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
    '                        ' SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error)
    '                        'Exit Sub


    '                        ' Then start from begin with new end date moved 1 day
    '                        Dim newDueDate As Date = oProd.DueDate.AddDays(1)

    '                        Day = newDueDate.DayOfWeek.ToString().Substring(0, 3)
    '                        If Day = "Sat" Then
    '                            newDueDate = newDueDate.AddDays(2)
    '                        End If
    '                        If Day = "Sun" Then
    '                            newDueDate = newDueDate.AddDays(1)
    '                        End If


    '                        oProd.StartDate = Date.Now
    '                        oProd.DueDate = newDueDate
    '                        Dim cnt3 As Integer = 0

    '                        Do Until cnt3 > oProd.Stages.Count - 1
    '                            oProd.Stages.SetCurrentLine(cnt3)
    '                            oProd.Stages.StartDate = Date.Now
    '                            oProd.Stages.EndDate = newDueDate

    '                            cnt3 = cnt3 + 1
    '                        Loop


    '                        iResult = oProd.Update()
    '                        If iResult < 0 Then
    '                            vCmp.GetLastError(iResult, Msg)
    '                            ErrorLog("CreateCopyOfBillOfMaterial", Msg)
    '                            'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
    '                            WriteInConsole(Msg)
    '                        Else
    '                            Msg = "Produktionsordre for artikel " + oProd.ItemNo.ToString + " er opdateret med ny dato " + HuskStartSlutDato
    '                            'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
    '                            WriteInConsole(Msg)
    '                        End If


    '                        sql = "Update RDR1 set U_RCS_AFS = '" + oProd.DueDate.ToString("yyyyMMdd") + "'"
    '                        sql = sql + " Where VisOrder = " + VisOrder.ToString + " and DocEntry = " + salesOrderDocEntry.ToString
    '                        rs4.DoQuery(sql)

    '                        SBO_Application.SetStatusBarMessage(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " Prøver med forfaldsdato: " + oProd.DueDate.ToString("dd-MM-yyyy"))
    '                        PlanOk = False


    '                        ' Exit Do

    '                        'oProd.StartDate = oProd.Stages.StartDate
    '                    End If

    '                    '  If PlanUnlimeted = False Then
    '                    oProd.Stages.SetCurrentLine(0)
    '                    oProd.Stages.EndDate = HuskProdDateEndNy
    '                    sql = "Select max(isnull(LeadTime,0)) as PQty from OITM "
    '                    sql = sql + "Where ItemCode in (select wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) "


    '                    rs4.DoQuery(sql)
    '                    Dim LeadTime As Integer
    '                    LeadTime = rs4.Fields.Item("PQty").Value

    '                    'rs 20210502
    '                    oProd.Stages.StartDate = HuskProdDateEndNy.AddDays(-1 * LeadTime)
    '                    ' End If



    '                    Dim teststart

    '                    teststart = oProd.Stages.StartDate

    '                    If HuskProdDateStartFromStart <= oProd.Stages.StartDate Or PlanUnlimeted = True Then
    '                        '   oProd.SaveXML("c:\rcs\prod.xml")
    '                        iResult = oProd.Update()
    '                        If iResult < 0 Then
    '                            vCmp.GetLastError(iResult, Msg)
    '                            ErrorLog("CreateCopyOfBillOfMaterial", Msg)
    '                            'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
    '                            WriteInConsole(Msg)
    '                            If PlanUnlimeted Then

    '                                'SBO_Application.StatusBar.SetText(HuskStartSlutDatoUB, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
    '                                WriteInConsole(HuskStartSlutDatoUB)
    '                            End If
    '                        Else


    '                            '  String code = "(select MAX(cast(Code as bigint)) + 1 from \"@RCS_CAP_COMMIT\")";
    '                            '          String delete = "Delete from \"@RCS_CAP_COMMIT\" where U_DocEntry = {0} And U_LineNum = {1} and U_ObjType = {2} and U_ResCode = {3}";
    '                            '          String insert = "insert Into \"@RCS_CAP_COMMIT\"(Code, Name, U_ResCode, U_DocEntry, U_LineNum, U_Capacity, U_Date, U_ResGrpCod, U_ObjType) values({0}, {0}, '{1}', {2}, {3}, {4}, '{5}', {6}, {7})";

    '                            SetResourceCapacityCommitted(oProd.AbsoluteEntry)

    '                            Msg = "Produktionsordre for artikel " + oProd.ItemNo.ToString + " er opdateret "
    '                            'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
    '                            'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
    '                            WriteInConsole(Msg, True)
    '                        End If
    '                    Else
    '                        If PlanOk = False Then
    '                            Msg = "Produktionsordre for artikel " + oProd.ItemNo.ToString + " er ikke opdateret! "
    '                            'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
    '                            'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
    '                            WriteInConsole(Msg, True)
    '                        End If

    '                    End If

    '                End If
    '            End If
    '        Loop

    '        'oSO.ReleaseComObject() ' = Nothing
    '        ' oProd.ReleaseComObject() ' = Nothing
    '        'oRe.ReleaseComObject() ' = Nothing

    '    Catch ex As Exception
    '        ErrLog("UpdateProductionOrderLineStartEndDate", ex)
    '    Finally
    '        rs.ReleaseComObject
    '        rs2.ReleaseComObject
    '        rs3.ReleaseComObject
    '        rs4.ReleaseComObject
    '        oProd.ReleaseComObject
    '    End Try


    'End Sub

    Private Sub UpdateProductionOrderLineStartEndDateNew(ByVal ProdDocEntry As Integer, ByVal ProdDocNum As String, ByVal PlanUnlimeted As Boolean, salesOrderDocEntry As Long, VisOrder As Integer)

        RCS_PLAN.UpdateProductionOrderLineStartEndDate(ProdDocEntry, ProdDocNum, PlanUnlimeted, salesOrderDocEntry, VisOrder, True)
        Exit Sub

        Dim tes

        Dim test As Object = ""
        Dim MaxPlanDays As Integer = Settings.Default.MaxPlanDays

        Dim rs As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim rs2 As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim rs3 As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim rs4 As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim rs5 As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim rs6 As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)

        Dim oProd As ProductionOrders = vCmp.GetBusinessObject(BoObjectTypes.oProductionOrders)
        Dim sb As New Text.StringBuilder
        Dim OworDocEntry As Integer = 0
        Try

            Dim iResult As Integer = 0
            Dim Msg As String = ""

            Dim sql As String
            Dim PlanOk As Boolean = False
            Dim OworRCS_PT As String = "SER"


            Dim PieceProduction As String = Settings.Default.PieceProduction


            sql = "select distinct DocEntry from owor  "
            sql = sql + "Where DocNum = " + ProdDocNum + ""
            rs.DoQuery(sql)
            OworDocEntry = rs.Fields.Item("DocEntry").Value

            sql = " Update owor Set U_RCS_PT = 'SER'  where DocEntry in (Select DocEntry from wor1 where(rtrim(CAST(Docentry as char)) +'-' + CAST(StageId as char) + '-' + CAST(ItemCode as char))  in "
            sql = sql + "(select top 1 Case when sum(Wor1.PlannedQty+ Wor1.U_RCS_ESKM * owor.PlannedQty) > " + PieceProduction + "  then rtrim(CAST(Wor1.Docentry as char)) +'-' + CAST(StageId as char)  + '-' + CAST(Wor1.ItemCode as char) else '0-0' "
            sql = sql + "End from wor1 inner join Owor On owor.DocEntry = wor1.DocEntry inner join ORSC On ORSC.ResCode = wor1.ItemCode and ORSC.QryGroup3 = 'Y' where StageID = 3 and isnull(OWOR.U_RCS_PT,'') <> 'SER' and OWor.DocEntry = " + OworDocEntry.ToString
            sql = sql + " Group by wor1.Docentry, StageId, Wor1.ITemCode, wor1.LineNum Order by wor1.LineNum DESC)) "

            rs4.DoQuery(sql)


            sql = " Update owor Set U_RCS_PT = 'STP'  where DocEntry in (Select DocEntry from wor1 where(rtrim(CAST(Docentry as char)) +'-' + CAST(StageId as char)  + '-' + CAST(ItemCode as char))  in "
            sql = sql + "(select top 1 Case when sum(Wor1.PlannedQty+ Wor1.U_RCS_ESKM * owor.PlannedQty) <= " + PieceProduction + "  then rtrim(CAST(Wor1.Docentry as char)) +'-' + CAST(StageId as char) + '-' + CAST(Wor1.ItemCode as char) else '0-0' "
            sql = sql + "End from wor1 inner join Owor On owor.DocEntry = wor1.DocEntry inner join ORSC On ORSC.ResCode = wor1.ItemCode and ORSC.QryGroup3 = 'Y' where StageID = 3 and isnull(OWOR.U_RCS_PT,'') <> 'STP' and OWor.DocEntry = " + OworDocEntry.ToString
            sql = sql + " Group by wor1.Docentry, StageId, Wor1.ITemCode, wor1.LineNum Order by wor1.LineNum DESC )) "

            rs4.DoQuery(sql)


            sql = "select U_RCS_PT from owor  "
            sql = sql + "Where DocNum = " + ProdDocNum + ""

            rs4.DoQuery(sql)

            OworRCS_PT = rs4.Fields.Item("U_RCS_PT").Value


            Console.WriteLine("Produktionsordre er: " + OworRCS_PT)




            Do Until PlanOk = True
                PlanOk = True
                sql = "select distinct DocEntry from owor  "
                sql = sql + "Where DocNum = " + ProdDocNum + ""
                rs.DoQuery(sql)

                If rs.EoF Then
                    Exit Try
                End If

                sql = "Select sum(case when PlannedQty < 1 then 1 else PlannedQty end) as PlannedQtyTotal from WOR1 inner join orsc on wor1.itemcode = orsc.rescode "
                sql = sql + "Where (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and ItemType = 290 and DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
                rs2.DoQuery(sql)

                sql = "Delete from ""@RCS_CAP_COMMIT"" where U_DocEntry = " + rs.Fields.Item("DocEntry").Value.ToString + ""
                rs3.DoQuery(sql)
                Dim PlannedQtyTotalHour As Double = rs2.Fields.Item("PlannedQtyTotal").Value / 60

                Dim CntDay As Double = 0
                Dim Delay As Double = 0
                Dim Day As String = ""
                Dim WeekendDays As Integer = 0

                sql = "Select sum(1) as DelayTotal from WOR1 inner join orsc on wor1.itemcode = orsc.rescode "
                sql = sql + "Where (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and  wor1.ItemType = 290 and wor1.DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
                rs2.DoQuery(sql)
                Dim DelayTotal As Double = DelayTotal + rs2.Fields.Item("DelayTotal").Value

                sql = "Select count(DocEntry) as ResTotal from WOR1  inner join orsc on wor1.itemcode = orsc.rescode "
                sql = sql + "Where (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and ItemType = 290 and DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
                rs2.DoQuery(sql)
                Dim ResTotal As Double = rs2.Fields.Item("ResTotal").Value

                sql = "Select  ORSC.ResGrpCod, wor1.StageId , wor4.SeqNum-1 As SeqNum ,sum(WOR1.PlannedQty+OWOR.PlannedQty*isnull(U_RCS_ESKM,0)-WOR1.IssuedQty) as PQty, sum(isnull(WOR1.U_RCS_Issued,0)) U_RCS_Issued, sum(isnull(U_RCS_ED,0)) as Nd, orsc.QryGroup7, orsc.QryGroup14, orsc.QryGroup6,  max(Case When orsc.QryGroup8 = 'Y' then 1 else 0 end) as QryGroup8, isnull(U_RCS_PU,'') as U_RCS_PU,  isnull(U_RCS_DEL,'') as U_RCS_DEL  from WOR1 "
                sql = sql + "inner Join OWOR on OWOR.DocEntry =  wor1.docentry inner Join wor4 on wor4.stageid = wor1.stageid And wor1.docentry = wor4.docentry inner join orsc on wor1.itemcode = orsc.rescode left join ocrd on SubString(wor1.itemcode,3,2) = OCRD.U_RCS_CONO "
                sql = sql + "Where  (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and  wor1.ItemType = 290 And wor1.DocEntry  =" + rs.Fields.Item("DocEntry").Value.ToString + " "
                sql = sql + "Group by SeqNum , wor1.StageId, orsc.QryGroup7, orsc.QryGroup14, orsc.QryGroup6,  isnull(U_RCS_PU,''),  isnull(U_RCS_DEL,''),  ORSC.ResGrpCod "
                sql = sql + "order by SeqNum "
                rs2.DoQuery(sql)

                Dim DateCntPlanned As Double
                Dim DateCntPlannedPerDate As Double
                Dim PlannedQty As Double

                Dim HuskProdDateStartFromStart As Date
                Dim HuskProdDateEndLatest As Date

                If oProd.GetByKey(rs.Fields.Item("DocEntry").Value) Then

                    HuskProdDateStartFromStart = oProd.StartDate
                    HuskProdDateStartNy = oProd.StartDate
                    HuskProdDateEndNy = oProd.DueDate
                    HuskProdDateEndLatest = oProd.DueDate

                    Dim Cnt As Integer = rs2.RecordCount - 1
                    Dim CntRecHusk As Integer = rs2.RecordCount - 1
                    Dim HuskResGrpcode As Integer = 0
                    Dim SameDate As Boolean = False
                    Dim HuskQtySameDate As Double = 0
                    Dim HuskSameDateOneShot As Boolean = False
                    Dim HPD As Double = 8
                    Dim STP As Double = 0
                    Dim SEP As Double = 0
                    Dim HuskStartSlutDato As String = ""
                    Dim HuskStartSlutDatoUB As String = ""

                    rs2.MoveLast()
                    Do Until Cnt < 0
                        Dim PQty As Double = rs2.Fields.Item("PQty").Value
                        Dim temp As String = rs2.Fields.Item("StageId").Value

                        If HuskResGrpcode = rs2.Fields.Item("ResGrpCod").Value Then

                            If oProd.Stages.EndDate = oProd.Stages.StartDate Then
                                SameDate = True
                                HuskSameDateOneShot = True
                            Else
                                SameDate = False
                            End If
                        Else
                            SameDate = False

                            Dim testDate As Date = oProd.Stages.StartDate
                            If HuskProdDateEndNy = oProd.Stages.StartDate Then
                                HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
                                If Day = "Sat" Then
                                    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                End If
                                If Day = "Sun" Then
                                    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
                                End If
                            End If

                        End If

                        Dim CntHusk As Integer = rs2.Fields.Item("Nd").Value
                        If rs2.Fields.Item("Nd").Value > 0 And 1 = 2 Then
                            WeekendDays = 0
                            CntHusk = rs2.Fields.Item("Nd").Value
                            CntDay = CntHusk
                            Dim CntDay2a As Double = rs2.Fields.Item("Nd").Value
                            Dim CntDayHuska As Double = rs2.Fields.Item("Nd").Value
                            If CntDay > 1 Then
                                Do Until (CntDay + WeekendDays <= 0)

                                    Day = HuskProdDateEndNy.AddDays(CntDayHuska - CntDay2a - 1).DayOfWeek.ToString().Substring(0, 3)
                                    If Day = "Sat" Or Day = "Sun" Then

                                        WeekendDays = WeekendDays + 1

                                    End If
                                    CntDay = CntDay - 1
                                    CntDay2a = CntDay2a + 1
                                Loop
                            End If

                            HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-(CntHusk) - WeekendDays)
                        End If

                        Dim stdate = oProd.Stages.StartDate
                        Dim endate = oProd.Stages.EndDate

                        Dim testseqNum = rs2.Fields.Item("SeqNum").Value
                        oProd.Stages.SetCurrentLine(rs2.Fields.Item("SeqNum").Value)

                        If HuskResGrpcode = rs2.Fields.Item("ResGrpCod").Value Then
                            sql = "Select isnull(U_RCS_NQD,'N') as NQD from ORSB "
                            sql = sql + "Where ResGrpCod = " + HuskResGrpcode.ToString()

                            rs4.DoQuery(sql)
                            Dim NQD As String
                            NQD = rs4.Fields.Item("NQD").Value

                            If NQD = "N" Then 'Ikke samme dato
                                If SameDate = True Then
                                    HuskProdDateEndNy = HuskProdDateEndLatest
                                End If
                            End If

                        Else
                            HuskQtySameDate = 0
                            sql = "Select isnull(U_RCS_NQD,'N') as NQD from ORSB "
                            sql = sql + "Where ResGrpCod = " + HuskResGrpcode.ToString()

                            rs4.DoQuery(sql)
                            Dim NQD As String
                            NQD = rs4.Fields.Item("NQD").Value

                            If NQD = "Y" Then 'Leverandør
                                HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)

                            End If

                            HuskResGrpcode = rs2.Fields.Item("ResGrpCod").Value
                        End If

                        sql = "Select isnull(U_RCS_HD,8) as HPD, isnull(U_RCS_STP,8) as STP, isnull(U_RCS_SEP,8) as SEP from ORSB "
                        sql = sql + "Where ResGrpCod = " + HuskResGrpcode.ToString()

                        rs4.DoQuery(sql)

                        HPD = rs4.Fields.Item("HPD").Value
                        If HPD < 1 Then ' BR
                            HPD = 8
                        End If

                        STP = rs4.Fields.Item("STP").Value

                        SEP = rs4.Fields.Item("SEP").Value

                        If OworRCS_PT = "SER" Then
                            HPD = SEP
                        Else
                            HPD = STP
                        End If



                        PlannedQty = rs2.Fields.Item("PQty").Value / 60
                        If PlannedQty < 1 Then
                            PlannedQty = 1
                        End If

                        If rs2.Fields.Item("QryGroup7").Value = "Y" Then
                            DateCntPlannedPerDate = (PlannedQty / 8) '+ 1 ' Antal dage
                        ElseIf rs2.Fields.Item("QryGroup6").Value = "Y" Then
                            DateCntPlannedPerDate = (PlannedQty / 8) + 1 ' Antal dage
                        Else
                            DateCntPlannedPerDate = PlannedQty / HPD ' Antal dage
                        End If

                        DateCntPlannedPerDate = Math.Ceiling(DateCntPlannedPerDate)
                        DateCntPlannedPerDate = DateCntPlannedPerDate + rs2.Fields.Item("Nd").Value
                        CntDay = DateCntPlannedPerDate
                        DateCntPlanned = DateCntPlannedPerDate
                        DateCntPlanned = Math.Ceiling(DateCntPlanned)

                        If rs2.Fields.Item("Nd").Value = 0 Then
                            If DateCntPlanned <= 1 Then
                                DateCntPlanned = 1
                            End If
                        Else
                            If DateCntPlanned <= 1 Then
                                DateCntPlanned = 1
                            End If
                        End If

                        Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)

                        If Day = "Sat" Then
                            HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                        End If
                        If Day = "Sun" Then
                            HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
                        End If

                        ' check for delevery
                        If rs2.Fields.Item("QryGroup6").Value = "Y" Then

                            Day = HuskProdDateEndNy.DayOfWeek.ToString()

                            Dim startdayvalue As String = "0"
                            Select Case Day.ToLower
                                Case "monday"
                                    startdayvalue = "1"
                                Case "tuesday"
                                    startdayvalue = "2"
                                Case "wednesday"
                                    startdayvalue = "3"
                                Case "thursday"
                                    startdayvalue = "4"
                                Case "friday"
                                    startdayvalue = "5"
                                Case "saturday"
                                    startdayvalue = "6"
                                Case "sunday"
                                    startdayvalue = "7"
                            End Select

                            Dim DayNum As String = startdayvalue
                            Dim DelDays As String = rs2.Fields.Item("U_RCS_DEL").Value

                            Dim CntDel As Integer = 0
                            Dim status As Boolean = False
                            Do Until CntDel > 1

                                Do Until Int(DayNum) <= 0
                                    If InStr(DelDays, Trim(DayNum)) > 0 Then
                                        ' all good
                                        status = True
                                        Exit Do
                                    Else
                                        DayNum = Str(Int(DayNum) - 1)
                                        HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                        Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
                                        If Day = "Sat" Then
                                            HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                        End If
                                        If Day = "Sun" Then
                                            HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
                                        End If
                                    End If

                                Loop
                                If status = True Then
                                    Exit Do
                                End If
                                HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
                                If Day = "Sat" Then
                                    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                End If
                                If Day = "Sun" Then
                                    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
                                End If
                                CntDel = CntDel + 1
                                DayNum = "5"
                            Loop

                            If Cnt < CntRecHusk Then
                                oProd.Stages.SetCurrentLine(rs2.Fields.Item("SeqNum").Value + 1)
                                oProd.Stages.StartDate = HuskProdDateEndNy
                                oProd.Stages.SetCurrentLine(rs2.Fields.Item("SeqNum").Value)
                            End If

                            If rs2.Fields.Item("QryGroup8").Value = "1" Then

                                oProd.Stages.StartDate = HuskProdDateEndNy.AddDays(-7).AddDays(CntHusk)
                                HuskProdDateEndNy = oProd.Stages.StartDate
                                DateCntPlanned = 1 + 5
                            End If

                        End If

                        CntDay = DateCntPlanned

                        Dim testdate4 As Date = oProd.Stages.EndDate
                        oProd.Stages.EndDate = HuskProdDateEndNy

                        HuskProdDateEndLatest = HuskProdDateEndNy
                        WeekendDays = 0
                        Dim CntDay2 As Double = CntDay
                        Dim CntDayHusk As Double = CntDay
                        If CntDay > 1 Then
                            Do Until ((CntDay - 1 + WeekendDays) <= 0)

                                Day = HuskProdDateEndNy.AddDays(CntDayHusk - CntDay2 - 1).DayOfWeek.ToString().Substring(0, 3)
                                If Day = "Sat" Or Day = "Sun" Then
                                    WeekendDays = WeekendDays + 1
                                End If
                                CntDay = CntDay - 1
                                CntDay2 = CntDay2 + 1

                                'you have to check day equal to "sat" Or "sun".
                            Loop
                        End If


                        'Day = oProd.Stages.EndDate.DayOfWeek.ToString().Substring(0, 3)
                        Day = oProd.Stages.EndDate.AddDays(-(DateCntPlanned + WeekendDays - 1)).DayOfWeek.ToString().Substring(0, 3)

                        If Day = "Sat" Then
                            oProd.Stages.StartDate = oProd.Stages.EndDate.AddDays(-(DateCntPlanned + WeekendDays - 1) - 1)
                            HuskProdDateEndNy = oProd.Stages.StartDate
                        End If
                        If Day = "Sun" Then
                            oProd.Stages.StartDate = oProd.Stages.EndDate.AddDays(-(DateCntPlanned + WeekendDays - 1) - 2)
                            HuskProdDateEndNy = oProd.Stages.StartDate
                        End If
                        If Day <> "Sun" And Day <> "Sat" Then
                            oProd.Stages.StartDate = oProd.Stages.EndDate.AddDays(-(DateCntPlanned + WeekendDays - 1))
                            HuskProdDateEndNy = oProd.Stages.StartDate
                        End If
                        ' check for for pickup
                        If rs2.Fields.Item("QryGroup6").Value = "Y" And rs2.Fields.Item("QryGroup8").Value = "1" Then
                            Dim DayNum As String

                            Day = HuskProdDateEndNy.DayOfWeek.ToString()

                            Dim startdayvalue As String = "0"
                            Select Case Day.ToLower
                                Case "monday"
                                    startdayvalue = "1"
                                Case "tuesday"
                                    startdayvalue = "2"
                                Case "wednesday"
                                    startdayvalue = "3"
                                Case "thursday"
                                    startdayvalue = "4"
                                Case "friday"
                                    startdayvalue = "5"
                                Case "saturday"
                                    startdayvalue = "6"
                                Case "sunday"
                                    startdayvalue = "7"
                            End Select

                            DayNum = startdayvalue

                            Dim DelDays As String = rs2.Fields.Item("U_RCS_PU").Value


                            Dim CntDel As Integer = 0
                            Dim status As Boolean = False
                            Do Until CntDel > 1

                                Do Until Int(DayNum) <= 0
                                    If InStr(DelDays, Trim(DayNum)) > 0 Then
                                        ' all good
                                        status = True
                                        Exit Do
                                    Else
                                        DayNum = Str(Int(DayNum) - 1)
                                        HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                        Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
                                        If Day = "Sat" Then
                                            HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                        End If
                                        If Day = "Sun" Then
                                            HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
                                        End If
                                    End If

                                Loop
                                If status = True Then
                                    Exit Do
                                End If

                                CntDel = CntDel + 1
                                DayNum = "5"
                            Loop

                            oProd.Stages.StartDate = HuskProdDateEndNy
                        End If

                        Dim HuskNewEndDate As Date
                        If rs2.Fields.Item("QryGroup7").Value = "Y" Or 1 = 1 Then
                            HuskNewEndDate = oProd.Stages.EndDate
                            If oProd.Stages.StartDate >= oProd.Stages.EndDate Then
                                HuskNewEndDate = oProd.Stages.StartDate
                            End If

                        Else
                            HuskNewEndDate = oProd.Stages.EndDate.AddDays(-1)
                            If oProd.Stages.StartDate >= oProd.Stages.EndDate.AddDays(-1) Then
                                HuskNewEndDate = oProd.Stages.StartDate
                            End If

                        End If

                        sql = "Select SeqNum-1 as pos From wor4 Where DocEntry = " + oProd.AbsoluteEntry.ToString + " and StageID = " + oProd.Stages.StageID.ToString
                        rs.DoQuery(sql)
                        Dim pos As String = rs.Fields.Item("pos").Value

                        sql = "select sum(ORCJ.Capacity) as Cap From ORCJ "
                        sql = sql + "Join ORSC On ORSC.ResCode = ORCJ.ResCode "
                        sql = sql + "Where ORSC.ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                        sql = sql + "And  ORCJ.CapDate BETWEEN '" + oProd.Stages.StartDate.ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' "
                        sql = sql + "And ORCJ.CapType = 'I' "
                        sql = sql + " And ORCJ.WhsCode = '01' "
                        rs.DoQuery(sql)

                        sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry "
                        sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                        sql = sql + "And  U_Date BETWEEN '" + oProd.Stages.StartDate.ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"
                        rs4.DoQuery(sql)

                        Dim test1, test2, test3, testdate1, testdate2
                        test1 = rs2.Fields.Item("PQty").Value
                        test2 = rs.Fields.Item("Cap").Value
                        test3 = rs4.Fields.Item("ComCap").Value

                        testdate1 = oProd.Stages.StartDate.ToString("MM-dd-yyyy")
                        testdate2 = HuskNewEndDate.ToString("MM-dd-yyyy")

                        Dim HuskPosEndDate As String = HuskNewEndDate.ToString("dd-MM-yyyy")
                        HuskStartSlutDatoUB = HuskStartSlutDatoUB + " Pos: " + pos.ToString + " Dato: " + oProd.Stages.StartDate.ToString("dd-MM-yyyy") + ":" + HuskNewEndDate.ToString("dd-MM-yyyy")

                        If 1 = 1 Then

                            Dim CntDay3 As Integer = 0
                            Dim Cnt2 As Integer = 0
                            Dim sql1 As String = ""
                            Dim Cnt4 As Integer = 0
                            Dim BookState As Boolean = False
                            Dim CanStartBook As Boolean = False
                            Dim CapPerDayTotal As Integer = 0
                            Dim CapPerDayTotalFirstRun As Integer = 0

                            Dim firtsRun As Boolean = True

                            sql = "select sum(ORCJ.Capacity) as Cap From ORCJ "
                            sql = sql + "Join ORSC On ORSC.ResCode = ORCJ.ResCode "
                            sql = sql + "Where ORSC.ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                            sql = sql + "And  ORCJ.CapDate BETWEEN '" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' "
                            sql = sql + "And ORCJ.CapType = 'I' "
                            sql = sql + " And ORCJ.WhsCode = '01' "
                            rs5.DoQuery(sql)
                            test2 = rs5.Fields.Item("Cap").Value

                            sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
                            sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                            sql = sql + "And  U_Date BETWEEN '" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"

                            rs6.DoQuery(sql)


                            test3 = rs6.Fields.Item("ComCap").Value

                            Do Until (rs2.Fields.Item("PQty").Value <= CapPerDayTotal)

                                sql = "select sum(ORCJ.Capacity) as Cap From ORCJ "
                                sql = sql + "Join ORSC On ORSC.ResCode = ORCJ.ResCode "
                                sql = sql + "Where ORSC.ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                                sql = sql + "And  ORCJ.CapDate BETWEEN '" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' "
                                sql = sql + "And ORCJ.CapType = 'I' "
                                sql = sql + " And ORCJ.WhsCode = '01' "
                                rs5.DoQuery(sql)
                                test2 = rs5.Fields.Item("Cap").Value

                                sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
                                sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                                sql = sql + "And  U_Date BETWEEN '" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"
                                rs6.DoQuery(sql)
                                test3 = rs6.Fields.Item("ComCap").Value

                                sql1 = "select sum(ORCJ.Capacity) as Cap From ORCJ "
                                sql1 = sql1 + "Join ORSC On ORSC.ResCode = ORCJ.ResCode "
                                sql1 = sql1 + "Where ORSC.ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                                sql1 = sql1 + "And  ORCJ.CapDate BETWEEN '" + HuskNewEndDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' "
                                sql1 = sql1 + "And ORCJ.CapType = 'I' "
                                sql1 = sql1 + " And ORCJ.WhsCode = '01' "
                                rs.DoQuery(sql1)
                                test2 = rs.Fields.Item("Cap").Value

                                sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
                                sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                                sql = sql + "And  U_Date BETWEEN '" + HuskNewEndDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"
                                rs4.DoQuery(sql)

                                test1 = rs2.Fields.Item("PQty").Value
                                test2 = rs.Fields.Item("Cap").Value
                                test3 = rs4.Fields.Item("ComCap").Value

                                'if capacity > comitted capacity
                                If (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value) > 0 Then
                                    ' if free capacity > max capacity per day
                                    If (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value) > (HPD * 60) Then
                                        CapPerDayTotal = CapPerDayTotal + (HPD * 60)
                                    Else
                                        If PlanUnlimeted = True And rs.Fields.Item("Cap").Value > 0 Then
                                            CapPerDayTotal = CapPerDayTotal + (HPD * 60)
                                        Else
                                            CapPerDayTotal = CapPerDayTotal + (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value)
                                        End If
                                    End If
                                    '  CntDay3 = CntDay3 - Math.Ceiling((rs2.Fields.Item("PQty").Value - (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value)) / (HPD * 60))
                                    CntDay3 = CntDay3 - 1
                                Else
                                    If PlanUnlimeted = True And rs.Fields.Item("Cap").Value > 0 Then
                                        CapPerDayTotal = CapPerDayTotal + (HPD * 60)
                                    End If
                                    CntDay3 = CntDay3 - 1
                                End If

                                HuskProdDateEndNy = HuskNewEndDate.AddDays(CntDay3 + 1) ' oProd.Stages.StartDate.AddDays(CntDay3)
                                Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
                                If Day = "Sat" Then
                                    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                End If
                                If Day = "Sun" Then
                                    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
                                End If

                                ' check for for pickup
                                If rs2.Fields.Item("QryGroup6").Value = "Y" And rs2.Fields.Item("QryGroup8").Value = "1" Then

                                    Day = HuskProdDateEndNy.DayOfWeek.ToString()

                                    Dim startdayvalue As String = "0"
                                    Select Case Day.ToLower
                                        Case "monday"
                                            startdayvalue = "1"
                                        Case "tuesday"
                                            startdayvalue = "2"
                                        Case "wednesday"
                                            startdayvalue = "3"
                                        Case "thursday"
                                            startdayvalue = "4"
                                        Case "friday"
                                            startdayvalue = "5"
                                        Case "saturday"
                                            startdayvalue = "6"
                                        Case "sunday"
                                            startdayvalue = "7"
                                    End Select


                                    Dim DayNum As String = startdayvalue
                                    Dim DelDays As String = rs2.Fields.Item("U_RCS_PU").Value

                                    Dim CntDel As Integer = 0
                                    Dim status As Boolean = False
                                    Do Until CntDel > 1

                                        Do Until Int(DayNum) <= 0
                                            If InStr(DelDays, Trim(DayNum)) > 0 Then
                                                ' all good
                                                status = True
                                                Exit Do
                                            Else
                                                DayNum = Str(Int(DayNum) - 1)
                                                HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                                Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
                                                If Day = "Sat" Then
                                                    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                                End If
                                                If Day = "Sun" Then
                                                    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
                                                End If
                                            End If

                                        Loop
                                        If status = True Then
                                            Exit Do
                                        End If
                                        CntDel = CntDel + 1
                                        DayNum = "5"
                                    Loop

                                End If




                                'Check for passing startdate
                                If (DateDiff(DateInterval.Day, Date.Now, HuskProdDateEndNy) < 0) Or (DateDiff(DateInterval.Day, oProd.StartDate, HuskProdDateEndNy) < 0) Then
                                    ' Then start from begin with new end date moved 1 day
                                    Dim newDueDate As Date = oProd.DueDate.AddDays(1)
                                    Day = newDueDate.DayOfWeek.ToString().Substring(0, 3)
                                    If Day = "Sat" Then
                                        newDueDate = newDueDate.AddDays(2)
                                    End If
                                    If Day = "Sun" Then
                                        newDueDate = newDueDate.AddDays(1)
                                    End If

                                    oProd.StartDate = Date.Now
                                    oProd.DueDate = newDueDate
                                    Dim cnt3 As Integer = 0

                                    Do Until cnt3 > oProd.Stages.Count - 1
                                        oProd.Stages.SetCurrentLine(cnt3)
                                        oProd.Stages.StartDate = Date.Now
                                        oProd.Stages.EndDate = newDueDate

                                        cnt3 = cnt3 + 1
                                    Loop

                                    iResult = oProd.Update()
                                    If iResult < 0 Then
                                        vCmp.GetLastError(iResult, Msg)
                                        ErrorLog("CreateCopyOfBillOfMaterial", Msg)
                                        'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                        WriteInConsole("Produktionsordre for artikel " + oProd.ItemNo.ToString + " fejlet: " + Msg)
                                    Else
                                        Msg = "Produktionsordre for artikel " + oProd.ItemNo.ToString + " er opdateret med ny dato. Stopper planlæggning ved pos: " + pos.ToString + " med slutdato: " + HuskNewEndDate.ToString("dd-MM-yyyy")
                                        'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                        WriteInConsole(Msg)
                                    End If

                                    sql = "Update RDR1 set U_RCS_AFS = '" + oProd.DueDate.ToString("yyyyMMdd") + "'"
                                    sql = sql + " Where VisOrder = " + VisOrder.ToString + " and DocEntry = " + salesOrderDocEntry.ToString
                                    rs4.DoQuery(sql)

                                    'SBO_Application.StatusBar.SetText("Planlægning fejler. Artikelnr.: " + oProd.ItemNo.ToString, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                    PlanOk = False

                                    Exit Do
                                End If

                                'Check for passing startdate
                                If (DateDiff(DateInterval.Day, Date.Now, HuskProdDateEndNy) < 0) Or (DateDiff(DateInterval.Day, oProd.StartDate, HuskProdDateEndNy) < 0) Then

                                    'SBO_Application.SetStatusBarMessage(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " Har passeret startdato stopper planlægning! " + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' mellem '" + HuskNewEndDate.ToString("MM-dd-yyyy"))
                                    'Console.WriteLine(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " Har passeret startdato stopper planlægning! " + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' mellem '" + HuskNewEndDate.ToString("MM-dd-yyyy"))
                                    Cnt = 0
                                    Exit Do
                                End If
                                'Check for max 60 days
                                If Cnt2 > MaxPlanDays Then
                                    'SBO_Application.SetStatusBarMessage(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + "Har passeret 30 dage stopper planlægning!" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' mellem '" + HuskNewEndDate.ToString("MM-dd-yyyy"))
                                    'Console.WriteLine(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + "Har passeret 30 dage stopper planlægning!" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' mellem '" + HuskNewEndDate.ToString("MM-dd-yyyy"))

                                    Dim Tekst = "Skal bruge " + test1.ToString + " Der er kapacitet ialt " + test2.ToString + " Der er brugt " + test3.ToString

                                    sql = "Select [@RCS_CAP_COMMIT].U_Date ,owor.docNum, round(Sum(U_Capacity)/60,2) as comcap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry "
                                    sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                                    sql = sql + "And  U_Date BETWEEN '" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString
                                    sql = sql + " group by [@RCS_CAP_COMMIT].U_Date,owor.DocNum "
                                    sql = sql + " order by U_Date"
                                    rs4.DoQuery(sql)

                                    test1 = rs4.Fields.Item("U_Date").Value
                                    test2 = rs4.Fields.Item("docNum").Value
                                    test3 = rs4.Fields.Item("ComCap").Value

                                    Do Until rs4.EoF
                                        test1 = rs4.Fields.Item("U_Date").Value
                                        test2 = rs4.Fields.Item("docNum").Value
                                        test3 = rs4.Fields.Item("ComCap").Value
                                        Tekst = Tekst + vbCrLf + test1.ToString + " " + test2.ToString + " " + test3.ToString
                                        rs4.MoveNext()
                                    Loop
                                    'SBO_Application.MessageBox(Tekst)
                                    WriteInConsole(Tekst)
                                    Cnt = 0
                                    Exit Do
                                End If
                                Cnt2 = Cnt2 + 1
                            Loop

                            oProd.Stages.StartDate = HuskProdDateEndNy
                            HuskStartSlutDato = HuskStartSlutDato + " Pos: " + pos.ToString + " Dato: " + HuskProdDateEndNy.ToString("dd-MM-yyyy") + ":" + HuskPosEndDate

                        End If
                        HuskQtySameDate = HuskQtySameDate + rs2.Fields.Item("PQty").Value

                        Dim HuskProdDateEndNy2 As Date
                        HuskProdDateEndNy2 = HuskProdDateEndNy
                        HuskProdDateEndNy2 = HuskProdDateEndNy2.AddDays(1)
                        Day = HuskProdDateEndNy2.DayOfWeek.ToString().Substring(0, 3)
                        If Day = "Sat" Then
                            HuskProdDateEndNy2 = HuskProdDateEndNy2.AddDays(-1)
                        End If
                        If Day = "Sun" Then
                            HuskProdDateEndNy2 = HuskProdDateEndNy2.AddDays(-2)
                        End If

                        oProd.Stages.EndDate = HuskNewEndDate

                        Cnt = Cnt - 1
                        If Cnt >= 0 Then
                            rs2.MovePrevious()

                        End If

                        testdate1 = oProd.Stages.StartDate.ToString("MM-dd-yyyy")
                        testdate2 = HuskNewEndDate.ToString("MM-dd-yyyy")

                        If PlanOk = False Then
                            Exit Do
                        End If
                    Loop


                    If PlanOk = True Then

                        Dim sd As Date = oProd.StartDate
                        Dim ssd As Date = oProd.Stages.StartDate
                        Dim test1a As Date = oProd.StartDate
                        Dim test2a As Date = oProd.Stages.StartDate

                        If oProd.StartDate > oProd.Stages.StartDate And PlanUnlimeted = False Then

                            Dim Missday

                            Missday = DateDiff(DateInterval.Day, oProd.StartDate, oProd.Stages.StartDate)

                            Msg = "Produktionsordre " + oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " er ikke opdateret! er forbi start dato!  "

                            ' Then start from begin with new end date moved 1 day
                            Dim newDueDate As Date = oProd.DueDate.AddDays(1)

                            Day = newDueDate.DayOfWeek.ToString().Substring(0, 3)
                            If Day = "Sat" Then
                                newDueDate = newDueDate.AddDays(2)
                            End If
                            If Day = "Sun" Then
                                newDueDate = newDueDate.AddDays(1)
                            End If

                            oProd.StartDate = Date.Now
                            oProd.DueDate = newDueDate
                            Dim cnt3 As Integer = 0

                            Do Until cnt3 > oProd.Stages.Count - 1
                                oProd.Stages.SetCurrentLine(cnt3)
                                oProd.Stages.StartDate = Date.Now
                                oProd.Stages.EndDate = newDueDate

                                cnt3 = cnt3 + 1
                            Loop

                            iResult = oProd.Update()
                            If iResult < 0 Then
                                vCmp.GetLastError(iResult, Msg)
                                ErrorLog("CreateCopyOfBillOfMaterial", Msg)
                                'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                WriteInConsole(Msg)
                            Else
                                Msg = "Produktionsordre for artikel " + oProd.ItemNo.ToString + " er opdateret med ny dato " + HuskStartSlutDato
                                'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                WriteInConsole(Msg)
                            End If

                            sql = "Update RDR1 set U_RCS_AFS = '" + oProd.DueDate.ToString("yyyyMMdd") + "'"
                            sql = sql + " Where VisOrder = " + VisOrder.ToString + " and DocEntry = " + salesOrderDocEntry.ToString
                            rs4.DoQuery(sql)

                            'SBO_Application.SetStatusBarMessage(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " Prøver med forfaldsdato: " + oProd.DueDate.ToString("dd-MM-yyyy"))
                            WriteInConsole(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " Prøver med forfaldsdato: " + oProd.DueDate.ToString("dd-MM-yyyy"))
                            PlanOk = False
                        End If

                        oProd.Stages.SetCurrentLine(0)
                        oProd.Stages.EndDate = HuskProdDateEndNy
                        sql = "Select max(isnull(LeadTime,0)) as PQty from OITM "
                        sql = sql + "Where ItemCode in (select wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) "

                        rs4.DoQuery(sql)
                        Dim LeadTime As Integer = rs4.Fields.Item("PQty").Value

                        oProd.Stages.StartDate = HuskProdDateEndNy.AddDays(-1 * LeadTime)

                        Dim teststart
                        teststart = oProd.Stages.StartDate

                        If HuskProdDateStartFromStart <= oProd.Stages.StartDate Or PlanUnlimeted = True Then
                            iResult = oProd.Update()
                            If iResult < 0 Then
                                vCmp.GetLastError(iResult, Msg)
                                ErrorLog("CreateCopyOfBillOfMaterial", Msg)
                                'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                WriteInConsole(Msg)
                                If PlanUnlimeted Then

                                    'SBO_Application.StatusBar.SetText(HuskStartSlutDatoUB, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                    WriteInConsole(HuskStartSlutDatoUB)
                                End If
                            Else

                                SetResourceCapacityCommitted(oProd.AbsoluteEntry)

                                Msg = "Produktionsordre for artikel " + oProd.ItemNo.ToString + " er opdateret "
                                'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                                WriteInConsole(Msg, True)
                            End If
                        Else
                            If PlanOk = False Then
                                Msg = "Produktionsordre for artikel " + oProd.ItemNo.ToString + " er ikke opdateret! "
                                'SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                WriteInConsole(Msg, True)
                            End If

                        End If

                    End If
                End If
            Loop

        Catch ex As Exception
            ErrLog("UpdateProductionOrderLineStartEndDate", ex)
        Finally
            rs.ReleaseComObject
            rs2.ReleaseComObject
            rs3.ReleaseComObject
            rs4.ReleaseComObject
            rs5.ReleaseComObject
            rs6.ReleaseComObject
            oProd.ReleaseComObject
        End Try


    End Sub

    Private Sub SetResourceCapacityCommitted(ByVal DocEntry As Integer)

        Dim rs As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim rs4 As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim sql As String = ""

        Try
            Dim sb As New Text.StringBuilder()
            sb.Append("SELECT").AppendLine()
            sb.Append("OWOR.DocEntry, OWOR.ObjType,  WOR1.LineNum").AppendLine()
            sb.Append(", ((WOR1.PlannedQty + isnull(WOR1.U_RCS_EOU, 0)+OWOR.PlannedQty*isnull(WOR1.U_RCS_ESKM,0)) -(isnull(U_Rcs_MATOK,0)/(OWOR.PlannedQty/WOR1.PlannedQty))) PlannedQty, WOR1.IssuedQty, WOR1.StartDate, DateAdd(day,0,WOR1.EndDate) as EndDate").AppendLine()
            sb.Append(", isnull(sum(U_AIG_ARBT),0)U_AIG_ARBT ").AppendLine()
            sb.Append(", ORSC.VisResCode").AppendLine()
            sb.Append(", ORSC.ResGrpCod").AppendLine()
            sb.Append("from OWOR ").AppendLine()
            sb.Append("inner join WOR1 on OWOR.DocEntry = WOR1.DocEntry").AppendLine()
            sb.Append("inner join ORSC on ORSC.VisResCode =  WOR1.ItemCode").AppendLine()
            sb.Append("left join ""@AIG_OGRD"" O on O.U_AIG_PROD = OWOR.DocNum and O.U_AIG_POS = ORSC.VisResCode  ").AppendLine()
            sb.Append("and U_AIG_TYPE = 'JobTid' and U_AIG_SLET = 'N' ").AppendLine()
            sb.Append("where  WOR1.ItemType = 290 ").AppendLine()
            sb.Append("and OWOR.DocEntry = ").Append(DocEntry).AppendLine()
            sb.Append("and (OWOR.status='R' or OWOR.status='P')").AppendLine()
            sb.Append("and (WOR1.EndDate >= GETDATE() or isnull(WOR1.U_RCS_Done, 'N') = 'N')").AppendLine()
            sb.Append("Group by OWOR.Docnum, WOR1.itemcode, WOR1.EndDate,  ORSC.resname, OWOR.originnum").AppendLine()
            sb.Append(", WOR1.plannedQty, IssuedQty, WOR1.LineNum, WOR1.VisOrder,  OWOR.DocEntry,  WOR1.U_RCS_Done, ORSC.VisResCode,  WOR1.U_RCS_DoneDate,  WOR1.StartDate").AppendLine()
            sb.Append(", ORSC.ResGrpCod,  WOR1.U_RCS_EOU, OWOR.ObjType, WOR1.U_RCS_ESKM, OWOR.PlannedQty, U_RCS_MATOK  ").AppendLine()
            sb.Append("order by OWOR.DocEntry, WOR1.LineNum")
            rs.DoQuery(sb.ToString())

            Dim produktionOrderLines As New Generic.List(Of ProduktionOrderLine)
            Dim pl As ProduktionOrderLine
            pl = Nothing
            Do While Not rs.EoF

                pl = New ProduktionOrderLine()
                pl.DocEntry = rs.ValueAsInteger("DocEntry")
                pl.ObjType = rs.ValueAsInteger("ObjType")
                pl.LineNum = rs.ValueAsInteger("LineNum")
                pl.PlannedQty = rs.ValueAsDouble("PlannedQty")
                pl.IssuedQty = rs.ValueAsDouble("IssuedQty")
                pl.StartDate = rs.ValueAsDate("StartDate")
                pl.U_AIG_ARBT = rs.ValueAsDouble("U_AIG_ARBT")
                If (pl.StartDate < DateTime.Now) Then
                    pl.StartDate = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day)
                End If

                pl.EndDate = rs.ValueAsDate("EndDate")
                If (pl.EndDate < DateTime.Now) Then
                    pl.EndDate = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day)
                End If

                pl.VisResCode = rs.ValueAsString("VisResCode")
                pl.ResGrpCod = rs.ValueAsInteger("ResGrpCod")
                pl.RestQty = pl.PlannedQty - pl.IssuedQty

                produktionOrderLines.Add(pl)

                rs.MoveNext()
            Loop

            GetWorkingDays(produktionOrderLines, rs)
            SetCapacityCommitted(produktionOrderLines, rs)


        Catch ex As Exception
            ErrLog("CreateProductionOrder", ex)
        Finally
            rs.ReleaseComObject
        End Try

    End Sub

    Private Sub GetWorkingDays(ByRef produktionOrderLines As Generic.List(Of ProduktionOrderLine), ByRef rs As Recordset)

        Dim sb As New Text.StringBuilder()
        sb.Append("select ").AppendLine()
        sb.Append("ORCJ.CapDate, sum(ORCJ.Capacity) as Capacity").AppendLine()
        sb.Append("from ORCJ").AppendLine()
        sb.Append("where 1=1").AppendLine()
        sb.Append("and ORCJ.CapType = 'I' ").AppendLine()
        sb.Append("and ORCJ.WhsCode = '01'").AppendLine()
        sb.Append("and CapDate Between '{0}' and '{1}'").AppendLine()
        sb.Append(" and Rescode in (select ResCode from ORSC where ResGrpCod = {2}) ").AppendLine()
        sb.Append("and Capacity > 0").AppendLine()
        ' sb.Append("group by ORCJ.CapDate order by CapDate")
        sb.Append("group by ORCJ.CapDate order by CapDate Desc")

        Dim sql As String = ""
        For Each pl As ProduktionOrderLine In produktionOrderLines
            If (pl.RestQty > 0) Then
                sql = String.Format(sb.ToString(), pl.StartDate.ToDBString(), pl.EndDate.ToDBString(), pl.ResGrpCod)
                rs.DoQuery(sql)
                Do Until rs.EoF


                    pl.WorkingDates.Add(rs.ValueAsDate("CapDate"))
                    rs.MoveNext()
                Loop

                If (pl.WorkingDates.Count > 0) Then
                    pl.CapacityPerDay = Math.Round(pl.RestQty / pl.WorkingDates.Count, 3)
                End If

            End If
        Next
        sql = ""
    End Sub

    Private Sub SetCapacityCommitted(ByRef produktionOrderLines As Generic.List(Of ProduktionOrderLine), ByRef rs As Recordset)
        Dim rs4 As Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim sql As String
        Dim HPD As Double
        Dim STP As Double = 0
        Dim SEP As Double = 0
        Dim Cap As Double = 0
        Dim ComCap As Double = 0
        Dim FreeCap As Double = 0
        Dim RememberRestQty As Double = 0
        Dim cnt As Integer = 0

        Dim OworDocEntry As Integer = 0

        Dim OworRCS_PT As String = "SER"

        For Each pl As ProduktionOrderLine In produktionOrderLines
            OworDocEntry = pl.DocEntry
            Exit For
        Next


        sql = "select U_RCS_PT from owor  "
        sql = sql + "Where DocEntry = " + OworDocEntry.ToString + ""

        rs4.DoQuery(sql)

        OworRCS_PT = rs4.Fields.Item("U_RCS_PT").Value

        sql = "select  Count(Code) count from ""@RCS_CAP_COMMIT"""
        rs.DoQuery(sql)
        If (rs.ValueAsInteger("count") < 1) Then
            sql = "insert Into ""@RCS_CAP_COMMIT""(Code, Name, U_ResCode, U_DocEntry, U_LineNum, U_Capacity, U_Date) values('1', '1', 'xx', -1, -1, 0, GETDATE())"
            rs.DoQuery(sql)
        End If

        Dim code As String = "(select MAX(cast(Code as bigint)) + 1 from ""@RCS_CAP_COMMIT"")"
        Dim delete As String = "Delete from ""@RCS_CAP_COMMIT"" where U_DocEntry = {0} And U_LineNum = {1} and U_ObjType = {2} and U_ResCode = {3}"
        Dim insert As String = "insert Into ""@RCS_CAP_COMMIT""(Code, Name, U_ResCode, U_DocEntry, U_LineNum, U_Capacity, U_Date, U_ResGrpCod, U_ObjType) values({0}, {0}, '{1}', {2}, {3}, {4}, '{5}', {6}, {7})"
        Dim sbInsert As New Text.StringBuilder()
        For Each pl As ProduktionOrderLine In produktionOrderLines

            sql = "Select isnull(U_RCS_HD,8) as HPD, isnull(U_RCS_STP,8) as STP, isnull(U_RCS_SEP,8) as SEP from ORSB "
            sql = sql + "Where ResGrpCod = " + pl.ResGrpCod.ToString()

            rs.DoQuery(sql)

            HPD = rs.Fields.Item("HPD").Value

            STP = rs.Fields.Item("STP").Value

            SEP = rs.Fields.Item("SEP").Value

            If OworRCS_PT = "SER" Then
                HPD = SEP
            Else
                HPD = STP
            End If

            sbInsert.Clear()
            sbInsert.Append(String.Format(delete, pl.DocEntry, pl.LineNum, pl.ObjType, pl.VisResCode)).AppendLine()
            If (pl.CapacityPerDay > 0 Or pl.RestQty > 0) Then

                If (pl.WorkingDates.Count > 0) Then
                    RememberRestQty = pl.RestQty
                    For Each workDate As Date In pl.WorkingDates

                        sql = "select sum(ORCJ.Capacity) as Cap From ORCJ "
                        sql = sql + "Join ORSC On ORSC.ResCode = ORCJ.ResCode "
                        sql = sql + "Where ORSC.ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = " + pl.VisResCode + ") "
                        sql = sql + "And  ORCJ.CapDate BETWEEN '" + workDate.ToString("MM-dd-yyyy") + "' and '" + workDate.ToString("MM-dd-yyyy") + "' "
                        sql = sql + "And ORCJ.CapType = 'I' "
                        sql = sql + " And ORCJ.WhsCode = '01' "
                        rs4.DoQuery(sql)
                        Cap = rs4.Fields.Item("Cap").Value

                        sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
                        sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = " + pl.VisResCode + "  ) "
                        sql = sql + "And  U_Date BETWEEN '" + workDate.ToString("MM-dd-yyyy") + "' and '" + workDate.ToString("MM-dd-yyyy") + "' and owor.Status <> 'C'"

                        'sql = sql + "And  U_Date BETWEEN '" + workDate.ToString("MM-dd-yyyy") + "' and '" + workDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + pl.DocEntry.ToString() + " and owor.Status <> 'C'"

                        rs4.DoQuery(sql)
                        ComCap = rs4.Fields.Item("ComCap").Value

                        FreeCap = Cap - ComCap
                        If FreeCap > 0 Then
                            If FreeCap > (HPD * 60) Then

                                If RememberRestQty > (HPD * 60) Then
                                    RememberRestQty = RememberRestQty - (HPD * 60)
                                    pl.CapacityPerDay = (HPD * 60)
                                Else
                                    pl.CapacityPerDay = RememberRestQty
                                    RememberRestQty = 0
                                End If
                            Else
                                If RememberRestQty > FreeCap Then
                                    RememberRestQty = RememberRestQty - FreeCap
                                    pl.CapacityPerDay = FreeCap
                                Else
                                    pl.CapacityPerDay = RememberRestQty
                                    RememberRestQty = 0
                                End If
                            End If

                            sbInsert.Append(String.Format(insert, code, pl.VisResCode, pl.DocEntry, pl.LineNum, pl.CapacityPerDay.ToDBString(), workDate.ToDBString(), pl.ResGrpCod.ToString(), pl.ObjType)).AppendLine()
                        End If
                        cnt = cnt + 1
                        If (cnt >= pl.WorkingDates.Count) Then
                            If RememberRestQty > 0 Then
                                pl.CapacityPerDay = RememberRestQty
                                RememberRestQty = 0
                                sbInsert.Append(String.Format(insert, code, pl.VisResCode, pl.DocEntry, pl.LineNum, pl.CapacityPerDay.ToDBString(), workDate.ToDBString(), pl.ResGrpCod.ToString(), pl.ObjType)).AppendLine()
                            End If
                        End If
                        'rs 2021-11-09
                        ' rs.DoQuery(sbInsert.ToString())
                        ' sbInsert.Clear()
                    Next
                    cnt = 0

                    rs.DoQuery(sbInsert.ToString())
                Else
                    sbInsert.Append(String.Format(insert, code, pl.VisResCode, pl.DocEntry, pl.LineNum, pl.RestQty.ToDBString(), pl.StartDate.ToDBString(), pl.ResGrpCod, pl.ObjType)).AppendLine()
                    rs.DoQuery(sbInsert.ToString())
                End If
            Else
                rs.DoQuery(sbInsert.ToString())
            End If
        Next
    End Sub















End Module
