Imports System.Linq
Imports System.Threading

Module RCS_PLAN
    'Public ItemCodeForRuoteCopy As String
    'Public SalesOrderRowNo As String
    'Public SalesOrderDocNum As String
    'Public SalesOrderDocEntry As String
    'Public SalesOrderItemQuantity As Double
    'Public SalesOrderItemProduceQuantity As Double
    'Public ProdOrderDocEntry As String
    'Public ProdOrderItemQuantity As Double
    'Public ProdOrderDocNum As String
    'Public ProdOrderItemcode As String
    'Public ProdOrderRowNo As String
    'Public SconnContext As String 'Used for Planlagt Ubergænset

    'Public Selekt As String
    'Public oGridPur As SAPbouiCOM.Grid

    'Private SubBomLines As New System.Collections.Generic.List(Of SubBomLine)
    '  Private IsCreateAllPro As Boolean
    ' Public OrderFormUID As String = ""


    Public Sub UpdateProductionOrderLineStartEndDate(ByVal ProdDocEntry As Integer, ByVal ProdDocNum As String, ByVal PlanUnlimeted As Boolean, salesOrderDocEntry As Long, VisOrder As Integer, PlanExternal As Boolean)

        Dim test As Object = ""
        Dim MaxPlanDays As Integer = Settings.Default.MaxPlanDays

        Dim rs As SAPbobsCOM.Recordset
        Dim rs2 As SAPbobsCOM.Recordset
        Dim rs3 As SAPbobsCOM.Recordset
        Dim rs4 As SAPbobsCOM.Recordset
        Dim rs5 As SAPbobsCOM.Recordset
        Dim rs6 As SAPbobsCOM.Recordset

        rs = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'rs1 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs2 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs3 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs4 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs5 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs6 = vCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim oProd As SAPbobsCOM.ProductionOrders
        oProd = vCmp.GetBusinessObject(BoObjectTypes.oProductionOrders)
        Dim sb As New System.Text.StringBuilder

        Dim OworDocEntry As Integer = 0
        Try


            Dim iResult As Integer = 0
            Dim Msg As String = ""


            Dim PieceProduction As String = Settings.Default.PieceProduction

            Dim sql As String
            Dim PlanOk As Boolean = False
            Dim OworRCS_PT As String = "SER"

            sql = "select distinct DocEntry from owor  "
            sql = sql + "Where DocNum = " + ProdDocNum + ""
            rs.DoQuery(sql)
            OworDocEntry = rs.Fields.Item("DocEntry").Value

            sql = " Update owor Set U_RCS_PT = 'SER'  where DocEntry in (Select DocEntry from wor1 where(rtrim(CAST(Docentry as char)) +'-' + CAST(StageId as char) + '-' + CAST(ItemCode as char))  in "
            sql = sql + "(select top 1 Case when sum(Wor1.PlannedQty+ Wor1.U_RCS_ESKM * owor.PlannedQty) > " + PieceProduction + "  then rtrim(CAST(Wor1.Docentry as char)) +'-' + CAST(StageId as char)  + '-' + CAST(Wor1.ItemCode as char) else '0-0' "
            sql = sql + "End from wor1 inner join Owor On owor.DocEntry = wor1.DocEntry inner join ORSC On ORSC.ResCode = wor1.ItemCode and ORSC.QryGroup3 = 'Y' where StageID = 3 and isnull(OWOR.U_RCS_PT,'') <> 'STP' and OWor.DocEntry = " + OworDocEntry.ToString
            sql = sql + " Group by wor1.Docentry, StageId, Wor1.ITemCode, wor1.LineNum Order by wor1.LineNum DESC)) "

            rs4.DoQuery(sql)

            sql = " Update owor Set U_RCS_PT = 'STP'  where DocEntry in (Select DocEntry from wor1 where(rtrim(CAST(Docentry as char)) +'-' + CAST(StageId as char)  + '-' + CAST(ItemCode as char))  in "
            sql = sql + "(select top 1 Case when sum(Wor1.PlannedQty+ Wor1.U_RCS_ESKM * owor.PlannedQty) <= " + PieceProduction + "  then rtrim(CAST(Wor1.Docentry as char)) +'-' + CAST(StageId as char) + '-' + CAST(Wor1.ItemCode as char) else '0-0' "
            sql = sql + "End from wor1 inner join Owor On owor.DocEntry = wor1.DocEntry inner join ORSC On ORSC.ResCode = wor1.ItemCode and ORSC.QryGroup3 = 'Y' where StageID = 3 and isnull(OWOR.U_RCS_PT,'') <> 'SER' and OWor.DocEntry = " + OworDocEntry.ToString
            sql = sql + " Group by wor1.Docentry, StageId, Wor1.ITemCode, wor1.LineNum Order by wor1.LineNum DESC )) "

            rs4.DoQuery(sql)

            sql = "select U_RCS_PT, U_RCS_DelDays from owor  " ' 20251117 HMGL tilføjet deldays
            sql = sql + "Where DocNum = " + ProdDocNum + ""

            rs4.DoQuery(sql)

            OworRCS_PT = rs4.Fields.Item("U_RCS_PT").Value
            Dim RcsDelDays As Double = rs4.Fields.Item("U_RCS_DelDays").Value ' 20251117 HMGL

            Console.WriteLine("Produktionsordre er: " + OworRCS_PT)

            Do Until PlanOk = True
                PlanOk = True
                sql = "select distinct DocEntry from owor  "
                sql = sql + "Where DocNum = " + ProdDocNum + ""


                rs.DoQuery(sql)



                If rs.EoF Then
                    Exit Try
                End If

                OworDocEntry = rs.Fields.Item("DocEntry").Value


                Dim PlannedQtyTotalHour As Double

                sql = "Select sum(case when PlannedQty < 1 then 1 else PlannedQty end) as PlannedQtyTotal from WOR1 inner join orsc on wor1.itemcode = orsc.rescode "
                'sql = sql + "Where (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and ItemType = 290 and DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
                ' 20251117 HMGL
                sql = sql + "Where (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or ORSC.QryGroup15 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and ItemType = 290 and DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
                rs2.DoQuery(sql)


                sql = "Delete from ""@RCS_CAP_COMMIT"" where U_DocEntry = " + rs.Fields.Item("DocEntry").Value.ToString + ""
                rs3.DoQuery(sql)

                PlannedQtyTotalHour = rs2.Fields.Item("PlannedQtyTotal").Value / 60

                Dim DelayTotal As Double
                'Dim WeekEndTotal As Double
                Dim CntDay As Double = 0
                Dim Delay As Double = 0
                Dim Day As String = ""
                Dim WeekendDays As Integer = 0

                sql = "Select sum(1) as DelayTotal from WOR1 inner join orsc on wor1.itemcode = orsc.rescode "
                'sql = sql + "Where (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and  wor1.ItemType = 290 and wor1.DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
                ' 20251117 HMGL
                sql = sql + "Where (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or ORSC.QryGroup15 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and  wor1.ItemType = 290 and wor1.DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
                rs2.DoQuery(sql)
                DelayTotal = DelayTotal + rs2.Fields.Item("DelayTotal").Value

                Dim ResTotal As Double

                sql = "Select count(DocEntry) as ResTotal from WOR1  inner join orsc on wor1.itemcode = orsc.rescode "
                'sql = sql + "Where (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and ItemType = 290 and DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
                ' 20251117 HMGL
                sql = sql + "Where (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or ORSC.QryGroup15 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and ItemType = 290 and DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
                rs2.DoQuery(sql)

                ResTotal = rs2.Fields.Item("ResTotal").Value



                '  sql = "Select WOR1.linenum, WOR1.StageId, wor4.SeqNum-1 as SeqNum ,WOR1.PlannedQty, WOR1.VisOrder, WOR1.U_RCS_Issued, WOR1.Itemcode, 0 as U_RCS_ND from WOR1 inner join orsc on wor1.itemcode = orsc.rescode "
                '  sql = sql + " inner join wor4 on wor4.stageid = wor1.stageid and wor1.docentry = wor4.docentry "
                ' sql = sql + "Where wor1.ItemType = 290 and wor1.DocEntry =" + rs.Fields.Item("DocEntry").Value.ToString + ""
                ' sql = sql + " order by VisOrder"

                'sql = "Select  ORSC.ResGrpCod, wor1.StageId , wor4.SeqNum-1 As SeqNum ,sum(WOR1.PlannedQty+OWOR.PlannedQty*isnull(U_RCS_ESKM,0)) as PQty, sum(WOR1.U_RCS_Issued), sum(isnull(U_RCS_ED,0)) as Nd, orsc.QryGroup7, orsc.QryGroup6, orsc.QryGroup8, isnull(U_RCS_PU,'') as U_RCS_PU,  isnull(U_RCS_DEL,'') as U_RCS_DEL  from WOR1 "
                'sql = sql + "inner Join OWOR on OWOR.DocEntry =  wor1.docentry inner Join wor4 on wor4.stageid = wor1.stageid And wor1.docentry = wor4.docentry inner join orsc on wor1.itemcode = orsc.rescode left join ocrd on SubString(wor1.itemcode,3,2) = OCRD.U_RCS_CONO "
                'sql = sql + "Where  (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and  wor1.ItemType = 290 And wor1.DocEntry  =" + rs.Fields.Item("DocEntry").Value.ToString + " "
                'sql = sql + "Group by SeqNum , wor1.StageId, orsc.QryGroup7, orsc.QryGroup6, orsc.QryGroup8, isnull(U_RCS_PU,''),  isnull(U_RCS_DEL,''),  ORSC.ResGrpCod "
                'sql = sql + "order by SeqNum "

                ' BR comment out 20220321
                'sql = "Select  ORSC.ResGrpCod, wor1.StageId , wor4.SeqNum-1 As SeqNum ,sum(WOR1.PlannedQty+OWOR.PlannedQty*isnull(U_RCS_ESKM,0)) as PQty, sum(isnull(WOR1.U_RCS_Issued,0)) U_RCS_Issued, sum(isnull(U_RCS_ED,0)) as Nd, orsc.QryGroup7, orsc.QryGroup14, orsc.QryGroup6,  max(Case When orsc.QryGroup8 = 'Y' then 1 else 0 end) as QryGroup8, isnull(U_RCS_PU,'') as U_RCS_PU,  isnull(U_RCS_DEL,'') as U_RCS_DEL  from WOR1 "
                sql = "Select  ORSC.ResGrpCod, wor1.StageId , wor4.SeqNum-1 As SeqNum ,sum(WOR1.PlannedQty+OWOR.PlannedQty*isnull(U_RCS_ESKM,0)-WOR1.IssuedQty) as PQty, "
                sql = sql + "sum(isnull(WOR1.U_RCS_Issued,0)) U_RCS_Issued, sum(isnull(U_RCS_ED,0)) as Nd, orsc.QryGroup7, orsc.QryGroup14, orsc.QryGroup15, orsc.QryGroup5, orsc.QryGroup6,  max(Case When orsc.QryGroup8 = 'Y' then 1 else 0 end) as QryGroup8, isnull(U_RCS_PU,'') as U_RCS_PU,  isnull(U_RCS_DEL,'') as U_RCS_DEL  from WOR1 "
                sql = sql + "inner Join OWOR on OWOR.DocEntry =  wor1.docentry inner Join wor4 on wor4.stageid = wor1.stageid And wor1.docentry = wor4.docentry inner join orsc on wor1.itemcode = orsc.rescode left join ocrd on SubString(wor1.itemcode,3,2) = OCRD.U_RCS_CONO "
                ' -> 20251117 HMGL
                'sql = sql + "Where  (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and  wor1.ItemType = 290 And wor1.DocEntry  =" + rs.Fields.Item("DocEntry").Value.ToString + " "
                'sql = sql + "Group by SeqNum , wor1.StageId, orsc.QryGroup7, orsc.QryGroup14, orsc.QryGroup5, orsc.QryGroup6,  isnull(U_RCS_PU,''),  isnull(U_RCS_DEL,''),  ORSC.ResGrpCod "
                sql = sql + "Where  (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or ORSC.QryGroup15 = 'Y' or  orsc.QryGroup6 = 'Y' or orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and  wor1.ItemType = 290 And wor1.DocEntry  =" + rs.Fields.Item("DocEntry").Value.ToString + " "
                sql = sql + "Group by SeqNum , wor1.StageId, orsc.QryGroup7, orsc.QryGroup14, orsc.QryGroup15, orsc.QryGroup5, orsc.QryGroup6,  isnull(U_RCS_PU,''),  isnull(U_RCS_DEL,''),  ORSC.ResGrpCod "
                ' <- 20251117 HMGL
                sql = sql + "order by SeqNum "

                '  Select Case ORSC.ResGrpCod, wor1.StageId , wor4.SeqNum-1 As SeqNum ,sum(WOR1.PlannedQty+OWOR.PlannedQty*isnull(U_RCS_ESKM,0)) As PQty, 
                'sum(isnull(WOR1.U_RCS_Issued, 0)) As U_RCS_Issued, sum(isnull(U_RCS_ED,0)) As Nd, orsc.QryGroup7, orsc.QryGroup14, orsc.QryGroup6, max(Case When orsc.QryGroup8 = 'Y' then 1 else 0 end) as QryGroup8,
                'isnull(U_RCS_PU,'') as U_RCS_PU,  isnull(U_RCS_DEL,'') as U_RCS_DEL, wor4.StartDate, wor4.endDate  from WOR1 inner Join OWOR on OWOR.DocEntry =  wor1.docentry 
                'inner Join wor4 on wor4.stageid = wor1.stageid And wor1.docentry = wor4.docentry inner join orsc on wor1.itemcode = orsc.rescode 
                'Left Join ocrd On SubString(wor1.itemcode, 3, 2) = OCRD.U_RCS_CONO Where  (orsc.QryGroup7 = 'Y' or ORSC.QryGroup14 = 'Y' or  orsc.QryGroup6 = 'Y' or
                'orsc.QryGroup2 = 'Y' or orsc.QryGroup3 = 'Y') and  wor1.ItemType = 290 And wor1.DocEntry  =12111 Group by SeqNum , wor1.StageId,
                'orsc.QryGroup7, orsc.QryGroup14, orsc.QryGroup6, isnull(U_RCS_PU,''),  isnull(U_RCS_DEL,''),  ORSC.ResGrpCod, wor4.StartDate, wor4.endDate order by SeqNum 

                rs2.DoQuery(sql)
                'Dim DateCnt As Double
                Dim DateCntPlanned As Double
                Dim DateCntPlannedPerDate As Double
                'Dim CapPerDate As Double
                Dim PlannedQty As Double

                Dim HuskProdDateStartFromStart As Date
                Dim HuskProdDateEndLatest As Date

                If oProd.GetByKey(rs.Fields.Item("DocEntry").Value) Then

                    HuskProdDateStartFromStart = oProd.StartDate
                    HuskProdDateStartNy = oProd.StartDate
                    HuskProdDateEndNy = oProd.DueDate
                    HuskProdDateEndLatest = oProd.DueDate
                    '  HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)


                    Dim CapPerDayTotal As Integer = 0

                    'DateCnt = DateCnt - WeekendDays

                    'DateCnt = DateCnt - (DelayTotal - 1)

                    '  CapPerDate = DateCnt / (PlannedQtyTotalHour)
                    Dim Cnt As Integer = rs2.RecordCount - 1
                    Dim CntRecHusk As Integer = rs2.RecordCount - 1
                    Dim HuskResGrpcode As Integer = 0
                    Dim SameDate As Boolean = False
                    Dim HuskQtySameDate As Double = 0
                    Dim HuskSameDateOneShot As Boolean = False
                    Dim HPD As Double = 8
                    Dim HPDTotal As Double = 8
                    Dim STP As Double = 0
                    Dim SEP As Double = 0
                    Dim HuskStartSlutDato As String = ""
                    Dim HuskStartSlutDatoUB As String = ""
                    Dim CapTotalSameRessource As Double = 0
                    'Dim testday

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

                            CapTotalSameRessource = 0
                            SameDate = False

                            'rs 20210927
                            ' If HuskSameDateOneShot = False Then

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

                            ' End If
                            'HuskSameDateOneShot = False
                            ' If HuskSameDateOneShot = True Then
                            'HuskSameDateOneShot = False
                            'If rs2.Fields.Item("QryGroup7").Value <> "Y" Then
                            '    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                            '    If Day = "Sat" Then
                            '        HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                            '    End If
                            '    If Day = "Sun" Then
                            '        HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
                            '    End If
                            'End If

                            '  End If


                        End If




                        '   HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-rs2.Fields.Item("Nd").Value)
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

                                    'you have to check day equal to "sat" Or "sun".
                                Loop
                            End If

                            HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-(CntHusk) - WeekendDays)


                        End If



                        ' Dim testVisOrder As String = rs2.Fields.Item("VisOrder").Value
                        '  oProd.Lines.SetCurrentLine(rs2.Fields.Item("VisOrder").Value)
                        Dim stdate = oProd.Stages.StartDate
                        Dim endate = oProd.Stages.EndDate
                        '   oProd.Update()

                        Dim testseqNum = rs2.Fields.Item("SeqNum").Value






                        oProd.Stages.SetCurrentLine(rs2.Fields.Item("SeqNum").Value)


                        'testday = oProd.Stages.EndDate
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

                        sql = "Select isnull(U_RCS_HD,8) as HPD, isnull(U_RCS_STP,8) as STP, isnull(U_RCS_SEP,8) as SEP, isnull(U_RCS_BD,0) as BD from ORSB "
                        sql = sql + "Where ResGrpCod = " + HuskResGrpcode.ToString()

                        rs4.DoQuery(sql)

                        Dim BackDates As Integer = rs4.Fields.Item("BD").Value

                        HPD = rs4.Fields.Item("HPD").Value
                        If HPD < 1 Then ' BR
                            HPD = 8
                        End If

                        HPDTotal = HPD

                        STP = rs4.Fields.Item("STP").Value

                        SEP = rs4.Fields.Item("SEP").Value


                        If OworRCS_PT = "SER" Then

                            '   If SEP > STP Then
                            '  SEP = SEP - STP
                            '  End If

                            HPD = SEP
                        Else
                            HPD = STP
                        End If



                        PlannedQty = rs2.Fields.Item("PQty").Value / 60
                        If PlannedQty < 1 Then
                            PlannedQty = 1
                        End If


                        If rs2.Fields.Item("QryGroup7").Value = "Y" Then

                            'DateCntPlannedPerDate = (PlannedQty / 8) + 1 ' Antal dage
                            DateCntPlannedPerDate = (PlannedQty / 8) '+ 1 ' Antal dage
                            DateCntPlannedPerDate = DateCntPlannedPerDate + RcsDelDays ' 20251117 HMGL

                        ElseIf rs2.Fields.Item("QryGroup6").Value = "Y" Then
                            DateCntPlannedPerDate = (PlannedQty / 8) + 1 ' Antal dage
                        ElseIf rs2.Fields.Item("QryGroup15").Value = "Y" Then ' 20251117 HMGL
                            HPD = 8
                            DateCntPlannedPerDate = PlannedQty / HPD ' Antal dage
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
                                'rs 20210927
                                ' DateCntPlanned = 2
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
                        'rs 20220207
                        oProd.Stages.EndDate = HuskProdDateEndNy

                        HuskProdDateEndLatest = HuskProdDateEndNy
                        '   CntDay = DateCntPlanned
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


                        Day = oProd.Stages.EndDate.DayOfWeek.ToString().Substring(0, 3)


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
                                '  HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                'Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
                                'If Day = "Sat" Then
                                '    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                'End If
                                'If Day = "Sun" Then
                                '    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
                                'End If
                                CntDel = CntDel + 1
                                DayNum = "5"
                            Loop

                            oProd.Stages.StartDate = HuskProdDateEndNy
                        End If

                        'rs 20210927 
                        Dim HuskNewEndDate As Date
                        If rs2.Fields.Item("QryGroup7").Value = "Y" Or 1 = 1 Then

                            If rs2.Fields.Item("QryGroup7").Value = "Y" Then
                                HuskNewEndDate = oProd.Stages.EndDate.AddDays(-RcsDelDays) ' 20260210 HMGL
                            Else
                            HuskNewEndDate = oProd.Stages.EndDate
                            End If


                            If oProd.Stages.StartDate >= oProd.Stages.EndDate Then
                                HuskNewEndDate = oProd.Stages.StartDate
                            End If

                        Else
                            HuskNewEndDate = oProd.Stages.EndDate.AddDays(-1)
                            If oProd.Stages.StartDate >= oProd.Stages.EndDate.AddDays(-1) Then
                                HuskNewEndDate = oProd.Stages.StartDate
                            End If

                        End If


                        ' Dim HuskNewEndDate As Date = oProd.Stages.EndDate.AddDays(-1)
                        'If oProd.Stages.StartDate >= oProd.Stages.EndDate.AddDays(-1) Then
                        'HuskNewEndDate = oProd.Stages.StartDate
                        'End If

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


                        ' Find committed capacited
                        sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry "
                        sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                        sql = sql + "And  U_Date BETWEEN '" + oProd.Stages.StartDate.ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"

                        '  sql = sql + "And  isnull(U_ResType,'" + OworRCS_PT + "') = '" + OworRCS_PT + "'"

                        rs4.DoQuery(sql)

                        ' Find committed capacited STP
                        sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCapSTP FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry "
                        sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                        sql = sql + "And  U_Date BETWEEN '" + oProd.Stages.StartDate.ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"
                        sql = sql + "And  isnull(U_ResType,'STP') = 'STP'"
                        'sql = sql + "And  isnull(U_ResType,'" + OworRCS_PT + "') = '" + OworRCS_PT + "'"

                        rs5.DoQuery(sql)


                        Dim test1, test2, test3, test4, testdate1, testdate2
                        test1 = rs2.Fields.Item("PQty").Value
                        test2 = rs.Fields.Item("Cap").Value
                        test3 = rs4.Fields.Item("ComCap").Value
                        test4 = rs5.Fields.Item("ComCapSTP").Value


                        Dim CapStpSer As Double

                        Dim StpSerDateSpan As Integer = DateDiff(DateInterval.Day, oProd.Stages.StartDate, HuskNewEndDate)


                        If OworRCS_PT = "SER" Then

                            If test2 > ((STP * 60) * StpSerDateSpan) Then
                                CapStpSer = test2 - ((STP * 60) * StpSerDateSpan)
                            Else
                                CapStpSer = test2
                            End If

                        Else
                            If test2 > ((STP * 60) * StpSerDateSpan) Then
                                CapStpSer = (STP * 60) * StpSerDateSpan
                            Else
                                CapStpSer = test2
                            End If

                        End If

                        'If rs2.Fields.Item("QryGroup5").Value = "Y" Or rs2.Fields.Item("QryGroup6").Value = "Y" Or rs2.Fields.Item("QryGroup7").Value = "Y" Or rs2.Fields.Item("QryGroup14").Value = "Y" Then
                        ' 20251117 HMGL
                        If rs2.Fields.Item("QryGroup5").Value = "Y" Or rs2.Fields.Item("QryGroup6").Value = "Y" Or rs2.Fields.Item("QryGroup7").Value = "Y" Or rs2.Fields.Item("QryGroup14").Value = "Y" Or rs2.Fields.Item("QryGroup15").Value = "Y" Then
                            CapStpSer = test2
                        End If


                        '  pos = (oProd.Stages.StageID - 1).ToString

                        testdate1 = oProd.Stages.StartDate.ToString("MM-dd-yyyy")
                        testdate2 = HuskNewEndDate.ToString("MM-dd-yyyy")

                        Dim HuskPosEndDate As String = HuskNewEndDate.ToString("dd-MM-yyyy")
                        HuskStartSlutDatoUB = HuskStartSlutDatoUB + " Pos: " + pos.ToString + " Dato: " + oProd.Stages.StartDate.ToString("dd-MM-yyyy") + ":" + HuskNewEndDate.ToString("dd-MM-yyyy")


                        '  If 1 = 1 And rs2.Fields.Item("PQty").Value > (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value - HuskQtySameDate) And PlanUnlimeted = False Then
                        'rs 2021-12-11 If 1 = 1 And PlanUnlimeted = False Or rs2.Fields.Item("PQty").Value > (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value - HuskQtySameDate) And PlanUnlimeted = False Then
                        'If rs2.Fields.Item("PQty").Value > (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value - HuskQtySameDate) Then
                        If 1 = 1 Then

                            Dim CntDay3 As Integer = 0
                            Dim Cnt2 As Integer = 0
                            Dim sql1 As String = ""
                            Dim Cnt4 As Integer = 0
                            Dim BookState As Boolean = False
                            Dim CanStartBook As Boolean = False
                            CapPerDayTotal = 0
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



                            StpSerDateSpan = DateDiff(DateInterval.Day, oProd.Stages.StartDate.AddDays(CntDay3), HuskNewEndDate) + 1


                            If OworRCS_PT = "SER" Then

                                If test2 > ((STP * 60) * StpSerDateSpan) Then
                                    CapStpSer = test2 - ((STP * 60) * StpSerDateSpan)
                                Else
                                    CapStpSer = test2
                                End If

                            Else
                                If test2 > ((STP * 60) * StpSerDateSpan) Then
                                    CapStpSer = (STP * 60) * StpSerDateSpan
                                Else
                                    CapStpSer = test2
                                End If

                            End If

                            'If rs2.Fields.Item("QryGroup5").Value = "Y" Or rs2.Fields.Item("QryGroup6").Value = "Y" Or rs2.Fields.Item("QryGroup7").Value = "Y" Or rs2.Fields.Item("QryGroup14").Value = "Y" Then
                            ' 20251117 HMGL
                            If rs2.Fields.Item("QryGroup5").Value = "Y" Or rs2.Fields.Item("QryGroup6").Value = "Y" Or rs2.Fields.Item("QryGroup7").Value = "Y" Or rs2.Fields.Item("QryGroup14").Value = "Y" Or rs2.Fields.Item("QryGroup15").Value = "Y" Then
                                CapStpSer = test2
                            End If


                            sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
                            sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                            sql = sql + "And  U_Date BETWEEN '" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"

                            sql = sql + "And  isnull(U_ResType,'" + OworRCS_PT + "') = '" + OworRCS_PT + "'"

                            rs6.DoQuery(sql)


                            test3 = rs6.Fields.Item("ComCap").Value


                            '   Do Until (rs2.Fields.Item("PQty").Value <= (rs5.Fields.Item("Cap").Value - rs6.Fields.Item("ComCap").Value - HuskQtySameDate)) And (rs2.Fields.Item("PQty").Value <= CapPerDayTotal)
                            '  Do Until (rs2.Fields.Item("PQty").Value <= (rs5.Fields.Item("Cap").Value - rs6.Fields.Item("ComCap").Value - HuskQtySameDate)) Or (rs2.Fields.Item("PQty").Value <= CapPerDayTotal)
                            Do Until (rs2.Fields.Item("PQty").Value <= CapPerDayTotal)
                                ' Do Until (rs2.Fields.Item("PQty").Value <= CapPerDayTotal - CapTotalSameRessource)20230404

                                If rs2.Fields.Item("QryGroup6").Value = "Y" Then
                                    '  Exit Do
                                End If
                                ' SBO_Application.SetStatusBarMessage("Ikke nok resource til artikel nr. " + oProd.ItemNo + " pos: " + oProd.Stages.Name + " Planlagt: " + rs2.Fields.Item("PQty").Value.ToString + " muligt forbrug: " + (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value).ToString + "I perioden " + oProd.Stages.StartDate.AddDays(CntDay3).ToString() + " " + HuskNewEndDate.ToString())


                                sql = "select sum(ORCJ.Capacity) as Cap From ORCJ "
                                sql = sql + "Join ORSC On ORSC.ResCode = ORCJ.ResCode "
                                sql = sql + "Where ORSC.ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                                sql = sql + "And  ORCJ.CapDate BETWEEN '" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' "
                                sql = sql + "And ORCJ.CapType = 'I' "
                                sql = sql + " And ORCJ.WhsCode = '01' "
                                rs5.DoQuery(sql)
                                test2 = rs5.Fields.Item("Cap").Value


                                StpSerDateSpan = DateDiff(DateInterval.Day, oProd.Stages.StartDate.AddDays(CntDay3), HuskNewEndDate) + 1


                                If OworRCS_PT = "SER" Then

                                    If test2 > ((STP * 60) * StpSerDateSpan) Then
                                        CapStpSer = test2 - ((STP * 60) * StpSerDateSpan)
                                    Else
                                        CapStpSer = test2
                                    End If

                                Else
                                    If test2 > ((STP * 60) * StpSerDateSpan) Then
                                        CapStpSer = (STP * 60) * StpSerDateSpan
                                    Else
                                        CapStpSer = test2
                                    End If

                                End If

                                'If rs2.Fields.Item("QryGroup5").Value = "Y" Or rs2.Fields.Item("QryGroup6").Value = "Y" Or rs2.Fields.Item("QryGroup7").Value = "Y" Or rs2.Fields.Item("QryGroup14").Value = "Y" Then
                                ' 20251117 HMGL
                                If rs2.Fields.Item("QryGroup5").Value = "Y" Or rs2.Fields.Item("QryGroup6").Value = "Y" Or rs2.Fields.Item("QryGroup7").Value = "Y" Or rs2.Fields.Item("QryGroup14").Value = "Y" Or rs2.Fields.Item("QryGroup15").Value = "Y" Then
                                    CapStpSer = test2
                                End If


                                sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
                                sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                                sql = sql + "And  U_Date BETWEEN '" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"
                                sql = sql + "And  isnull(U_ResType,'" + OworRCS_PT + "') = '" + OworRCS_PT + "'"

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


                                StpSerDateSpan = DateDiff(DateInterval.Day, HuskNewEndDate.AddDays(CntDay3), HuskNewEndDate.AddDays(CntDay3)) + 1

                                Dim BackDatesDate = Now.AddDays(BackDates)

                                'Check for weekend
                                WeekendDays = 0
                                CntDayHusk = BackDates
                                CntDay = BackDates
                                CntDay2 = 0
                                If CntDay > 1 Then

                                    Do Until ((CntDay - 1 + WeekendDays) <= 0)

                                        Day = BackDatesDate.AddDays(CntDayHusk - CntDay2 - 1).DayOfWeek.ToString().Substring(0, 3)
                                        If Day = "Sat" Or Day = "Sun" Then

                                            WeekendDays = WeekendDays + 1

                                        End If
                                        CntDay = CntDay - 1
                                        CntDay2 = CntDay2 + 1

                                        'you have to check day equal to "sat" Or "sun".
                                    Loop

                                End If

                                BackDatesDate = BackDatesDate.AddDays(WeekendDays)


                                If OworRCS_PT = "SER" Then

                                    If test2 > ((STP * 60) * StpSerDateSpan) Then
                                        If HuskNewEndDate.AddDays(CntDay3) <= BackDatesDate Then
                                            CapStpSer = test2
                                            HPD = test2 / 60
                                        Else
                                            CapStpSer = test2 - ((STP * 60) * StpSerDateSpan)
                                        End If
                                    Else
                                        CapStpSer = test2
                                    End If

                                Else
                                    If test2 > ((STP * 60) * StpSerDateSpan) Then
                                        CapStpSer = (STP * 60) * StpSerDateSpan
                                    Else
                                        CapStpSer = test2
                                    End If

                                End If
                                'If rs2.Fields.Item("QryGroup5").Value = "Y" Or rs2.Fields.Item("QryGroup6").Value = "Y" Or rs2.Fields.Item("QryGroup7").Value = "Y" Or rs2.Fields.Item("QryGroup14").Value = "Y" Then
                                ' 20251117 HMGL
                                If rs2.Fields.Item("QryGroup5").Value = "Y" Or rs2.Fields.Item("QryGroup6").Value = "Y" Or rs2.Fields.Item("QryGroup7").Value = "Y" Or rs2.Fields.Item("QryGroup14").Value = "Y" Or rs2.Fields.Item("QryGroup15").Value = "Y" Then
                                    CapStpSer = test2
                                End If

                                sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
                                sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                                sql = sql + "And  U_Date BETWEEN '" + HuskNewEndDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"

                                ' sql = sql + "And  isnull(U_ResType,'" + OworRCS_PT + "') = '" + OworRCS_PT + "'"

                                rs4.DoQuery(sql)


                                sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCapSERSTP FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
                                sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = (select top 1 wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) ) "
                                sql = sql + "And  U_Date BETWEEN '" + HuskNewEndDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and '" + HuskNewEndDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + oProd.DocumentNumber.ToString + " and owor.Status <> 'C'"
                                sql = sql + "And  isnull(U_ResType,'') = '" + OworRCS_PT + "'"

                                rs6.DoQuery(sql)
                                test4 = rs6.Fields.Item("ComCapSERSTP").Value



                                test1 = rs2.Fields.Item("PQty").Value
                                test2 = rs.Fields.Item("Cap").Value
                                test3 = rs4.Fields.Item("ComCap").Value


                                'if capacity > comitted capacity
                                If ((CapStpSer - rs6.Fields.Item("ComCapSERSTP").Value) > 0) And ((rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value) > 0) Then

                                    ' if free capacity > max capacity per day
                                    If (CapStpSer - rs6.Fields.Item("ComCapSERSTP").Value) > (HPD * 60) Then

                                        CapPerDayTotal = CapPerDayTotal + (HPD * 60)

                                    Else
                                        If PlanUnlimeted = True And CapStpSer > 0 Then
                                            CapPerDayTotal = CapPerDayTotal + (HPD * 60)
                                        Else

                                            If (CapStpSer - rs6.Fields.Item("ComCapSERSTP").Value) > (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value) Then
                                                CapPerDayTotal = CapPerDayTotal + (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value)
                                            Else
                                                CapPerDayTotal = CapPerDayTotal + (CapStpSer - rs6.Fields.Item("ComCapSERSTP").Value)
                                            End If



                                        End If


                                    End If

                                    '  CntDay3 = CntDay3 - Math.Ceiling((rs2.Fields.Item("PQty").Value - (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value)) / (HPD * 60))
                                    CntDay3 = CntDay3 - 1
                                Else
                                    If PlanUnlimeted = True And CapStpSer > 0 Then
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
                                '  rs.MoveNext()

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
                                        '  HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                        'Day = HuskProdDateEndNy.DayOfWeek.ToString().Substring(0, 3)
                                        'If Day = "Sat" Then
                                        '    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-1)
                                        'End If
                                        'If Day = "Sun" Then
                                        '    HuskProdDateEndNy = HuskProdDateEndNy.AddDays(-2)
                                        'End If
                                        CntDel = CntDel + 1
                                        DayNum = "5"
                                    Loop

                                    ' oProd.Stages.StartDate = HuskProdDateEndNy
                                End If




                                'Check for passing startdate
                                'rs 2021-1´2-11 If DateDiff(DateInterval.Day, oProd.StartDate, HuskProdDateEndNy) < 0 Then
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
                                        If PlanExternal Then

                                            'Show message in Console
                                            Console.WriteLine(Msg)
                                            Console.WriteLine()

                                        Else
                                            SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                        End If


                                    Else
                                        Msg = "Produktionsordre for artikel " + oProd.ItemNo.ToString + " er opdateret med ny dato. Stopper planlæggning ved pos: " + pos.ToString + " med slutdato: " + HuskNewEndDate.ToString("dd-MM-yyyy")

                                        If PlanExternal Then

                                            'Show message in Console
                                            Console.WriteLine(Msg)
                                            Console.WriteLine()

                                        Else
                                            SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                        End If


                                    End If


                                    sql = "Update RDR1 set U_RCS_AFS = '" + oProd.DueDate.ToString("yyyyMMdd") + "'"
                                    sql = sql + " Where VisOrder = " + VisOrder.ToString + " and DocEntry = " + salesOrderDocEntry.ToString
                                    rs4.DoQuery(sql)

                                    'SBO_Application.SetStatusBarMessage(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " Prøver med forfaldsdato: " + oProd.DueDate.ToString("dd-MM-yyyy"))
                                    'SBO_Application.SetStatusBarMessage(HuskStartSlutDato)

                                    If PlanExternal Then

                                        'Show message in Console
                                        Console.WriteLine("Planlægning fejler. Artikelnr.: " + oProd.ItemNo.ToString)
                                        Console.WriteLine()

                                    Else

                                        SBO_Application.StatusBar.SetText("Planlægning fejler. Artikelnr.: " + oProd.ItemNo.ToString, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                    End If



                                    PlanOk = False


                                    Exit Do
                                End If

                                'Check for passing startdate
                                'rs 2021-12-11 If DateDiff(DateInterval.Day, oProd.StartDate, HuskProdDateEndNy) < 0 Then
                                If (DateDiff(DateInterval.Day, Date.Now, HuskProdDateEndNy) < 0) Or (DateDiff(DateInterval.Day, oProd.StartDate, HuskProdDateEndNy) < 0) Then


                                    If PlanExternal Then

                                        'Show message in Console
                                        Console.WriteLine(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " Har passeret startdato stopper planlægning! " + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' mellem '" + HuskNewEndDate.ToString("MM-dd-yyyy"))
                                        Console.WriteLine()

                                    Else

                                        SBO_Application.SetStatusBarMessage(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " Har passeret startdato stopper planlægning! " + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' mellem '" + HuskNewEndDate.ToString("MM-dd-yyyy"))
                                    End If


                                    Cnt = 0
                                    Exit Do
                                End If
                                'Check for max 60 days
                                If Cnt2 > MaxPlanDays Then

                                    If PlanExternal Then

                                        'Show message in Console
                                        Console.WriteLine(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + "Har passeret 30 dage stopper planlægning!" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' mellem '" + HuskNewEndDate.ToString("MM-dd-yyyy"))
                                        Console.WriteLine()

                                    Else

                                        SBO_Application.SetStatusBarMessage(oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + "Har passeret 30 dage stopper planlægning!" + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + "' mellem '" + HuskNewEndDate.ToString("MM-dd-yyyy"))
                                    End If



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

                                    If PlanExternal Then

                                        'Show message in Console
                                        Console.WriteLine(Tekst)
                                        Console.WriteLine()

                                    Else

                                        SBO_Application.MessageBox(Tekst)
                                    End If





                                    Cnt = 0
                                    Exit Do
                                End If
                                Cnt2 = Cnt2 + 1
                            Loop
                            '  SBO_Application.SetStatusBarMessage("Ikke nok resource til artikel nr. " + oProd.ItemNo + " pos: " + oProd.Stages.Name + " Planlagt: " + rs2.Fields.Item("PQty").Value.ToString + " muligt forbrug: " + (rs.Fields.Item("Cap").Value - rs4.Fields.Item("ComCap").Value).ToString + "I perioden " + oProd.Stages.StartDate.AddDays(CntDay3).ToString("MM-dd-yyyy") + " " + HuskNewEndDate.ToString("MM-dd-yyyy"))

                            'RS20220207  
                            oProd.Stages.StartDate = HuskProdDateEndNy
                            HuskStartSlutDato = HuskStartSlutDato + " Pos: " + pos.ToString + " Dato: " + HuskProdDateEndNy.ToString("dd-MM-yyyy") + ":" + HuskPosEndDate

                            'SBO_Application.SetStatusBarMessage("Ikke nok resource til artikel nr. " + oProd.ItemNo + " pos: " + oProd.Stages.Name + " Planlagt: " + rs2.Fields.Item("PQty").Value.ToString + " muligt forbrug: " + (test2 - test3).ToString)

                        End If

                        CapTotalSameRessource = CapPerDayTotal

                        HuskQtySameDate = HuskQtySameDate + rs2.Fields.Item("PQty").Value

                        'SBO_Application.SetStatusBarMessage("Fundet resource til artikel nr. " + oProd.ItemNo + " Pos : " + oProd.Stages.Name + " Start: " + oProd.Stages.StartDate.ToString + " Slut: " + HuskNewEndDate.ToString, BoMessageTime.bmt_Long, False)
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

                        'rs 20210927    
                        If rs2.Fields.Item("QryGroup7").Value = "Y" Then
                            ' oProd.Stages.StartDate = HuskProdDateEndNy2
                            '  oProd.Stages.StartDate = oProd.Stages.StartDate.AddDays(1)
                        End If

                        'oProd.Stages.StartDate = HuskProdDateEndNy2

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

                        'If PlanUnlimeted = True And oProd.StartDate > oProd.Stages.StartDate Then
                        '    Dim Cnt2 As Integer = 0
                        '    Do Until Cnt2 > oProd.Stages.Count - 1
                        '        oProd.Stages.SetCurrentLine(Cnt2)
                        '        oProd.Stages.StartDate = Date.Now
                        '        oProd.Stages.EndDate = oProd.DueDate

                        '        Cnt2 = Cnt2 + 1
                        '    Loop

                        'End If
                        Dim test1a As Date = oProd.StartDate
                        Dim test2a As Date = oProd.Stages.StartDate

                        If oProd.StartDate > oProd.Stages.StartDate And PlanUnlimeted = False Then

                            Dim Missday

                            Missday = DateDiff(DateInterval.Day, oProd.StartDate, oProd.Stages.StartDate)

                            Msg = "Produktionsordre " + oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " er ikke opdateret! er forbi start dato!  "
                            'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
                            ' SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error)
                            'Exit Sub


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

                                If PlanExternal Then

                                    'Show message in Console
                                    Console.WriteLine(Msg)
                                    Console.WriteLine()

                                Else

                                    SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                End If





                            Else
                                Msg = "Produktionsordre for artikel " + oProd.ItemNo.ToString + " er opdateret med ny dato " + HuskStartSlutDato


                                If PlanExternal Then

                                    'Show message in Console
                                    Console.WriteLine(Msg)
                                    Console.WriteLine()

                                Else

                                    SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                End If






                            End If


                            sql = "Update RDR1 set U_RCS_AFS = '" + oProd.DueDate.ToString("yyyyMMdd") + "'"
                            sql = sql + " Where VisOrder = " + VisOrder.ToString + " and DocEntry = " + salesOrderDocEntry.ToString
                            rs4.DoQuery(sql)

                            Msg = oProd.ItemNo.ToString + " Pos: " + (oProd.Stages.StageID - 1).ToString + " Prøver med forfaldsdato: " + oProd.DueDate.ToString("dd-MM-yyyy")

                            If PlanExternal Then

                                'Show message in Console
                                Console.WriteLine(Msg)
                                Console.WriteLine()

                            Else

                                SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            End If




                            PlanOk = False


                            ' Exit Do

                            'oProd.StartDate = oProd.Stages.StartDate
                        End If

                        '  If PlanUnlimeted = False Then
                        oProd.Stages.SetCurrentLine(0)
                        oProd.Stages.EndDate = HuskProdDateEndNy
                        sql = "Select max(isnull(LeadTime,0)) as PQty from OITM "
                        sql = sql + "Where ItemCode in (select wor1.itemcode from wor1 where wor1.stageid = " + oProd.Stages.StageID.ToString + " and wor1.docentry = " + oProd.AbsoluteEntry.ToString + " ) "


                        rs4.DoQuery(sql)
                        Dim LeadTime As Integer
                        LeadTime = rs4.Fields.Item("PQty").Value

                        'rs 20210502
                        oProd.Stages.StartDate = HuskProdDateEndNy.AddDays(-1 * LeadTime)
                        ' End If



                        Dim teststart

                        teststart = oProd.Stages.StartDate

                        If HuskProdDateStartFromStart <= oProd.Stages.StartDate Or PlanUnlimeted = True Then
                            '   oProd.SaveXML("c:\rcs\prod.xml")
                            iResult = oProd.Update()
                            If iResult < 0 Then
                                vCmp.GetLastError(iResult, Msg)
                                ErrorLog("CreateCopyOfBillOfMaterial", Msg)


                                If PlanExternal Then

                                    'Show message in Console
                                    Console.WriteLine(Msg)
                                    Console.WriteLine()

                                Else

                                    SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                End If



                                If PlanUnlimeted Then
                                    Msg = HuskStartSlutDatoUB
                                    If PlanExternal Then

                                        'Show message in Console
                                        Console.WriteLine(Msg)
                                        Console.WriteLine()

                                    Else

                                        SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                    End If




                                End If
                            Else


                                '  String code = "(select MAX(cast(Code as bigint)) + 1 from \"@RCS_CAP_COMMIT\")";
                                '          String delete = "Delete from \"@RCS_CAP_COMMIT\" where U_DocEntry = {0} And U_LineNum = {1} and U_ObjType = {2} and U_ResCode = {3}";
                                '          String insert = "insert Into \"@RCS_CAP_COMMIT\"(Code, Name, U_ResCode, U_DocEntry, U_LineNum, U_Capacity, U_Date, U_ResGrpCod, U_ObjType) values({0}, {0}, '{1}', {2}, {3}, {4}, '{5}', {6}, {7})";

                                SetResourceCapacityCommitted(oProd.AbsoluteEntry)

                                Msg = "Produktionsordre for artikel " + oProd.ItemNo.ToString + " er opdateret "
                                'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
                                If PlanExternal Then

                                    'Show message in Console
                                    Console.WriteLine(Msg)
                                    Console.WriteLine()

                                Else

                                    SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                End If

                            End If
                        Else
                            If PlanOk = False Then
                                Msg = "Produktionsordre for artikel " + oProd.ItemNo.ToString + " er ikke opdateret! "
                                'Msg = String.Format(Msg, oProd.DocumentNumber.ToString())
                                If PlanExternal Then

                                    'Show message in Console
                                    Console.WriteLine(Msg)
                                    Console.WriteLine()

                                Else

                                    SBO_Application.StatusBar.SetText(Msg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                End If

                            End If

                        End If

                    End If
                End If
            Loop

            'oSO.ReleaseComObject() ' = Nothing
            ' oProd.ReleaseComObject() ' = Nothing
            'oRe.ReleaseComObject() ' = Nothing

        Catch ex As Exception
            ErrLog("UpdateProductionOrderLineStartEndDate", ex)
        Finally
            rs.ReleaseComObject
            rs2.ReleaseComObject
            rs3.ReleaseComObject
            rs4.ReleaseComObject
            oProd.ReleaseComObject
        End Try


    End Sub

    Private Sub SetResourceCapacityCommitted(ByVal DocEntry As Integer)

        Dim rs As SAPbobsCOM.Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim rs4 As SAPbobsCOM.Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim sql As String = ""

        Try
            Dim sb As New System.Text.StringBuilder()
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

            Dim produktionOrderLines As New System.Collections.Generic.List(Of ProduktionOrderLine)
            Dim pl As ProduktionOrderLine
            pl = Nothing
            Do While Not rs.EoF

                '   Sql = "Select isnull(U_RCS_HD,8) as HPD from ORSB "
                '  sql = sql + "Where ResGrpCod = " + rs.ValueAsInteger("ResGrpCod").ToString()

                '  rs4.DoQuery(Sql)

                '  HPD = rs4.Fields.Item("HPD").Value

                '    DateCntPlannedPerDate = PlannedQty / HPD ' Antal dage








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


                'RememberRestQty = pl.RestQty
                'Do Until RememberRestQty <= 0

                '    sql = "select sum(ORCJ.Capacity) as Cap From ORCJ "
                '    sql = sql + "Join ORSC On ORSC.ResCode = ORCJ.ResCode "
                '    sql = sql + "Where ORSC.ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = " + rs.ValueAsString("VisResCode") + ") "
                '    sql = sql + "And  ORCJ.CapDate BETWEEN '" + pl.StartDate.ToString("MM-dd-yyyy") + "' and '" + pl.StartDate.ToString("MM-dd-yyyy") + "' "
                '    sql = sql + "And ORCJ.CapType = 'I' "
                '    sql = sql + " And ORCJ.WhsCode = '01' "
                '    rs4.DoQuery(sql)
                '    Cap = rs4.Fields.Item("Cap").Value

                '    sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
                '    sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = " + rs.ValueAsString("VisResCode") + "  ) "
                '    sql = sql + "And  U_Date BETWEEN '" + pl.StartDate.ToString("MM-dd-yyyy") + "' and '" + pl.StartDate.ToString("MM-dd-yyyy") + "' and U_DocEntry <> " + rs.ValueAsInteger("DocEntry").ToString() + " and owor.Status <> 'C'"

                '    rs4.DoQuery(sql)
                '    ComCap = rs4.Fields.Item("ComCap").Value

                '    FreeCap = Cap - ComCap
                '    If FreeCap > 0 Then
                '        If FreeCap > (HPD * 60) Then

                '            If pl.RestQty > (HPD * 60) Then
                '                RememberRestQty = RememberRestQty - (HPD * 60)
                '                pl.RestQty = (HPD * 60)
                '            Else
                '                pl.RestQty = RememberRestQty
                '                RememberRestQty = 0
                '            End If
                '        Else
                '            If pl.RestQty > FreeCap Then
                '                RememberRestQty = RememberRestQty - FreeCap
                '                pl.RestQty = FreeCap
                '            Else
                '                pl.RestQty = RememberRestQty
                '                RememberRestQty = 0
                '            End If
                '        End If



                produktionOrderLines.Add(pl)

                '  End If

                'pl.StartDate = pl.StartDate.AddDays(1)


                ' Loop

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

    Private Sub GetWorkingDays(ByRef produktionOrderLines As System.Collections.Generic.List(Of ProduktionOrderLine), ByRef rs As SAPbobsCOM.Recordset)

        Dim sb As New System.Text.StringBuilder()
        sb.Append("select ").AppendLine()
        sb.Append("ORCJ.CapDate, sum(ORCJ.Capacity) as Capacity").AppendLine()
        sb.Append("from ORCJ").AppendLine()
        sb.Append("where 1=1").AppendLine()
        sb.Append("and ORCJ.CapType = 'I' ").AppendLine()
        sb.Append("and ORCJ.WhsCode = '01'").AppendLine()
        sb.Append("and CapDate Between '{0}' and '{1}'").AppendLine()
        sb.Append(" and Rescode in (select ResCode from ORSC where ResGrpCod = {2}) ").AppendLine()
        sb.Append("and Capacity > 0").AppendLine()
        sb.Append("group by ORCJ.CapDate order by CapDate")
        'sb.Append("group by ORCJ.CapDate order by CapDate Desc")

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

    Private Sub SetCapacityCommitted(ByRef produktionOrderLines As System.Collections.Generic.List(Of ProduktionOrderLine), ByRef rs As SAPbobsCOM.Recordset)
        Dim rs4 As SAPbobsCOM.Recordset = vCmp.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim sql As String
        Dim HPD As Double
        Dim STP As Double = 0
        Dim SEP As Double = 0
        Dim Cap As Double = 0
        Dim ComCap As Double = 0
        Dim ComCapTotal As Double = 0
        Dim FreeCap As Double = 0
        Dim FreeCapTotal As Double = 0
        Dim RememberRestQty As Double = 0
        Dim cnt As Integer = 0
        Dim OworDocEntry As Integer = 0

        Dim OworRCS_PT As String = "SER"
        Dim pl As ProduktionOrderLine


        For Each pl1 As ProduktionOrderLine In produktionOrderLines
            OworDocEntry = pl1.DocEntry
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
        ' Dim delete As String = "Delete from ""@RCS_CAP_COMMIT"" where U_DocEntry = {0} And U_LineNum = {1} and U_ObjType = {2} and U_ResCode = {3} and isnull(U_ResType,'" + OworRCS_PT + "') = '{4}'"
        Dim delete As String = "Delete from ""@RCS_CAP_COMMIT"" where U_DocEntry = {0} And U_LineNum = {1} and U_ObjType = {2} and U_ResCode = {3} "
        Dim insert As String = "insert Into ""@RCS_CAP_COMMIT""(Code, Name, U_ResCode, U_DocEntry, U_LineNum, U_Capacity, U_Date, U_ResGrpCod, U_ObjType, U_ResType) values({0}, {0}, '{1}', {2}, {3}, {4}, '{5}', {6}, {7}, '{8}')"
        Dim sbInsert As New System.Text.StringBuilder()

        For i As Integer = produktionOrderLines.Count - 1 To 0 Step -1

            pl = produktionOrderLines(i)

            ' For Each pl As ProduktionOrderLine In produktionOrderLines

            sql = "Select isnull(U_RCS_HD,8) as HPD, isnull(U_RCS_STP,8) as STP, isnull(U_RCS_SEP,8) as SEP, isnull(U_RCS_BD,0) as BD from ORSB "
            sql = sql + "Where ResGrpCod = " + pl.ResGrpCod.ToString()

            rs.DoQuery(sql)


            Dim BackDates As Integer = rs.Fields.Item("BD").Value


            HPD = rs.Fields.Item("HPD").Value


            STP = rs.Fields.Item("STP").Value

            SEP = rs.Fields.Item("SEP").Value


            If OworRCS_PT = "SER" Then

                '    If SEP > STP Then
                '   SEP = SEP - STP
                'End If

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

                        Dim CapStpSer As Double

                        Dim StpSerDateSpan As Integer = DateDiff(DateInterval.Day, workDate, workDate) + 1

                        Dim BackDatesDate = Now.AddDays(BackDates)

                        'Check for weekend
                        Dim Day As String
                        Dim WeekendDays As Integer = 0
                        Dim CntDayHusk As Integer = BackDates
                        Dim CntDay As Integer = BackDates
                        Dim CntDay2 As Integer = 0
                        If CntDay > 1 Then

                            Do Until ((CntDay - 1 + WeekendDays) <= 0)

                                Day = BackDatesDate.AddDays(CntDayHusk - CntDay2 - 1).DayOfWeek.ToString().Substring(0, 3)
                                If Day = "Sat" Or Day = "Sun" Then

                                    WeekendDays = WeekendDays + 1

                                End If
                                CntDay = CntDay - 1
                                CntDay2 = CntDay2 + 1

                                'you have to check day equal to "sat" Or "sun".
                            Loop

                        End If

                        BackDatesDate = BackDatesDate.AddDays(WeekendDays)



                        sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
                        sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = " + pl.VisResCode + "  ) "
                        sql = sql + "And  U_Date BETWEEN '" + workDate.ToString("MM-dd-yyyy") + "' and '" + workDate.ToString("MM-dd-yyyy") + "' and owor.Status <> 'C'"

                        ' sql = sql + "And  isnull(U_ResType,'" + OworRCS_PT + "') = '" + OworRCS_PT + "'"
                        sql = sql + "And  isnull(U_ResType,'') = '" + OworRCS_PT + "'"

                        rs4.DoQuery(sql)
                        ComCap = rs4.Fields.Item("ComCap").Value



                        If OworRCS_PT = "SER" Then

                            If Cap > ((STP * 60) * StpSerDateSpan) Then

                                If workDate <= BackDatesDate Then
                                    CapStpSer = Cap
                                    HPD = Cap / 60
                                Else
                                    CapStpSer = Cap - ((STP * 60) * StpSerDateSpan)
                                End If


                            Else
                                ' CapStpSer = Cap - (STP * 60) - ComCap
                                CapStpSer = Cap
                            End If

                        Else
                            If Cap > ((STP * 60) * StpSerDateSpan) Then
                                CapStpSer = (STP * 60) * StpSerDateSpan
                            Else
                                'CapStpSer = Cap - (SEP * 60) - ComCap
                                CapStpSer = Cap
                            End If

                        End If




                        sql = "Select isnull(sum(isnull([U_Capacity],0)),0) as ComCap FROM [@RCS_CAP_COMMIT] inner join owor on U_DocEntry=DocEntry  "
                        sql = sql + "Where U_ResGrpCod = (Select ORSC.ResGrpCod from ORSC where ORSC.ResCode = " + pl.VisResCode + "  ) "
                        sql = sql + "And  U_Date BETWEEN '" + workDate.ToString("MM-dd-yyyy") + "' and '" + workDate.ToString("MM-dd-yyyy") + "' and owor.Status <> 'C'"
                        rs4.DoQuery(sql)
                        ComCapTotal = rs4.Fields.Item("ComCap").Value

                        ' FreeCap = Cap - ComCap
                        FreeCap = CapStpSer - ComCap

                        FreeCapTotal = (Cap - ComCapTotal)


                        If FreeCap > 0 And FreeCapTotal > 0 Then
                            If (FreeCap > (HPD * 60)) And (FreeCapTotal > (HPD * 60)) Then

                                If RememberRestQty > (HPD * 60) Then
                                    RememberRestQty = RememberRestQty - (HPD * 60)
                                    pl.CapacityPerDay = (HPD * 60)
                                Else
                                    pl.CapacityPerDay = RememberRestQty
                                    RememberRestQty = 0
                                End If
                            Else
                                If (RememberRestQty > FreeCap) And (FreeCap < FreeCapTotal) Then
                                    RememberRestQty = RememberRestQty - FreeCap
                                    pl.CapacityPerDay = FreeCap
                                Else
                                    If (RememberRestQty > FreeCapTotal) Then
                                        RememberRestQty = RememberRestQty - FreeCapTotal
                                        pl.CapacityPerDay = FreeCapTotal

                                    Else
                                        pl.CapacityPerDay = RememberRestQty
                                        RememberRestQty = 0
                                    End If

                                End If
                            End If
                            If pl.CapacityPerDay > 0 Then
                                sbInsert.Append(String.Format(insert, code, pl.VisResCode, pl.DocEntry, pl.LineNum, pl.CapacityPerDay.ToDBString(), workDate.ToDBString(), pl.ResGrpCod.ToString(), pl.ObjType, OworRCS_PT)).AppendLine()
                            End If

                        End If
                        cnt = cnt + 1
                        If (cnt >= pl.WorkingDates.Count) Then
                            If RememberRestQty > 0 Then
                                pl.CapacityPerDay = RememberRestQty
                                RememberRestQty = 0
                                If pl.CapacityPerDay > 0 Then
                                    sbInsert.Append(String.Format(insert, code, pl.VisResCode, pl.DocEntry, pl.LineNum, pl.CapacityPerDay.ToDBString(), workDate.ToDBString(), pl.ResGrpCod.ToString(), pl.ObjType, OworRCS_PT)).AppendLine()
                                End If

                            End If
                        End If
                        'rs 2021-11-09
                        ' rs.DoQuery(sbInsert.ToString())
                        ' sbInsert.Clear()
                    Next
                    cnt = 0

                    rs.DoQuery(sbInsert.ToString())
                Else
                    sbInsert.Append(String.Format(insert, code, pl.VisResCode, pl.DocEntry, pl.LineNum, pl.RestQty.ToDBString(), pl.StartDate.ToDBString(), pl.ResGrpCod, pl.ObjType, OworRCS_PT)).AppendLine()
                    rs.DoQuery(sbInsert.ToString())
                End If
            Else
                rs.DoQuery(sbInsert.ToString())
            End If
            'Next
        Next i
    End Sub

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

End Module