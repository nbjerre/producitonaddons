Option Strict Off
Option Explicit On
Friend Class ChangeSysForm

    Public Property PlanUnlimetedName As String
    Public Property PlanUnlimeted As Boolean
    Public Property SalesOrderDocEntry As Long
    Public Property LineNum As Integer
    Public Property ProdDocNum As Integer
    Public Property UserName As String


    Private Sub Class_Initialize_Renamed()

        'Dim planUnlimeted As Boolean
        'Dim salesOrderDocEntry As Long
        Dim path As String = Windows.Forms.Application.StartupPath()

        Dim result As Response = New Response
        Dim preText As String = "\PLAN_UNLIMITED_"

        'Show message in Console
        Console.WriteLine($"{Me.PlanUnlimetedName} er startet")
        Console.WriteLine()

        If LineNum >= 0 Then
            result = AIG_KGJT_PlanUnlimited.PlanAllProductionOrdersPerLine(PlanUnlimeted, SalesOrderDocEntry, LineNum)
            'preText = "\PLAN_UNLIMITED_LINE_"
        ElseIf ProdDocNum > 0 Then
            result = AIG_KGJT_PlanUnlimited.UpdateProductionOrderLineStartEndDateUnLimited(PlanUnlimeted, ProdDocNum)
            preText = "\PLAN_UNLIMITED_PROD_"
        Else
            result = AIG_KGJT_PlanUnlimited.PlanAllProductionOrders(PlanUnlimeted, SalesOrderDocEntry)
        End If


        If result.Success Then
            path = path + preText + Me.UserName + "_OK.txt"

            Try
                If File.Exists(path) Then
                    File.Delete(path)
                End If
                File.Create(path)
            Catch ex As Exception
                'Console.WriteLine($"Kan ikke oprette {preText + Me.UserName}_OK.txt filen")
            End Try

            'Show message in Console
            Console.WriteLine()
            Console.WriteLine($"{Me.PlanUnlimetedName} er fuldført")
            'Console.ReadKey()

        Else
            path = path + preText + Me.UserName + "_FAILED.txt"
            If File.Exists(path) Then
                'Dim file As StreamWriter
                'file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
                'file.WriteLine()
                'file.WriteLine(result.Message)
                'file.Close()
                Using file As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(path, True)
                    file.WriteLine()
                    file.WriteLine(result.Message)
                End Using
            Else
                'Dim fs As FileStream
                'fs = File.Create(path)
                'Dim info As Byte() = New Text.UTF8Encoding(True).GetBytes(result.Message)
                'fs.Write(info, 0, info.Length)
                'fs.Close()
                Using fs As FileStream = File.Create(path)
                    Dim info As Byte() = New Text.UTF8Encoding(True).GetBytes(result.Message)
                    fs.Write(info, 0, info.Length)
                End Using

            End If

            'Show message in Console
            Console.WriteLine()
            Console.WriteLine($"{Me.PlanUnlimetedName} fejlet. FEJL: " & result.Message)
            'Console.ReadKey()

        End If

    End Sub

    'Public Sub New()
    '    MyBase.New()
    '    Class_Initialize_Renamed()
    'End Sub

    Public Sub New(btnItemUID As String, salesOrderDocEntry As Long, lineNum As Integer, prodDocNum As Integer, userName As String)
        MyBase.New()
        Me.PlanUnlimetedName = GetPlanUnlimitedName(btnItemUID)
        Me.PlanUnlimeted = GetPlanUnlimitedStatus(btnItemUID)
        Me.SalesOrderDocEntry = salesOrderDocEntry
        Me.ProdDocNum = prodDocNum
        Me.LineNum = lineNum
        Me.UserName = userName

        Class_Initialize_Renamed()
    End Sub

    Private Function GetPlanUnlimitedName(ByRef btnItemUID As String) As String

        Dim name As String

        Select Case btnItemUID.ToLower()
            Case "btnpu"
                name = "Planlæg Ubegransæt i Kundeordre"
            Case "btnpul"
                name = "Planlæg Linje i Kundeordre"
            Case "btnpr"
                name = "Planlæg Alle i Kundeordre"
            Case "btnpl"
                name = "Planlæg Rute datoer i Produktionsordre"
            Case "btnplu"
                name = "Planlæg Ubegrænset i Produktionsordre"
            Case Else
                name = ""
        End Select

        GetPlanUnlimitedName = name
    End Function

    Private Function GetPlanUnlimitedStatus(ByRef btnItemUID As String) As Boolean

        Dim pu As Boolean
        Select Case btnItemUID.ToLower()
            'Case "btnpu" 'Planlæg Ubegransæt (Kundeordre)
            '    pu = True
            'Case "btnpul" 'Planlæg Linje (Kundeordre)
            '    pu = False
            Case "btnpr" 'Planlæg Alle (Kundeordre)
                pu = False
            Case "btnpl" 'Planlæg Rute datoer (Produktionsordre)
                pu = False
                'Case "btnplu" 'Planlæg Ubegrænset (Produktionsordre)
                '    pu = True
            Case Else
                pu = True
        End Select

        GetPlanUnlimitedStatus = pu
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class