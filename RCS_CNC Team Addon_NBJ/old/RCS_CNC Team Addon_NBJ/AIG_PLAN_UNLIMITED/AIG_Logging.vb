Module AIG_Logging

    Public Sub ErrLog(ByVal SubMetode As String, ByRef exc As Exception)
        Dim Path As String
        Dim objwriter As System.IO.StreamWriter = Nothing
        Dim TimeStamp As String
        Dim MessageString As String
        Dim username As String = ""
        Try
            Try
                username = vCmp.UserName
            Catch ex As Exception
                username = "NoName"
            End Try

            Path = (System.Windows.Forms.Application.StartupPath).ToString '+ "\LogFiles"
            If Not IO.Directory.Exists(Path) Then
                IO.Directory.CreateDirectory(Path)
            End If
            Path = Path + "\AIG_ERRORLOG-" + username + ".log"

            TimeStamp = Now + ":" + Right("00" + Now.Second.ToString, 2) + "," + Now.Millisecond.ToString

            If Not exc Is Nothing Then
                MessageString = TimeStamp + vbCrLf + "Sub: " + SubMetode + " Msg: " + exc.Message + vbCrLf + "Trace: " + exc.StackTrace + vbCrLf
            Else
                MessageString = SubMetode
            End If


            objwriter = New System.IO.StreamWriter(Path, True)
            objwriter.Write(MessageString)
            objwriter.WriteLine()
            objwriter.Flush()

        Catch ex As Exception

        Finally
            If Not objwriter Is Nothing Then
                objwriter.Close()
                objwriter.Dispose()
            End If
        End Try
    End Sub


    Public Sub PrintToErrorLog(ByVal SubMetode As String, ByVal ErrorMsg As String, Optional ByVal UserName As String = "")
        Dim Path As String
        Dim objwriter As System.IO.StreamWriter
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
                objwriter.Dispose()
            End If
        End Try
    End Sub

    Public Sub PrintToDebugLog(ByVal Msg As String, Optional ByVal UserName As String = "")
        Dim path As String
        Dim objwriter As System.IO.StreamWriter
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
            '            objwriter.Close()

        Catch ex As Exception
        Finally
            If Not objwriter Is Nothing Then
                objwriter.Close()
                objwriter.Dispose()
            End If
        End Try
    End Sub

End Module
