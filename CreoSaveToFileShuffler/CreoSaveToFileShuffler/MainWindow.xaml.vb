Imports pfcls

Class MainWindow

    Dim asyncConnection As IpfcAsyncConnection = Nothing
    Dim model As IpfcModel
    Dim activeserver As IpfcServer
    Dim paramown As IpfcParameterOwner
    Dim ipbaseparam As IpfcBaseParameter
    Dim paramval As IpfcParamValue
    Dim session As IpfcBaseSession
    Dim Moditem As CMpfcModelItem
    Dim State As String = ""
    Dim FileEnd As String = ""
    Dim ConvertType As Boolean

    Dim FileNameComplete As String


    Private Sub myWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles myWindow.Loaded
        myInfo.Text = "*****"
        Call SaveObjectToDisk()
    End Sub

    Private Sub SaveObjectToDisk()

        Try
            asyncConnection = (New CCpfcAsyncConnection).Connect(Nothing, Nothing, Nothing, Nothing)
            session = asyncConnection.Session
            activeserver = session.GetActiveServer
            model = session.CurrentModel

            If model Is Nothing Then
                MsgBox("Model is not present",, "Script message")
                asyncConnection.Disconnect(1)
                Environment.Exit(0)
            End If

            If activeserver.IsObjectCheckedOut(activeserver.ActiveWorkspace, model.FileName) Then
                MsgBox("Please check in model first...",, "Script Message")
                asyncConnection.Disconnect(1)
                Environment.Exit(0)
            End If

            Select Case model.ReleaseLevel
                Case "Concept"
                    State = "C"
                Case "Design"
                    State = "D"
                Case "Pre-Released"
                    State = "P"
                Case "Released"
                    State = "R"
                Case Else
            End Select

            Select Case model.Type
                Case 0
                    FileEnd = ".stp"
                    ConvertType = True
                Case 1
                    FileEnd = ".stp"
                    ConvertType = True
                Case 2
                    FileEnd = ".pdf"
                    ConvertType = False
                Case Else
                    MsgBox("Model not supported. Only Drawings, Models or Assemblies", "Script Message")
                    asyncConnection.Disconnect(1)
                    Environment.Exit(0)
            End Select

            FileNameComplete = model.FullName + "_" + model.Revision + "_" + model.Version + "_" + State + FileEnd

            Call ExportFileToDisc(FileNameComplete, ConvertType)

            asyncConnection.Disconnect(1)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub ExportFileToDisc(FileNameComplete As String, ConvertType As Boolean)
        myInfo.Text = FileNameComplete.ToString()
    End Sub
End Class
