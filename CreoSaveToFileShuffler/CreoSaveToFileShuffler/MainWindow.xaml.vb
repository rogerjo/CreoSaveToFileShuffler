Imports pfcls

Class MainWindow

    Dim asyncConnection As IpfcAsyncConnection = Nothing
    Dim model As IpfcModel
    Dim activeserver As IpfcServer
    Dim paramval As IpfcParamValue
    Dim session As IpfcBaseSession
    Dim Moditem As CMpfcModelItem
    Dim State As String = ""
    Dim FileEnd As String = ""
    Dim ConvertType As Boolean

    Dim FileNameComplete As String

    Sub Creo_Connect()

        Dim asyncConnection As IpfcAsyncConnection = Nothing

        Try
            myInfo.Text = "Connecting..."

            asyncConnection = (New CCpfcAsyncConnection).Connect(Nothing, Nothing, Nothing, Nothing)
            session = asyncConnection.Session
            activeserver = session.GetActiveServer
            model = session.CurrentModel
            myInfo.Text = "Connection established"

        Catch ex As Exception
            MsgBox(ex.Message.ToString + Chr(13) + ex.StackTrace.ToString)
            If Not asyncConnection Is Nothing AndAlso asyncConnection.IsRunning Then
                asyncConnection.Disconnect(1)
            End If
            myInfo.Text = "Error occurred while connecting"

        End Try
    End Sub
    Private Sub MyWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles myWindow.Loaded
        myInfo.Text = "Working..."
        Call Creo_Connect()

        Call SaveObjectToDisk()
    End Sub

    Private Sub SaveObjectToDisk()

        'Try
        '    asyncConnection = (New CCpfcAsyncConnection).Connect(Nothing, Nothing, Nothing, Nothing)
        'session = asyncConnection.Session
        'activeserver = session.GetActiveServer
        'model = session.CurrentModel
        'Catch ex As Exception
        '    myInfo.Text = "No session"
        'End Try

        Try
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

                Case 1
                    FileEnd = ".stp"

                Case 2
                    FileEnd = ".pdf"

                Case Else
                    MsgBox("Model not supported. Only Drawings, Models or Assemblies", "Script Message")
                    asyncConnection.Disconnect(1)
                    Environment.Exit(0)
            End Select

            FileNameComplete = model.FullName + "_" + model.Revision + "_" + model.Version + "_" + State + FileEnd

            Call ExportFileToDisc(FileNameComplete, model.Type)
        Catch ex As Exception

        End Try

    End Sub

    Sub TestForDir(workdir As String)
        Try
            If Dir(workdir, vbDirectory) = "" Then
                MkDir(workdir)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ExportFileToDisc(FileNameComplete As String, ConvertType As Integer)
        Dim Workdir As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString() & "\Fileshuffler Files\"
        Dim Destination As String = Workdir & FileNameComplete

        TestForDir(Workdir)

        myInfo.Text = FileNameComplete.ToString()


        Try
            If (ConvertType = 0 Or ConvertType = 1) Then 'Export assy and model to STEP
                Dim cDesExStep As CCpfcSTEP3DExportInstructions
                Dim DesFlags As IpfcGeometryFlags
                Dim Des3DEx As IpfcExport3DInstructions
                Dim DesEx As IpfcExportInstructions
                Dim DesExStep As IpfcSTEP3DExportInstructions

                cDesExStep = New CCpfcSTEP3DExportInstructions
                DesFlags = (New CCpfcGeometryFlags).Create()
                DesFlags.AsSolids = True
                DesExStep = cDesExStep.Create(EpfcAssemblyConfiguration.EpfcEXPORT_ASM_FLAT_FILE, DesFlags)
                Des3DEx = DesExStep
                DesEx = Des3DEx

                session.CurrentModel.Export(Destination, Des3DEx)


            ElseIf (ConvertType = 2) Then 'Export drawing to PDF
                Dim expdf As IpfcPDFExportInstructions
                Dim pdfopt As IpfcPDFOption
                Dim EpfcPDFOPT_LAUNCH_VIEWER As Boolean
                Dim Drawing As IpfcModel2D
                Dim sheet As IpfcSheetOwner

                Drawing = CType(session.CurrentModel, IpfcModel2D)
                Drawing.Regenerate()


                EpfcPDFOPT_LAUNCH_VIEWER = True
                expdf = (New CCpfcPDFExportInstructions).Create()
                pdfopt = (New CCpfcPDFOption).Create()
                pdfopt.OptionValue = (New CMpfcArgument).CreateBoolArgValue(EpfcPDFOPT_LAUNCH_VIEWER)

                session.CurrentModel.Export(Destination, CType(expdf, IpfcExportInstructions))

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MyButton_Click(sender As Object, e As RoutedEventArgs) Handles myButton.Click
        Try
            asyncConnection.Disconnect(1)
        Catch ex As Exception

        End Try
        Me.Close()
    End Sub
End Class
