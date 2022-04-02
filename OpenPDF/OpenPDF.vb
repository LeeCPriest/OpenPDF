Imports EdmLib

Public Class OpenPDF
    Implements IEdmAddIn5

    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo
        poInfo.mbsAddInName = "OpenPDF"
        poInfo.mbsCompany = "Written by Lee Priest - leeclarkepriest@gmail.com"

        'Specify information to display in the add-in's Properties dialog box
        poInfo.mbsDescription = "Opens PDF files using PDM datacard button." & vbCrLf & "Use 'Name of add-in' field to specify variable to read." & vbCrLf & "Syntax 'OpenPDF:[Variable_PDF_Path]'"
        poInfo.mlAddInVersion = 1
        poInfo.mlRequiredVersionMajor = 8
        poInfo.mlRequiredVersionMinor = 0

        'Notify the add-in when a file data card button is clicked
        poCmdMgr.AddHook(EdmCmdType.EdmCmd_CardButton)

    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData As Array) Implements IEdmAddIn5.OnCmd
        'To trigger, the field called 'Name of add-in' associated with the PDM datacard button must be of the following format: OpenPDF:[Variable_PDF_Path]

        If UCase(Left(poCmd.mbsComment, 8)) = "OPENPDF:" Then
            Dim eCmdData As EdmCmdData = ppoData(0) 'get the first entry of the cmd data array (should only be one entry, since triggered by button on datacard)

            'get the PDF path from the datacard variable specified
            Dim eVault As EdmVault5 = poCmd.mpoVault 'get the vault object
            Dim eFile As IEdmFile5 = eVault.GetObject(EdmObjectType.EdmObject_File, eCmdData.mlObjectID1) 'get the file object
            Dim eFileCard As IEdmEnumeratorVariable8 = eFile.GetEnumeratorVariable() 'get the datacard object

            'get the configuration information from the file
            Dim eFileConfigs As EdmStrLst5 = eFile.GetConfigurations()
            Dim eConfigName As String
            Dim ConfigList As New ArrayList

            Dim pos As IEdmPos5 = eFileConfigs.GetHeadPosition
            While Not pos.IsNull
                eConfigName = eFileConfigs.GetNext(pos)
                ConfigList.Add(eConfigName)
            End While

            Dim SelectedConfig As String = ""

            If ConfigList.Count = 2 Then 'if there are 2 configurations (@ and only one model config), select the model config 
                SelectedConfig = ConfigList.Item(1)
            ElseIf ConfigList.Count > 2 Then 'if there are more than 2 configurations, open the form to select, otherwise default to the 
                Dim ConfigSelectForm As New ConfigSelect
                ConfigSelectForm.ConfigList = ConfigList
                ConfigSelectForm.ShowDialog()
                SelectedConfig = ConfigSelectForm.ConfigListBox.Items(ConfigSelectForm.ConfigListBox.SelectedIndex)
                ConfigSelectForm.Dispose()
            End If

            Dim PDFPath As String = ""

            'Get the name of the variable from which to read the PDF path
            Dim VarName As String
            VarName = Right(poCmd.mbsComment, Len(poCmd.mbsComment) - 8)

            Try
                eFileCard.GetVar(VarName, SelectedConfig, PDFPath) 'get the PDF path from the specified variable, for the selected config
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "OpenPDF Error")
            End Try

            eFileCard.CloseFile(False) 'close the datacard object

            If IO.File.Exists(PDFPath) = True Then 'check if the PDF exists

                Try
                    Process.Start(PDFPath) 'open the selected PDF
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation, "OpenPDF Error")
                End Try

            Else
                MsgBox("The file path '" & PDFPath & "' does not exist. Unable to open.", MsgBoxStyle.Exclamation, "OpenPDF Error")
            End If


        End If

    End Sub

End Class
