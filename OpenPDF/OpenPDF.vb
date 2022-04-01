Imports EdmLib

Public Class OpenPDF
    Implements IEdmAddIn5

    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo
        poInfo.mbsAddInName = "OpenPDF"
        poInfo.mbsCompany = "Written by Lee Priest - leeclarkepriest@gmail.com"

        'Specify information to display in the add-in's Properties dialog box
        poInfo.mbsDescription = "Opens PDF files using PDM datacard button." & vbCrLf & "Use Name add-in name field to specify variable to read." & vbCrLf & "Syntax 'OpenPDF:[Variable_PDF_Path]'"
        poInfo.mlAddInVersion = 1
        poInfo.mlRequiredVersionMajor = 8
        poInfo.mlRequiredVersionMinor = 0

        'Notify the add-in when a file data card button is clicked
        poCmdMgr.AddHook(EdmCmdType.EdmCmd_CardButton)

    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData As Array) Implements IEdmAddIn5.OnCmd
        'To trigger the field called 'Name of add-in' associated with the PDM datacard button must be of the following format: OpenPDF:[Variable_PDF_Path]

        If UCase(Left(poCmd.mbsComment, 8)) = "OPENPDF:" Then
            Dim eCmdData As EdmCmdData = ppoData(0) 'get the first entry of the cmd data array (should only be one entry, since triggered by button on datacard)

            'Get the name of the variable to update.
            Dim VarName As String
            VarName = Right(poCmd.mbsComment, Len(poCmd.mbsComment) - 8)

            Dim PDFPath As String = ""

            'get the PDF path from the datacard variable specified
            Dim eVault As EdmVault5 = poCmd.mpoVault 'get the vault object
            Dim eFile As IEdmFile5 = eVault.GetObject(EdmObjectType.EdmObject_File, eCmdData.mlObjectID1) 'get the file object
            Dim eFileCard As IEdmEnumeratorVariable8 = eFile.GetEnumeratorVariable() 'get the datacard object
            eFileCard.GetVarFromDb(VarName, "@", PDFPath) 'read the PDF path from the specified variable

            'define the process info used to open the PDF file
            Dim ProcStartInfo As New ProcessStartInfo()
            ProcStartInfo.FileName = "powershell.exe"
            ProcStartInfo.Arguments = "Invoke-Item" & " " & PDFPath

            Try
                Process.Start(ProcStartInfo)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "OpenPDF Error")
            End Try
        End If
    End Sub

End Class
