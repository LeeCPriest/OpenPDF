Imports EdmLib
Imports SldWorks
Imports SwConst
Imports System.Reflection
Public Class PDMTools
    Implements IEdmAddIn5

    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo
        poInfo.mbsAddInName = "TruSteel_PDMTools"
        poInfo.mbsCompany = "Written by Lee Priest - leeclarkepriest@gmail.com"

        'Specify information to display in the add-in's Properties dialog box
        poInfo.mbsDescription = "Use 'Name of add-in' field in datacard button properties as follows:
                                1. To open PDF files - 'OpenPDF:[Variable_PDF_Path]'
                                2. To add new configuration to assy/part files - 'AddCfgPN'"
        poInfo.mlAddInVersion = 1.1
        poInfo.mlRequiredVersionMajor = 8
        poInfo.mlRequiredVersionMinor = 0

        'Notify the add-in when a file data card button is clicked
        poCmdMgr.AddHook(EdmCmdType.EdmCmd_CardButton)

    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData As Array) Implements IEdmAddIn5.OnCmd
        'To trigger, the field called 'Name of add-in' associated with the PDM datacard button must be of the following format: OpenPDF:[Variable_PDF_Path]

        If poCmd.meCmdType = EdmCmdType.EdmCmd_CardButton Then
            Dim eCmdData As EdmCmdData = ppoData(0) 'get the first entry of the cmd data array (should only be one entry, since triggered by button on datacard)

            'get the PDF path from the datacard variable specified
            Dim eVault As EdmVault5 = poCmd.mpoVault 'get the vault object
            Dim eFile As IEdmFile5 = eVault.GetObject(EdmObjectType.EdmObject_File, eCmdData.mlObjectID1) 'get the file object
            Dim eFileCard As IEdmEnumeratorVariable8 = eFile.GetEnumeratorVariable() 'get the datacard object

            If UCase(Left(poCmd.mbsComment, 8)) = "OPENPDF:" Then
                OpenPDF(poCmd, eVault, eFile, eFileCard)
            ElseIf UCase(Left(poCmd.mbsComment, 8)) = "ADDCFGPN" Then
                AddConfigPN(poCmd, eVault, eFile, eFileCard)
            End If
        End If

    End Sub

    Private Sub OpenPDF(ByVal poCmd As EdmCmd, ByVal eVault As EdmVault5, ByVal eFile As IEdmFile5, ByVal eFileCard As IEdmEnumeratorVariable8)

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
            ConfigSelectForm.InfoText.Text = "Select the configuration to open the pdf:"
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
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, MethodInfo.GetCurrentMethod.ToString)
        End Try

        eFileCard.CloseFile(False) 'close the datacard object

        If IO.File.Exists(PDFPath) = True Then 'check if the PDF exists

            Try
                Process.Start(PDFPath) 'open the selected PDF
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MethodInfo.GetCurrentMethod.ToString)
            End Try

        Else
            MsgBox("The file path '" & PDFPath & "' does not exist. Unable to open.", MsgBoxStyle.Exclamation, MethodInfo.GetCurrentMethod.ToString)
        End If

    End Sub

    Private Sub AddConfigPN(ByVal poCmd As EdmCmd, ByVal eVault As EdmVault5, ByVal eFile As IEdmFile5, ByVal eFileCard As IEdmEnumeratorVariable8)

        Try
            Dim eUserID As Integer = eVault.GetLoggedInWindowsUserID(eVault.Name)
            Dim eUserMgr As IEdmUserMgr7 = eVault

            Dim swFileName As String = eFile.Name
            Dim swFileType As swDocumentTypes_e
            If UCase(Right(swFileName, 6)) = "SLDASM" Then
                swFileType = swDocumentTypes_e.swDocASSEMBLY
            ElseIf UCase(Right(swFileName, 6)) = "SLDDRW" Then
                swFileType = swDocumentTypes_e.swDocDRAWING
            ElseIf UCase(Right(swFileName, 6)) = "SLDPRT" Then
                swFileType = swDocumentTypes_e.swDocPART
            End If

            If swFileType <> swDocumentTypes_e.swDocDRAWING Then
                'check if the file is checked out
                If eFile.IsLocked = True And eFile.LockedByUserID = eUserID Then
                    'get active file in SolidWorks
                    Dim SWApp As SldWorks.SldWorks = Nothing
                    Try
                        SWApp = Runtime.InteropServices.Marshal.GetActiveObject("SldWorks.Application") 'if SolidWorks is already running, attach to that process
                    Catch
                    End Try

                    If SWApp Is Nothing Then
                        MsgBox("SolidWorks must be opened to add a new configuration.", MsgBoxStyle.Exclamation, MethodInfo.GetCurrentMethod.ToString)

                        '*** This attempt to open SolidWorks fails because the add-ins don't load, for some reason ***
                        'If MsgBox("SolidWorks must be opened to add a new configuration. Continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, My.Application.Info.AssemblyName) = vbYes Then
                        '    Dim swAppType = System.Type.GetTypeFromProgID("SldWorks.Application")
                        '    SWApp = TryCast(System.Activator.CreateInstance(swAppType), SldWorks.SldWorks)
                        '    SWApp.Visible = True
                        '    SWApp.UserControl = True
                        'End If
                    End If

                    If SWApp IsNot Nothing Then

                        Dim SWFile As ModelDoc2 = SWApp.ActiveDoc

                        If SWFile Is Nothing Then
                            MsgBox("The file " & swFileName & " must be opened to add a new configuration.", MsgBoxStyle.Exclamation, MethodInfo.GetCurrentMethod.ToString)

                            '*** This attempt to open the file fails because the PDM add-in isn't updatating, for some reason ***
                            'If MsgBox("The file " & swFileName & " must be opened to add a new configuration. Continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, My.Application.Info.AssemblyName) = vbYes Then
                            '    SWApp.OpenDoc(eFile.LockPath, swFileType)
                            'End If
                        End If

                        If SWFile IsNot Nothing Then

                            If SWFile.GetPathName = eFile.LockPath Then 'if not the file associated with the datacard, offer to open it (ADD THIS FUNCTION)

                                'get the configuration information from the file
                                Dim ConfigNames() As String = SWFile.GetConfigurationNames
                                Dim ConfigList As New ArrayList(ConfigNames)

                                Dim SelectedConfig As String = ""

                                'ask user to select config to copy from
                                If ConfigList.Count = 1 Then 'if there is 1 configuration, return that
                                    SelectedConfig = ConfigList.Item(0)
                                ElseIf ConfigList.Count > 1 Then 'if there is more than 1 configuration, open the form to select, otherwise default to the 
                                    Dim ConfigSelectForm As New ConfigSelect
                                    ConfigSelectForm.ConfigList = ConfigList
                                    ConfigSelectForm.InfoText.Text = "Select the configuration to copy:"
                                    ConfigSelectForm.ShowDialog()
                                    SelectedConfig = ConfigSelectForm.ConfigListBox.Items(ConfigSelectForm.ConfigListBox.SelectedIndex)
                                    ConfigSelectForm.Dispose()
                                End If

                                If SelectedConfig <> "" Then
                                    'ask user for configuration name
                                    Dim cfgExists As Boolean = True
                                    Dim newCfgName As String = ""
                                    Do While cfgExists = True
                                        newCfgName = InputBox("Please provide the name of the new configuration to add", MethodInfo.GetCurrentMethod.ToString, " ")

                                        If newCfgName = "" Then
                                            Exit Do
                                        ElseIf newCfgName <> " " Then
                                            cfgExists = False

                                            For i = 0 To ConfigList.Count - 1
                                                If ConfigList.Item(i) = newCfgName Then
                                                    MsgBox("A configuration named " & newCfgName & " already exists. Please enter a new name.", MsgBoxStyle.Exclamation, MethodInfo.GetCurrentMethod.ToString)
                                                    cfgExists = True
                                                End If
                                            Next
                                        Else
                                            MsgBox("A blank configuration name is not valid. Please enter a new name.", MsgBoxStyle.Exclamation, MethodInfo.GetCurrentMethod.ToString)
                                        End If
                                    Loop

                                    If newCfgName <> "" Then 'if the user hasn't cancelled
                                        ' create config
                                        Dim ConfigMgr As ConfigurationManager = SWFile.ConfigurationManager
                                        SWFile.ShowConfiguration2(SelectedConfig)
                                        ConfigMgr.AddConfiguration2(newCfgName, "1", "1", 0, "", "1", True)

                                        ' get the next serial #
                                        Dim eSerialGen As IEdmSerNoGen7 = eVault.CreateUtility(EdmUtility.EdmUtil_SerNoGen)
                                        Dim SerialNum As IEdmSerNoValue = eSerialGen.AllocSerNoValue("Purchase Parts SN")

                                        ' set serial # custom property
                                        Dim swCustPropMgr As CustomPropertyManager = SWFile.Extension.CustomPropertyManager(newCfgName)

                                        Try
                                            swCustPropMgr.Set("CAD_Manufacturer", "") 'clear CAD_Manufacturer custom property
                                        Catch
                                        End Try

                                        Try
                                            swCustPropMgr.Set("CAD_ManufacturerPartNumber", "")  'clear CAD_ManufacturerPartNumber custom property
                                        Catch
                                        End Try

                                        Try
                                            swCustPropMgr.Add3("PDM_DocumentNumber", swCustomInfoType_e.swCustomInfoText, SerialNum.Value, True)
                                        Catch ex As Exception
                                            MsgBox(ex.Message, MsgBoxStyle.Critical, My.Application.Info.AssemblyName)
                                        Finally
                                            MsgBox("The configuration " & newCfgName & " has been added.", MsgBoxStyle.Information, MethodInfo.GetCurrentMethod.ToString)
                                        End Try
                                    End If

                                End If
                            Else
                                MsgBox("The file " & swFileName & " is not open in SolidWorks.", MsgBoxStyle.Exclamation, MethodInfo.GetCurrentMethod.ToString)
                            End If
                        End If
                    End If
                Else
                    MsgBox("The file must be checked out to you in order to complete this operation", MsgBoxStyle.Exclamation, MethodInfo.GetCurrentMethod.ToString)
                End If
            Else
                MsgBox("This action can only be performed on parts and assemblies.", MsgBoxStyle.Exclamation, MethodInfo.GetCurrentMethod.ToString)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, MethodInfo.GetCurrentMethod.ToString)
        End Try

    End Sub

End Class
