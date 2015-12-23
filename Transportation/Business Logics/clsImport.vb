Imports System.IO
Imports System.Diagnostics.Process
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Net
Imports System.Xml
Imports Microsoft.VisualBasic
Imports System
Imports System.Threading

Public Class clsImport
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oStaticText As SAPbouiCOM.StaticText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private oFolder As SAPbouiCOM.Folder
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oNewItem, oItem1 As SAPbouiCOM.Item
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Dim strFileName As String
    Dim strSelectedFilepath, sPath, strSelectedFolderPath As String
    Dim dtDatatable As SAPbouiCOM.DataTable
    Dim blnErrorflag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Methods"
    Private Sub LoadForm()
        oApplication.Utilities.LoadForm(xml_Import, frm_Import)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.DataSources.UserDataSources.Add("path", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        'oCombobox = oForm.Items.Item("4").Specific
        'oCombobox.ValidValues.Add("UDT", "Import All files")
        ''oCombobox.ValidValues.Add("SKU", "SKU")
        '' oCombobox.ValidValues.Add("BP", "Business Partner")
        'oCombobox.ValidValues.Add("SHP", "Invoice Import")
        'oCombobox.ValidValues.Add("ASN", "Receipt Import")
        ''oCombobox.ValidValues.Add("ADJ", "Adjustment Import")
        ''oCombobox.ValidValues.Add("HOLD", "Hold Import")
        'oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        ' oForm.Items.Item("4").Enabled = False
        ' oForm.Items.Item("4").DisplayDesc = True

        oForm.DataSources.UserDataSources.Add("fld1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld4", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld5", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        oForm.DataSources.UserDataSources.Add("fld6", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld7", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld8", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld9", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        oForm.DataSources.UserDataSources.Add("fld11", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld12", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld13", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld14", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        oForm.DataSources.UserDataSources.Add("fld15", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld16", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld17", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld18", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        oForm.DataSources.UserDataSources.Add("fld19", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld20", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld21", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("fld22", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        oFolder = oForm.Items.Item("10").Specific
        oFolder.DataBind.SetBound(True, "", "fld1")

        oFolder = oForm.Items.Item("11").Specific
        oFolder.DataBind.SetBound(True, "", "fld22")
        oFolder = oForm.Items.Item("13").Specific
        oFolder.DataBind.SetBound(True, "", "fld21")
        oForm.DataSources.UserDataSources.Add("fld32", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        oFolder = oForm.Items.Item("1000002").Specific
        oFolder.DataBind.SetBound(True, "", "fld32")

        oEditText = oForm.Items.Item("6").Specific
        oEditText.DataBind.SetBound(True, "", "path")
        oForm.Freeze(True)
        AddSubFolders(oForm)
        AddSubFolders1(oForm)
        AddSubFolders2(oForm)
        AddSubFolders3(oForm)
        oForm.Freeze(False)
        oForm.PaneLevel = 1
    End Sub

#Region "AddSubFolders"
    Private Sub AddSubFolders(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            For i As Integer = 1 To 6
                oNewItem = aForm.Items.Add("Folder" & i, SAPbouiCOM.BoFormItemTypes.it_FOLDER)
                If i = 1 Then
                    oItem1 = aForm.Items.Item("10")
                    oNewItem.Left = oItem1.Left + (i - 1) + 5
                    oNewItem.Top = oItem1.Top + 30
                    oNewItem.Width = oItem1.Width + 20
                Else
                    oItem1 = aForm.Items.Item("Folder" & i - 1)
                    oNewItem.Left = oItem1.Left + oItem1.Width + 1
                    oNewItem.Top = oItem1.Top
                    oNewItem.Width = oItem1.Width + 20
                End If
                'oNewItem.Width = oItem1.Width + 10
                oNewItem.Height = 20
                oNewItem.FromPane = 2
                oNewItem.ToPane = 7
                Dim oFolder As SAPbouiCOM.Folder
                oFolder = oNewItem.Specific
                If i = 1 Then
                    oFolder.Caption = "Goods In/Delivery Notes"
                    oFolder.DataBind.SetBound(True, "", "fld4")
                ElseIf i = 2 Then
                    oFolder.Caption = "Goods In/Delivery Notes/Costs"
                    oFolder.DataBind.SetBound(True, "", "fld5")
                ElseIf i = 3 Then
                    oFolder.Caption = "Goods In/Supplier Invoices"
                    oFolder.DataBind.SetBound(True, "", "fld6")
                ElseIf i = 4 Then
                    oFolder.Caption = "Goods In/Supplier Invoices/Positions and Costs"
                    oFolder.DataBind.SetBound(True, "", "fld7")
                ElseIf i = 5 Then
                    oFolder.Caption = "Inter Branch Transfers"
                Else
                    oFolder.Caption = "Inter Branch Transfers Confirmation"
                End If
                oFolder.DataBind.SetBound(True, "", "FolderDS")
                If i = 1 Then
                    oFolder.Select()
                    ' oFolder.GroupWith("10")
                Else
                    oFolder.GroupWith("Folder" & i - 1)
                End If
            Next i
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
        End Try
    End Sub
    Private Sub AddSubFolders1(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            For i As Integer = 1 To 6
                oNewItem = aForm.Items.Add("Finance" & i, SAPbouiCOM.BoFormItemTypes.it_FOLDER)
                If i = 1 Then
                    oItem1 = aForm.Items.Item("10")
                    oNewItem.Left = oItem1.Left + (i - 1) + 5
                    oNewItem.Top = oItem1.Top + 30
                    oNewItem.Width = oItem1.Width + 20
                Else
                    oItem1 = aForm.Items.Item("Finance" & i - 1)
                    oNewItem.Left = oItem1.Left + oItem1.Width + 1
                    oNewItem.Top = oItem1.Top
                    oNewItem.Width = oItem1.Width + 20
                End If
                'oNewItem.Width = oItem1.Width + 10
                oNewItem.Height = 20
                oNewItem.FromPane = 8
                oNewItem.ToPane = 13
                Dim oFolder As SAPbouiCOM.Folder
                oFolder = oNewItem.Specific
                If i = 1 Then
                    oFolder.Caption = "Supplier Return Delivery Note"
                    oFolder.DataBind.SetBound(True, "", "fld21")
                ElseIf i = 2 Then
                    oFolder.Caption = "Customer Delivery Notes"
                    ' oFolder.DataBind.SetBound(True, "", "fld11")
                ElseIf i = 3 Then
                    oFolder.Caption = "Supplier Return"
                    ' oFolder.DataBind.SetBound(True, "", "fld12")
                ElseIf i = 4 Then
                    oFolder.Caption = "Supplier Return Positions"
                    '  oFolder.DataBind.SetBound(True, "", "fld13")
                ElseIf i = 5 Then
                    oFolder.Caption = "Supplier Return Costs"
                Else
                    oFolder.Caption = "Customer Credit Invoices"
                End If
                oFolder.DataBind.SetBound(True, "", "FolderDS")
                If i = 1 Then
                    oFolder.Select()
                Else
                    oFolder.GroupWith("Finance" & i - 1)
                End If
            Next i
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
        End Try
    End Sub
    Private Sub AddSubFolders2(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            For i As Integer = 1 To 6
                oNewItem = aForm.Items.Add("Sales" & i, SAPbouiCOM.BoFormItemTypes.it_FOLDER)
                If i = 1 Then
                    oItem1 = aForm.Items.Item("10")
                    oNewItem.Left = oItem1.Left + (i - 1) + 5
                    oNewItem.Top = oItem1.Top + 30
                    oNewItem.Width = oItem1.Width + 20
                Else
                    oItem1 = aForm.Items.Item("Sales" & i - 1)
                    oNewItem.Left = oItem1.Left + oItem1.Width + 1
                    oNewItem.Top = oItem1.Top
                    oNewItem.Width = oItem1.Width + 20
                End If
                oNewItem.Height = 20
                oNewItem.FromPane = 14
                oNewItem.ToPane = 19
                Dim oFolder As SAPbouiCOM.Folder
                oFolder = oNewItem.Specific
                If i = 1 Then
                    oFolder.Caption = "Stock Correction"
                    oFolder.DataBind.SetBound(True, "", "fld18")
                ElseIf i = 2 Then
                    oFolder.Caption = "Stock Take"
                    oFolder.DataBind.SetBound(True, "", "fld19")
                ElseIf i = 3 Then
                    oFolder.Caption = "Sales TurnOver"
                    oFolder.DataBind.SetBound(True, "", "fld20")
                ElseIf i = 4 Then
                    oFolder.Caption = "Incoming Payments"
                ElseIf i = 5 Then
                    oFolder.Caption = "Sales Differences"
                Else
                    oFolder.Caption = "Sales Drop off Not Applicable"
                End If
                oFolder.DataBind.SetBound(True, "", "FolderDS")
                If i = 1 Then
                    oFolder.Select()
                Else
                    oFolder.GroupWith("Sales" & i - 1)
                End If
            Next i
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
        End Try
    End Sub
    Private Sub AddSubFolders3(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            For i As Integer = 1 To 5
                oNewItem = aForm.Items.Add("Payments" & i, SAPbouiCOM.BoFormItemTypes.it_FOLDER)
                If i = 1 Then
                    oItem1 = aForm.Items.Item("10")
                    oNewItem.Left = oItem1.Left + (i - 1) + 5
                    oNewItem.Top = oItem1.Top + 30
                    oNewItem.Width = oItem1.Width + 20
                Else
                    oItem1 = aForm.Items.Item("Payments" & i - 1)
                    oNewItem.Left = oItem1.Left + oItem1.Width + 1
                    oNewItem.Top = oItem1.Top
                    oNewItem.Width = oItem1.Width + 20
                End If
                oNewItem.Height = 20
                oNewItem.FromPane = 20
                oNewItem.ToPane = 24
                Dim oFolder As SAPbouiCOM.Folder
                oFolder = oNewItem.Specific
                If i = 1 Then
                    oFolder.Caption = "Sales Pick Up Not Applicable"
                    oFolder.DataBind.SetBound(True, "", "fld4")
                ElseIf i = 2 Then
                    oFolder.Caption = "Sales Sold Gift Vouchers"
                ElseIf i = 3 Then
                    oFolder.Caption = "Prepayments Not Applicable"
                ElseIf i = 4 Then
                    oFolder.Caption = "Sales Expenses"
                ElseIf i = 5 Then
                    oFolder.Caption = "Credit Balance"
                End If
                oFolder.DataBind.SetBound(True, "", "FolderDS")
                If i = 1 Then
                    oFolder.Select()
                Else
                    oFolder.GroupWith("Payments" & i - 1)
                End If
            Next i
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
        End Try
    End Sub

#End Region



#Region "Browse File"

    '*****************************************************************
    'Type               : Procedure    
    'Name               : BrowseFile
    'Parameter          : Form
    'Return Value       : 
    'Author             :  Senthil Kumar B
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Browse a  File
    '******************************************************************
    Private Sub BrowseFile(ByVal Form As SAPbouiCOM.Form)
        'ShowFileDialog(Form)
    End Sub
#End Region

#Region "ShowFileDialog"

    '*****************************************************************
    'Type               : Procedure
    'Name               : ShowFileDialog
    'Parameter          :
    'Return Value       :
    'Author             : Senthil Kumar B 
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To open a File Browser
    '******************************************************************

    Private Sub fillopen()
        Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
        mythr.SetApartmentState(Threading.ApartmentState.STA)
        mythr.Start()
        mythr.Join()

    End Sub

    Private Sub ShowFileDialog()
        Dim oDialogBox As New OpenFileDialog
        Dim strFileName, strMdbFilePath As String
        Dim oEdit As SAPbouiCOM.EditText
        Dim oProcesses() As Process
        Try
            oProcesses = Process.GetProcessesByName("SAP Business One")
            If oProcesses.Length <> 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    Dim MyWindow As New WindowWrapper(oProcesses(i).MainWindowHandle)
                    If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                        strMdbFilePath = oDialogBox.FileName
                        strSelectedFilepath = oDialogBox.FileName
                        strFileName = strSelectedFilepath
                        strSelectedFolderPath = strFileName
                        If strSelectedFolderPath.EndsWith("\") Then
                            strSelectedFolderPath = strSelectedFilepath.Substring(0, strSelectedFolderPath.Length - 1)
                        End If
                    Else
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
#End Region


#Region "Write into ErrorLog File"
    Private Sub WriteErrorlog(ByVal aMessage As String, ByVal aPath As String)
        Dim aSw As System.IO.StreamWriter
        If File.Exists(aPath) Then
        End If
        aSw = New StreamWriter(aPath, True)
        aMessage = Now.ToString("dd-MM-yyyy hh:mm") & "--> " & aMessage
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
#End Region

#Region "Import"
    Private Sub Import(ByVal aForm As SAPbouiCOM.Form)
        Dim strvalue, strTime, strFileName1 As String
        Dim stpath As String
        oCombobox = aForm.Items.Item("4").Specific
        strvalue = oCombobox.Selected.Value
        'If strvalue = "" Then
        '    oApplication.Utilities.Message("Select the Document Type", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Exit Sub
        'End If
        stpath = oApplication.Utilities.getEdittextvalue(oForm, "6")
        If stpath = "" Then
            oApplication.Utilities.Message("File Name  missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        'If Directory.Exists(stpath) = False Then
        '    oApplication.Utilities.Message("Folder does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Exit Sub
        'End If
        strTime = Now.ToShortTimeString.Replace(":", "")
        strFilename1 = Now.Date.ToString("ddMMyyyy")
        strFileName1 = strFileName1 & strTime
        strImportErrorLog = System.Windows.Forms.Application.StartupPath & "\ImportLog"
        If Directory.Exists(strImportErrorLog) = False Then
            Directory.CreateDirectory(strImportErrorLog)
        End If
        strImportErrorLog = strImportErrorLog & "\Import_" & strFileName1 & ".txt"
        Try
            'If ReadImportFiles(aForm) = False Then
            '    Exit Sub
            'End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End Try
        oApplication.Utilities.WriteErrorHeader(strImportErrorLog, "Import Reading files Processing...")

        oApplication.Utilities.WriteErrorHeader(strImportErrorLog, "Import Reading files Process Completed....")
        oApplication.Utilities.WriteErrorHeader(strImportErrorLog, "Document Creation Processing...")
       
        oApplication.Utilities.WriteErrorHeader(strImportErrorLog, "Document Creation Process Completed....")
        If 1 = 1 Then
            Dim x As System.Diagnostics.ProcessStartInfo
            x = New System.Diagnostics.ProcessStartInfo
            x.UseShellExecute = True
            sPath = strImportErrorLog ' System.Windows.Forms.Application.StartupPath & "\ImportLog.txt"
            If File.Exists(sPath) Then
                x.FileName = sPath
                System.Diagnostics.Process.Start(x)
                x = Nothing
            End If
        End If
        oApplication.Utilities.Message("Export process completed successfully.....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
#End Region

#Region "Read Payroll Interface file"


#Region "Read Import files"
    Private Function ReadImportFiles(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim strvalue As String
            Dim stpath, strImpLogFolder As String
            oCombobox = aForm.Items.Item("4").Specific
            strvalue = oCombobox.Selected.Value
            stpath = oApplication.Utilities.getEdittextvalue(oForm, "6")
            strImpLogFolder = System.Windows.Forms.Application.StartupPath & "\ImportLog.txt"
            strImpLogFolder = strImportErrorLog

            If stpath = "" Then
                oApplication.Utilities.Message("Import folder path is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            sPath = System.Windows.Forms.Application.StartupPath & "\test.txt"
            If File.Exists(sPath) Then
                File.Delete(sPath)
            End If
            If validateFolderPaths(stpath, oCombobox.Selected.Value) = False Then
                Return False
            End If


            Select Case oCombobox.Selected.Value
                Case "SHP"
                    readSOImport(stpath & "\Import\XSO_Export", aForm, sPath)
                Case "ASN"
                    readASNImport(stpath & "\Import\XASN_Export", aForm, sPath)
                Case "ADJ"
                    readADJImport(stpath & "\Import\XINV_Export", aForm, sPath)
                Case "HOLD"
                    readHOLImport(stpath & "\Import\XHOL_Export", aForm, sPath)
                Case "UDT"
                    readSOImport(stpath & "\Import\XSO_Export", aForm, sPath)
                    readASNImport(stpath & "\Import\XASN_Export", aForm, sPath)
                    readADJImport(stpath & "\Import\XINV_Export", aForm, sPath)
                    readHOLImport(stpath & "\Import\XHOL_Export", aForm, sPath)
            End Select
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        Return True
    End Function
#End Region

#Region "Validate Folder path"
    Private Function validateFolderPaths(ByVal aPath As String, ByVal choice As String) As Boolean
        Dim strFolder As String
        Select Case choice
            Case "SHP"
                strFolder = aPath & "\Import\XSO_Export"
                If Directory.Exists(aPath & "\Import\XSO_Export") = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Case "ASN"
                strFolder = aPath & "\Import\XASN_Export"
                If Directory.Exists(strFolder) = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Case "ADJ"
                strFolder = aPath & "\Import\XINV_Export"
                If Directory.Exists(strFolder) = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Case "HOLD"
                strFolder = aPath & "\Import\XHOL_Export"
                If Directory.Exists(strFolder) = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Case "UDT"
                strFolder = aPath & "\Import\XSO_Export"
                If Directory.Exists(aPath & "\Import\XSO_Export") = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                strFolder = aPath & "\Import\XASN_Export"
                If Directory.Exists(strFolder) = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                strFolder = aPath & "\Import\XINV_Export"
                If Directory.Exists(strFolder) = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                strFolder = aPath & "\Import\XHOL_Export"
                If Directory.Exists(strFolder) = False Then
                    oApplication.Utilities.Message("Folder does not exist: " & strFolder, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
        End Select
        Return True
    End Function
#End Region
#Region "Read SO Import"

    Private Sub readSOImport(ByVal aFolderpath As String, ByVal aform As SAPbouiCOM.Form, ByVal aPath As String)
        Dim di As New IO.DirectoryInfo(aFolderpath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strSokey, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strLineno, strImpDocType, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading Shipment files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = "Processing Reading Shipment file..."
            oApplication.Utilities.WriteErrorlog("Reading shipment files...", strImportErrorLog)
            strsuccessfolder = aFolderpath
            strsuccessfolder = aFolderpath & "\Success"
            strErrorfolder = aFolderpath & "\Error"
            If Directory.Exists(strsuccessfolder) = False Then
                Directory.CreateDirectory(strsuccessfolder)
            End If
            If Directory.Exists(strErrorfolder) = False Then
                Directory.CreateDirectory(strErrorfolder)
            End If
            For Each fi In aryFi
                strFilename = fi.FullName
                strSuccessFile = strsuccessfolder & "\" & fi.Name
                strErrorFile = strErrorfolder & "\" & fi.Name
                sr = New StreamReader(fi.FullName, System.Text.Encoding.Default) 'IO.File.OpenText(fil)
                sPath = aPath
                Dim strLIneStrin As String()
                Try
                    oApplication.Utilities.WriteErrorlog("Reading Shipment File Processing...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Shipment File Processing...File Name : " & fi.Name, strImportErrorLog)
                    'oApplication.Utilities.WriteErrorlog("File Name : " & fi.Name, sPath)
                    Dim oRec, oRecUpdate As SAPbobsCOM.Recordset
                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecUpdate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery("Select * from [@Z__XSO] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    If oRec.RecordCount > 0 Then
                        oRec.DoQuery("Delete  from [@Z__XSO] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    End If
                    Do While (sr.Peek <> -1)
                        linje = ""
                        linje = sr.ReadLine()
                        strLIneStrin = linje.Split(vbTab)
                        If strLIneStrin.Length > 0 Then
                            strStorekey = strLIneStrin.GetValue(0)
                            strSokey = strLIneStrin.GetValue(1)
                            strType = strLIneStrin.GetValue(2)
                            If strType = "R" Then
                                strImpDocType = "R"
                            Else
                                strImpDocType = "INVTRN"

                            End If
                            strOrderKey = strLIneStrin.GetValue(3)
                            strShipdate = strLIneStrin.GetValue(4)
                            strSKU = strLIneStrin.GetValue(5)
                            strQty = strLIneStrin.GetValue(6)
                            strbatch = strLIneStrin.GetValue(7)
                            strmfgdate = strLIneStrin.GetValue(8)
                            strexpdate = strLIneStrin.GetValue(9)
                            strLineno = strLIneStrin.GetValue(10)
                            strdate = strShipdate
                            strdate = strdate.ToString.Replace("-", "")
                            DAY = strdate.Substring(0, 2)
                            MONTH = strdate.Substring(2, 2)
                            YEAR = strdate.Substring(4, 4)
                            DATE1 = DAY & MONTH & YEAR
                            dtShipdate = GetDateTimeValue(DATE1)
                            strdate = strmfgdate
                            If strdate <> "" Then

                                strdate = strdate.ToString.Replace("-", "")
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtMfrDate = GetDateTimeValue(DATE1)
                            End If
                            strdate = strexpdate
                            If strdate <> "" Then
                                strdate = strdate.ToString.Replace("-", "")
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtExpDate = GetDateTimeValue(DATE1)
                            End If
                            Dim oUsertable As SAPbobsCOM.UserTable
                            Dim strsql, sCode, strUpdateQuery As String
                            strsql = oApplication.Utilities.getMaxCode("@Z__XSO", "CODE")
                            oUsertable = oApplication.Company.UserTables.Item("Z__XSO")
                            oUsertable.Code = strsql
                            oUsertable.Name = strsql & "M"
                            ' oUsertable.UserFields.Fields.Item("U_Z_DocType").Value = "SO"
                            oUsertable.UserFields.Fields.Item("U_Z_StoreKey").Value = strStorekey
                            oUsertable.UserFields.Fields.Item("U_Z_SAPDocKey").Value = strSokey
                            oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType
                            oUsertable.UserFields.Fields.Item("U_Z_ImpDocType").Value = strImpDocType
                            oUsertable.UserFields.Fields.Item("U_Z_OrderKey").Value = strOrderKey
                            oUsertable.UserFields.Fields.Item("U_Z_Receiptdate").Value = dtShipdate
                            oUsertable.UserFields.Fields.Item("U_Z_SKU").Value = strSKU
                            oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = CDbl(strQty)
                            oUsertable.UserFields.Fields.Item("U_Z_BatchNo").Value = strbatch
                            oUsertable.UserFields.Fields.Item("U_Z_MfrDate").Value = dtMfrDate
                            oUsertable.UserFields.Fields.Item("U_Z_ExpDate").Value = dtExpDate
                            oUsertable.UserFields.Fields.Item("U_Z_LineNo").Value = strLineno
                            oUsertable.UserFields.Fields.Item("U_Z_FileName").Value = fi.Name
                            oUsertable.UserFields.Fields.Item("U_Z_Imported").Value = "N"
                            oUsertable.UserFields.Fields.Item("U_Z_ImpMethod").Value = "M"
                            If oUsertable.Add <> 0 Then
                                MsgBox(oApplication.Company.GetLastErrorDescription)
                                oApplication.Utilities.WriteErrorlog("Error --> " & oApplication.Company.GetLastErrorDescription & " File Name : " & fi.Name, sPath)
                                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oDelrec.DoQuery("Delete from [@Z__XSO] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                            End If
                        End If
                    Loop
                    oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oDelrec.DoQuery("Update [@Z__XSO] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    sr.Close()
                    If File.Exists(strSuccessFile) Then
                        File.Delete(strSuccessFile)
                    End If
                    File.Move(fi.FullName, strSuccessFile)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                    'Return True
                Catch ex As Exception

                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oApplication.Utilities.WriteErrorlog("Reading SO File Failed...File Name : " & fi.Name, strImportErrorLog)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, strImportErrorLog)
                    sr.Close()
                    If File.Exists(strErrorFile) Then
                        File.Delete(strErrorFile)
                    End If
                    File.Move(fi.FullName, strErrorFile)
                    oApplication.Utilities.WriteErrorlog("Reading SO file failed: Filename : " & fi.Name & " Moved to Error folder", strImportErrorLog)
                    ' Return False
                End Try
            Next

            oApplication.Utilities.Message("Reading Shipment file completed...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.WriteErrorlog("Reading Shipment file completed", strImportErrorLog)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = ""
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readASNImport(ByVal aFolderpath As String, ByVal aform As SAPbouiCOM.Form, ByVal apath As String)
        Dim di As New IO.DirectoryInfo(aFolderpath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, Desgfolder, strsokey, strOrderKey, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strLineno, strImpDocType, strType, strdate, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading ASN files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = "Processing Reading ASN file..."
            oApplication.Utilities.WriteErrorlog("Reading ASN Files...", strImportErrorLog)
            strsuccessfolder = aFolderpath
            strsuccessfolder = aFolderpath & "\Success"
            strErrorfolder = aFolderpath & "\Error"
            If Directory.Exists(strsuccessfolder) = False Then
                Directory.CreateDirectory(strsuccessfolder)
            End If
            If Directory.Exists(strErrorfolder) = False Then
                Directory.CreateDirectory(strErrorfolder)
            End If

            For Each fi In aryFi
                strFilename = fi.FullName
                strSuccessFile = strsuccessfolder & "\" & fi.Name
                strErrorFile = strErrorfolder & "\" & fi.Name
                sr = New StreamReader(fi.FullName, System.Text.Encoding.Default) 'IO.File.OpenText(fil)
                sPath = apath
                'If File.Exists(sPath) Then
                '    File.Delete(sPath)
                'End If
                Dim strLIneStrin As String()
                Try
                    Dim oRec, oRecUpdate As SAPbobsCOM.Recordset
                    oApplication.Utilities.WriteErrorlog("Reading ASN File Processing...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading ASN File Processing...File Name : " & fi.Name, strImportErrorLog)
                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecUpdate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery("Select * from [@Z_XASN] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    If oRec.RecordCount > 0 Then
                        oRec.DoQuery("Delete from [@Z_XASN] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    End If
                    Do While (sr.Peek <> -1)
                        linje = ""
                        linje = sr.ReadLine()
                        strLIneStrin = linje.Split(vbTab)
                        If strLIneStrin.Length > 0 Then
                            strStorekey = strLIneStrin.GetValue(0)
                            strsokey = strLIneStrin.GetValue(1)
                            strType = strLIneStrin.GetValue(2)
                            If strType = "" Then
                                strImpDocType = ""
                            End If
                            strImpDocType = "ST"
                            Select Case strType.ToUpper
                                Case "NORMAL"
                                    strImpDocType = "GRPO"
                                Case "I"
                                    strImpDocType = "GRPO"
                                Case "RETRUN ORDER"
                                    strImpDocType = "RETURNS"
                                Case "OR"
                                    strImpDocType = "RETURNS"
                                Case "RETURN INVOICE"
                                    strImpDocType = "ARCR"
                                Case "IR"
                                    strImpDocType = "ARCR"
                                Case "TRN"
                                    strImpDocType = "ST"
                                Case "TRS"
                                    strImpDocType = "ST"
                            End Select

                            strShipdate = strLIneStrin.GetValue(3)
                            strSKU = strLIneStrin.GetValue(4)
                            strQty = strLIneStrin.GetValue(5)
                            strbatch = strLIneStrin.GetValue(6)
                            strmfgdate = strLIneStrin.GetValue(7)
                            strexpdate = strLIneStrin.GetValue(8)
                            strSusr1 = strLIneStrin.GetValue(9)
                            strSur2 = strLIneStrin.GetValue(10)
                            strholdcode = strLIneStrin.GetValue(11)
                            strLineno = strLIneStrin.GetValue(12)

                            strdate = strShipdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtShipdate = GetDateTimeValue(DATE1)

                            End If

                            strdate = strmfgdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtMfrDate = GetDateTimeValue(DATE1)
                            End If

                            strdate = strexpdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtExpDate = GetDateTimeValue(DATE1)
                            End If
                            Dim oUsertable As SAPbobsCOM.UserTable
                            Dim strsql As String
                            strsql = oApplication.Utilities.getMaxCode("@Z_XASN", "CODE")
                            oUsertable = oApplication.Company.UserTables.Item("Z_XASN")
                            oUsertable.Code = strsql
                            oUsertable.Name = strsql & "M"
                            'oUsertable.UserFields.Fields.Item("U_Z_DocType").Value = "ASN"
                            oUsertable.UserFields.Fields.Item("U_Z_StoreKey").Value = strStorekey
                            oUsertable.UserFields.Fields.Item("U_Z_SAPDocKey").Value = strsokey
                            oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType
                            oUsertable.UserFields.Fields.Item("U_Z_ImpDocType").Value = strImpDocType
                            oUsertable.UserFields.Fields.Item("U_Z_Receiptdate").Value = dtShipdate
                            oUsertable.UserFields.Fields.Item("U_Z_SKU").Value = strSKU
                            oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = CDbl(strQty)
                            oUsertable.UserFields.Fields.Item("U_Z_BatchNo").Value = strbatch
                            oUsertable.UserFields.Fields.Item("U_Z_MfrDate").Value = dtMfrDate
                            oUsertable.UserFields.Fields.Item("U_Z_ExpDate").Value = dtExpDate
                            oUsertable.UserFields.Fields.Item("U_Z_LineNo").Value = strLineno
                            oUsertable.UserFields.Fields.Item("U_Z_Susr").Value = strSusr1
                            oUsertable.UserFields.Fields.Item("U_Z_Susr2").Value = strSur2
                            oUsertable.UserFields.Fields.Item("U_Z_HoldCode").Value = strholdcode
                            oUsertable.UserFields.Fields.Item("U_Z_FileName").Value = fi.Name
                            oUsertable.UserFields.Fields.Item("U_Z_Imported").Value = "N"

                            oUsertable.UserFields.Fields.Item("U_Z_ImpMethod").Value = "M"
                            If oUsertable.Add <> 0 Then
                                MsgBox(oApplication.Company.GetLastErrorDescription)
                                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oDelrec.DoQuery("Delete from [@Z_XASN] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                            End If
                        End If


                    Loop
                    oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oDelrec.DoQuery("Update [@Z_XASN] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")

                    sr.Close()
                    File.Move(fi.FullName, strSuccessFile)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                    'Return True
                Catch ex As Exception
                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oApplication.Utilities.WriteErrorlog("Reading ADN File Failed...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading ADN file Failed...File Name : " & fi.Name, strImportErrorLog)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, strImportErrorLog)
                    sr.Close()
                    If File.Exists(strErrorFile) Then
                        File.Delete(strErrorFile)
                    End If
                    File.Move(fi.FullName, strErrorFile)
                    oApplication.Utilities.WriteErrorlog("Reading SO file failed: Filename : " & fi.Name & " Moved to Error folder", strImportErrorLog)

                    ' Return False
                End Try
            Next
            oApplication.Utilities.Message("Reading ASN Import completed...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.WriteErrorlog("Reading ADN File Completed", strImportErrorLog)

            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = ""
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readADJImport(ByVal aFolderpath As String, ByVal aform As SAPbouiCOM.Form, ByVal apath As String)
        Dim di As New IO.DirectoryInfo(aFolderpath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading ADJ files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = "Processing Reading ADJ file..."
            oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            strsuccessfolder = aFolderpath
            strsuccessfolder = aFolderpath & "\Success"
            strErrorfolder = aFolderpath & "\Error"
            If Directory.Exists(strsuccessfolder) = False Then
                Directory.CreateDirectory(strsuccessfolder)
            End If
            If Directory.Exists(strErrorfolder) = False Then
                Directory.CreateDirectory(strErrorfolder)
            End If

            For Each fi In aryFi
                strFilename = fi.FullName
                strSuccessFile = strsuccessfolder & "\" & fi.Name
                strErrorFile = strErrorfolder & "\" & fi.Name
                sr = New StreamReader(fi.FullName, System.Text.Encoding.Default) 'IO.File.OpenText(fil)
                sPath = apath 'System.Windows.Forms.Application.StartupPath & "\ImportLog.txt"

                Dim strLIneStrin As String()
                Try
                    'WriteErrorlog("File Name : " & strFilename, sPath)
                    'WriteErrorlog("Import Process Starting.....", sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Adjustment File Processing...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Adjustment File Processing...File Name : " & fi.Name, strImportErrorLog)
                    Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTemp.DoQuery("SELECT T0.[DfltWhs] FROM OADM T0")
                    strwhs = oTemp.Fields.Item(0).Value

                    oRecUpdate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery("Select * from [@Z_XADJ] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    If oRec.RecordCount > 0 Then
                        oRec.DoQuery("Delete from [@Z_XADJ] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    End If
                    Do While (sr.Peek <> -1)
                        linje = ""
                        linje = sr.ReadLine()
                        strLIneStrin = linje.Split(vbTab)
                        If strLIneStrin.Length > 7 Then
                            strStorekey = strLIneStrin.GetValue(0)
                            strsokey = strLIneStrin.GetValue(1)
                            strSKU = strLIneStrin.GetValue(2)
                            strbatch = strLIneStrin.GetValue(3)
                            strmfgdate = strLIneStrin.GetValue(4)
                            strexpdate = strLIneStrin.GetValue(5)
                            strQty = strLIneStrin.GetValue(6)
                            If strQty.Contains("-") Then
                                strImpDocType = "Goods Issue"
                            Else
                                strImpDocType = "Goods Recipt"
                            End If
                            strremarks = strLIneStrin.GetValue(7)

                            strdate = strmfgdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtMfrDate = GetDateTimeValue(DATE1)
                            End If


                            strdate = strexpdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtExpDate = GetDateTimeValue(DATE1)
                            End If
                            Dim oUsertable As SAPbobsCOM.UserTable
                            Dim strsql As String
                            strsql = oApplication.Utilities.getMaxCode("@Z_XADJ", "CODE")
                            oUsertable = oApplication.Company.UserTables.Item("Z_XADJ")
                            oUsertable.Code = strsql
                            oUsertable.Name = strsql & "M"
                            oUsertable.UserFields.Fields.Item("U_Z_StoreKey").Value = strStorekey
                            oUsertable.UserFields.Fields.Item("U_Z_Adjkey").Value = strsokey
                            oUsertable.UserFields.Fields.Item("U_Z_ImpDocType").Value = strImpDocType
                            oUsertable.UserFields.Fields.Item("U_Z_SKU").Value = strSKU
                            oUsertable.UserFields.Fields.Item("U_Z_Remarks").Value = strremarks
                            oUsertable.UserFields.Fields.Item("U_Z_BatchNo").Value = strbatch
                            oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = CDbl(strQty)
                            oUsertable.UserFields.Fields.Item("U_Z_MfrDate").Value = dtMfrDate
                            oUsertable.UserFields.Fields.Item("U_Z_ExpDate").Value = dtExpDate
                            oUsertable.UserFields.Fields.Item("U_Z_Whs").Value = strwhs
                            oUsertable.UserFields.Fields.Item("U_Z_FileName").Value = fi.Name
                            oUsertable.UserFields.Fields.Item("U_Z_Imported").Value = "N"
                            oUsertable.UserFields.Fields.Item("U_Z_ImpMethod").Value = "M"
                            If oUsertable.Add <> 0 Then
                                MsgBox(oApplication.Company.GetLastErrorDescription)
                                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                            End If
                        End If
                    Loop
                    oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    sr.Close()
                    File.Move(fi.FullName, strSuccessFile)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                    'Return True
                Catch ex As Exception
                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oApplication.Utilities.WriteErrorlog("Reading ADJ File Failed...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading ADJ File Failed...File Name : " & fi.Name, strImportErrorLog)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, strImportErrorLog)
                    sr.Close()
                    If File.Exists(strErrorFile) Then
                        File.Delete(strErrorFile)
                    End If
                    File.Move(fi.FullName, strErrorFile)
                    oApplication.Utilities.WriteErrorlog("Reading SO file failed: Filename : " & fi.Name & " Moved to Error folder", strImportErrorLog)

                    ' Return False
                End Try
            Next
            oApplication.Utilities.Message("Reading Adjustment file completed...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.WriteErrorlog("Reading Adjustment file completed", strImportErrorLog)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = ""
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readHOLImport(ByVal aFolderpath As String, ByVal aform As SAPbouiCOM.Form, ByVal apath As String)
        Dim di As New IO.DirectoryInfo(aFolderpath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strfrmwhs, strtowhs, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading HOLD files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = "Processing Reading ADJ file..."
            oApplication.Utilities.WriteErrorlog("Reading HOLD Files processing..", strImportErrorLog)
            strsuccessfolder = aFolderpath
            strsuccessfolder = aFolderpath & "\Success"
            strErrorfolder = aFolderpath & "\Error"
            If Directory.Exists(strsuccessfolder) = False Then
                Directory.CreateDirectory(strsuccessfolder)
            End If
            If Directory.Exists(strErrorfolder) = False Then
                Directory.CreateDirectory(strErrorfolder)
            End If

            For Each fi In aryFi
                strFilename = fi.FullName
                strSuccessFile = strsuccessfolder & "\" & fi.Name
                strErrorFile = strErrorfolder & "\" & fi.Name
                sr = New StreamReader(fi.FullName, System.Text.Encoding.Default) 'IO.File.OpenText(fil)
                sPath = apath 'System.Windows.Forms.Application.StartupPath & "\ImportLog.txt"

                Dim strLIneStrin As String()
                Try
                    'WriteErrorlog("File Name : " & strFilename, sPath)
                    'WriteErrorlog("Import Process Starting.....", sPath)
                    oApplication.Utilities.WriteErrorlog("Reading HOLD File Processing...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading HOLD File Processing...File Name : " & fi.Name, strImportErrorLog)
                    Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTemp.DoQuery("SELECT T0.[DfltWhs] FROM OADM T0")
                    strwhs = oTemp.Fields.Item(0).Value
                    oRecUpdate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery("Select * from [@Z_XHOL] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    If oRec.RecordCount > 0 Then
                        oRec.DoQuery("Delete from [@Z_XHOL] where U_Z_FileName='" & fi.Name & "' and U_Z_Imported='N'")
                    End If
                    Do While (sr.Peek <> -1)
                        linje = ""
                        linje = sr.ReadLine()
                        strLIneStrin = linje.Split(vbTab)
                        If strLIneStrin.Length > 7 Then
                            strfrmwhs = strLIneStrin.GetValue(0)
                            strtowhs = strLIneStrin.GetValue(1)
                            strremarks = strLIneStrin.GetValue(2)
                            strSKU = strLIneStrin.GetValue(3)
                            strbatch = strLIneStrin.GetValue(4)
                            strmfgdate = strLIneStrin.GetValue(5)
                            strexpdate = strLIneStrin.GetValue(6)
                            strQty = strLIneStrin.GetValue(7)

                            strdate = strmfgdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtMfrDate = GetDateTimeValue(DATE1)
                            End If
                            strQty = strQty.Replace(".", CompanyDecimalSeprator)
                            strdate = strexpdate.Replace("-", "")
                            If strdate <> "" Then
                                DAY = strdate.Substring(0, 2)
                                MONTH = strdate.Substring(2, 2)
                                YEAR = strdate.Substring(4, 4)
                                DATE1 = DAY & MONTH & YEAR
                                dtExpDate = GetDateTimeValue(DATE1)
                            End If
                            strImpDocType = "ST"
                            Dim oUsertable As SAPbobsCOM.UserTable
                            Dim strsql As String
                            strsql = oApplication.Utilities.getMaxCode("@Z_XHOL", "CODE")
                            oUsertable = oApplication.Company.UserTables.Item("Z_XHOL")
                            oUsertable.Code = strsql
                            oUsertable.Name = strsql & "M"
                            oUsertable.UserFields.Fields.Item("U_Z_FrmWhs").Value = strfrmwhs
                            oUsertable.UserFields.Fields.Item("U_Z_ToWhs").Value = strtowhs
                            oUsertable.UserFields.Fields.Item("U_Z_ImpDocType").Value = strImpDocType
                            oUsertable.UserFields.Fields.Item("U_Z_SKU").Value = strSKU
                            oUsertable.UserFields.Fields.Item("U_Z_Remarks").Value = strremarks
                            oUsertable.UserFields.Fields.Item("U_Z_BatchNo").Value = strbatch
                            oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = CDbl(strQty)
                            oUsertable.UserFields.Fields.Item("U_Z_MfrDate").Value = dtMfrDate
                            oUsertable.UserFields.Fields.Item("U_Z_ExpDate").Value = dtExpDate

                            oUsertable.UserFields.Fields.Item("U_Z_FileName").Value = fi.Name
                            oUsertable.UserFields.Fields.Item("U_Z_Imported").Value = "N"
                            oUsertable.UserFields.Fields.Item("U_Z_ImpMethod").Value = "M"
                            If oUsertable.Add <> 0 Then
                                MsgBox(oApplication.Company.GetLastErrorDescription)
                                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oDelrec.DoQuery("Delete from [@Z_XHOL] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                            End If
                        End If
                    Loop
                    oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oDelrec.DoQuery("Update [@Z_XHOL] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    sr.Close()
                    File.Move(fi.FullName, strSuccessFile)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                    oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                    'Return True
                Catch ex As Exception
                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oApplication.Utilities.WriteErrorlog("Reading HOLD File Failed...File Name : " & fi.Name, sPath)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, sPath)
                    oApplication.Utilities.WriteErrorlog("Reading HOLD File Failed...File Name : " & fi.Name, strImportErrorLog)
                    oApplication.Utilities.WriteErrorlog("Error -> " & ex.Message, strImportErrorLog)
                    sr.Close()
                    If File.Exists(strErrorFile) Then
                        File.Delete(strErrorFile)
                    End If
                    File.Move(fi.FullName, strErrorFile)
                    oApplication.Utilities.WriteErrorlog("Reading HOLD file failed: Filename : " & fi.Name & " Moved to Error folder", strImportErrorLog)

                    ' Return False
                End Try
            Next
            oApplication.Utilities.Message("Reading HOLD file completed...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.WriteErrorlog("Reading HOLD file completed", strImportErrorLog)
            oStaticText = aform.Items.Item("9").Specific
            oStaticText.Caption = ""
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Read Interface Files"
    Private Sub readFiles(ByVal aform As SAPbouiCOM.Form, ByVal aPath As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strSokey, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strLineno, strImpDocType, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading  files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'strsuccessfolder = aFolderpath
            'strsuccessfolder = aFolderpath & "\Success"
            'strErrorfolder = aFolderpath & "\Error"
            'If Directory.Exists(strsuccessfolder) = False Then
            '    Directory.CreateDirectory(strsuccessfolder)
            'End If
            'If Directory.Exists(strErrorfolder) = False Then
            '    Directory.CreateDirectory(strErrorfolder)
            'End If
            strFilename = aPath
            'strSuccessFile = strsuccessfolder & "\" & fi.Name
            'strErrorFile = strErrorfolder & "\" & fi.Name
            sr = New StreamReader(aPath, System.Text.Encoding.Default) 'IO.File.OpenText(fil)
            sPath = aPath
            Dim strLIneStrin As String()
            Try
                'oApplication.Utilities.WriteErrorlog("File Name : " & fi.Name, sPath)
                Dim oRec, oRecUpdate As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecUpdate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Do While (sr.Peek <> -1)
                    linje = ""
                    linje = sr.ReadLine()
                    strLIneStrin = linje.Split(vbTab)
                    If strLIneStrin.Length > 0 Then
                        Select Case strLIneStrin.GetValue(0)
                            Case "GIDN"
                                ReadGIDN(linje)
                            Case "GIDNCO"
                                ReadGIDNCO(linje)
                                'Case "GIIV"
                                '    ReadGIIV(linje)
                                'Case "GIIVLG"
                                '    ReadGIIVLG(linje)
                                'Case "GIIVLC"
                                '    ReadGIIVLC(linje)
                                'Case "IBTDN"
                                '    ReadIBTDN(linje)
                                'Case "IBTCO"
                                '    ReadIBTCO(linje)
                                'Case "SURDN"
                                '    ReadSURDN(linje)
                                'Case "CUSDN"
                                '    ReadCUSDN(linje)
                                'Case "SURIV"
                                '    ReadSURIV(linje)
                                'Case "SURIVLG"
                                '    ReadSURIVLG(linje)
                                'Case "SURIVLC"
                                '    ReadSURIVLC(linje)
                                'Case "CURIV"
                                '    ReadCURIV(linje)
                                'Case "STKCOR"
                                '    readSTKCOR(linje)
                                'Case "STKTAK"
                                '    readSTKTAK(linje)
                            Case "SLSTRN"
                                readSLSTRN(linje)
                            Case "SLSPAY"
                                readSLSPAY(linje)
                            Case "SLSDIF"
                                readSLSDIF(linje)
                                'Case "SLSDRP"
                                '    readSLSDRP(linje)
                                'Case "SLSPUP"
                                '    readSLSPUP(linje)
                                'Case "SLSGFT"
                                '    readSLSGFT(linje)
                                'Case "SLSPRE"
                                '    readSLSPRE(linje)
                                'Case "SLSEXP"
                                '    readSLSEXP(linje)
                                'Case "CREBAL"
                                '    readCREBAL(linje)
                        End Select
                    End If
                Loop
                'Return True
            Catch ex As Exception
                '   oApplication.Utilities.WriteErrorlog("Reading SO file failed: Filename : " & fi.Name & " Moved to Error folder", strImportErrorLog)
                oApplication.Utilities.Message("Readin selected file failed : " & ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            End Try
            oApplication.Utilities.Message("Reading process completed", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            '  oApplication.Utilities.WriteErrorlog("Reading  file completed", strImportErrorLog)
          
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub ReadGIDN(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading ADJ files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(vbTab)

                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strAccount, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strOrigin = strLIneStrin.GetValue(2)
                    strgoodsno = strLIneStrin.GetValue(3)
                    strVat = strLIneStrin.GetValue(4)
                    strSupplier = strLIneStrin.GetValue(5)
                    strBrand = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strGoodsDate = strLIneStrin.GetValue(8)
                    strSupDocNo = strLIneStrin.GetValue(9)
                    strCurrency = strLIneStrin.GetValue(10)
                    strGoodsBranch = strLIneStrin.GetValue(11)
                    strvalue = strLIneStrin.GetValue(12)
                    strPaydate = strLIneStrin.GetValue(13)
                    strPeriodPay = strLIneStrin.GetValue(14)
                    strvatpercentage = strLIneStrin.GetValue(15)
                    strsuppliertype = strLIneStrin.GetValue(16)
                    strdate = strGoodsDate
                    If strdate <> "" Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    End If
                    strdate = strPaydate
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtExpDate = GetDateTimeValue(DATE1)
                    Else
                        strPaydate = ""
                    End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@GIDN", "Code")

                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from [@GIDN] where U_Z_Imported='Y' and U_Z_Originator='" & strOrigin & "' and  U_Z_GoodsNo=" & strgoodsno & " and U_Z_Brand='" & strBrand & "'")
                    If oTest.RecordCount > 0 Then
                        Exit Sub
                    Else
                        oTest.DoQuery("Delete from [@GIDN] where U_Z_Imported<>'Y' and U_Z_Originator='" & strOrigin & "' and  U_Z_GoodsNo=" & strgoodsno & " and U_Z_Brand='" & strBrand & "'")

                    End If
                    oUsertable = oApplication.Company.UserTables.Item("GIDN")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Originator").Value = strOrigin
                    oUsertable.UserFields.Fields.Item("U_Z_GoodsNo").Value = strgoodsno
                    oUsertable.UserFields.Fields.Item("U_Z_VatKey").Value = strVat
                    oUsertable.UserFields.Fields.Item("U_Z_Supplier").Value = strSupplier
                    oUsertable.UserFields.Fields.Item("U_Z_SupplierDocNo").Value = strSupDocNo
                    oUsertable.UserFields.Fields.Item("U_Z_Brand").Value = strBrand
                    If strGoodsDate <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_GoodsDate").Value = dtMfrDate
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_AcctNo").Value = strAccount
                    oUsertable.UserFields.Fields.Item("U_Z_Currency").Value = strCurrency
                    oUsertable.UserFields.Fields.Item("U_Z_GoodsBranch").Value = strGoodsBranch
                    oUsertable.UserFields.Fields.Item("U_Z_Value").Value = strvalue
                    If strPaydate <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_PayDate").Value = dtExpDate
                    End If

                    oUsertable.UserFields.Fields.Item("U_Z_PeriodPay").Value = strPeriodPay
                    oUsertable.UserFields.Fields.Item("U_Z_VatPercentage").Value = strvatpercentage
                    oUsertable.UserFields.Fields.Item("U_Z_SupplierType").Value = strsuppliertype
                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False

            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Sub ReadSURDN(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading SURDN files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String

                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strOrigin = strLIneStrin.GetValue(2)
                    strDeliveryNoteNo = strLIneStrin.GetValue(3)
                    strVatKey = strLIneStrin.GetValue(4)
                    strToSupplier = strLIneStrin.GetValue(5)
                    strBrand = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strDeliveryNoteDt = strLIneStrin.GetValue(8)
                    strFromBranch = strLIneStrin.GetValue(9)
                    strComments = strLIneStrin.GetValue(10)
                    strQuantity = strLIneStrin.GetValue(11)
                    strDelNotePrice = strLIneStrin.GetValue(12)
                    strSalesPriceNet = strLIneStrin.GetValue(13)
                    strSalesPriceVat = strLIneStrin.GetValue(14)
                    stroSalesPriceNet = strLIneStrin.GetValue(15)
                    strOSalesPriceVAT = strLIneStrin.GetValue(16)
                    strperrateofVAT = strLIneStrin.GetValue(17)
                    strsuppliertype = strLIneStrin.GetValue(18)
                    strdate = strDeliveryNoteDt
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDeliveryNoteDt = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@SURDN", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("SURDN")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Originator").Value = strOrigin
                    oUsertable.UserFields.Fields.Item("U_Z_DeliveryNoteNo").Value = oApplication.Utilities.getDocumentQuantity(strDeliveryNoteNo)
                    oUsertable.UserFields.Fields.Item("U_Z_VatKey").Value = strVatKey
                    oUsertable.UserFields.Fields.Item("U_Z_ToSupplier").Value = strToSupplier
                    oUsertable.UserFields.Fields.Item("U_Z_Brand").Value = strBrand

                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    If strDeliveryNoteDt <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_DeliveryNoteDt").Value = dtMfrDate
                    End If

                    oUsertable.UserFields.Fields.Item("U_Z_FromBranch").Value = strFromBranch
                    oUsertable.UserFields.Fields.Item("U_Z_Comment").Value = strComments
                    oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = oApplication.Utilities.getDocumentQuantity(strQuantity)
                    oUsertable.UserFields.Fields.Item("U_Z_DelNotePrice").Value = oApplication.Utilities.getDocumentQuantity(strDelNotePrice)

                    oUsertable.UserFields.Fields.Item("U_Z_SalePriceNet").Value = oApplication.Utilities.getDocumentQuantity(strSalesPriceNet)
                    oUsertable.UserFields.Fields.Item("U_Z_SalePriceVat").Value = oApplication.Utilities.getDocumentQuantity(strSalesPriceVat)
                    oUsertable.UserFields.Fields.Item("U_Z_OsalePriceNet").Value = oApplication.Utilities.getDocumentQuantity(stroSalesPriceNet)
                    oUsertable.UserFields.Fields.Item("U_Z_OSalePriceVat").Value = oApplication.Utilities.getDocumentQuantity(strOSalesPriceVAT)

                    oUsertable.UserFields.Fields.Item("U_Z_PerRateofVAT").Value = oApplication.Utilities.getDocumentQuantity(strperrateofVAT)
                    oUsertable.UserFields.Fields.Item("U_Z_SupplierType").Value = oApplication.Utilities.getDocumentQuantity(strsuppliertype)
                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub ReadCUSDN(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading SURDN files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String

                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strOrigin = strLIneStrin.GetValue(2)
                    strDeliveryNoteNo = strLIneStrin.GetValue(3)
                    strVatKey = strLIneStrin.GetValue(4)
                    strSupplier = strLIneStrin.GetValue(5)
                    strBrand = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strDeliveryNoteDt = strLIneStrin.GetValue(8)
                    strFromBranch = strLIneStrin.GetValue(9)
                    strToCustomer = strLIneStrin.GetValue(10)
                    strComments = strLIneStrin.GetValue(11)
                    strQuantity = strLIneStrin.GetValue(12)
                    strDelNotePrice = strLIneStrin.GetValue(13)
                    strSalesPriceNet = strLIneStrin.GetValue(14)
                    strSalesPriceVat = strLIneStrin.GetValue(15)
                    stroSalesPriceNet = strLIneStrin.GetValue(16)
                    strOSalesPriceVAT = strLIneStrin.GetValue(17)
                    strperrateofVAT = strLIneStrin.GetValue(18)
                    strsuppliertype = strLIneStrin.GetValue(19)
                    strdate = strDeliveryNoteDt
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDeliveryNoteDt = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@CUSDN", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("CUSDN")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Originator").Value = strOrigin
                    oUsertable.UserFields.Fields.Item("U_Z_DeliveryNoteNo").Value = oApplication.Utilities.getDocumentQuantity(strDeliveryNoteNo)
                    oUsertable.UserFields.Fields.Item("U_Z_VatKey").Value = strVatKey
                    oUsertable.UserFields.Fields.Item("U_Z_Supplier").Value = strSupplier
                    oUsertable.UserFields.Fields.Item("U_Z_Brand").Value = strBrand

                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    If strDeliveryNoteDt <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_DeliveryNoteDt").Value = dtMfrDate
                    End If

                    oUsertable.UserFields.Fields.Item("U_Z_FromBranch").Value = strFromBranch
                    oUsertable.UserFields.Fields.Item("U_Z_ToCustomer").Value = strToCustomer
                    oUsertable.UserFields.Fields.Item("U_Z_Comment").Value = strComments
                    oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = oApplication.Utilities.getDocumentQuantity(strQuantity)
                    oUsertable.UserFields.Fields.Item("U_Z_DelNotePrice").Value = oApplication.Utilities.getDocumentQuantity(strDelNotePrice)

                    oUsertable.UserFields.Fields.Item("U_Z_SalePriceNet").Value = oApplication.Utilities.getDocumentQuantity(strSalesPriceNet)
                    oUsertable.UserFields.Fields.Item("U_Z_SalePriceVat").Value = oApplication.Utilities.getDocumentQuantity(strSalesPriceVat)
                    oUsertable.UserFields.Fields.Item("U_Z_OsalePriceNet").Value = oApplication.Utilities.getDocumentQuantity(stroSalesPriceNet)
                    oUsertable.UserFields.Fields.Item("U_Z_OSalePriceVat").Value = oApplication.Utilities.getDocumentQuantity(strOSalesPriceVAT)

                    oUsertable.UserFields.Fields.Item("U_Z_PerRateofVAT").Value = oApplication.Utilities.getDocumentQuantity(strperrateofVAT)
                    oUsertable.UserFields.Fields.Item("U_Z_SupplierType").Value = oApplication.Utilities.getDocumentQuantity(strsuppliertype)
                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub ReadSURIV(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading SURIV files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strInvoiceNo, strInvDate, strDeliveryNote, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String

                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strOrigin = strLIneStrin.GetValue(2)
                    strInvoiceNo = strLIneStrin.GetValue(3)
                    strInvDate = strLIneStrin.GetValue(4)
                    strDeliveryNote = strLIneStrin.GetValue(5)
                    strFromBranch = strLIneStrin.GetValue(6)
                    strToSupplier = strLIneStrin.GetValue(7)
                    strComments = strLIneStrin.GetValue(8)
                    
                    strdate = strInvDate
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strInvDate = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@SURIV", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("SURIV")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Originator").Value = strOrigin
                    oUsertable.UserFields.Fields.Item("U_Z_InvoiceNo").Value = oApplication.Utilities.getDocumentQuantity(strInvoiceNo)
                    If strInvDate <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_InvDate").Value = dtMfrDate
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_DeliveryNote").Value = strDeliveryNote
                    oUsertable.UserFields.Fields.Item("U_Z_FromBranch").Value = strFromBranch

                    oUsertable.UserFields.Fields.Item("U_Z_ToSupplier").Value = strToSupplier
                    oUsertable.UserFields.Fields.Item("U_Z_Comment").Value = strComments
                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub ReadSURIVLG(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading SURIVLG files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strInvoiceNo, strInvDate, strDeliveryNote, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String

                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strOrigin = strLIneStrin.GetValue(2)
                    strInvoiceNo = strLIneStrin.GetValue(3)
                    strSupplier = strLIneStrin.GetValue(4)
                    strBrand = strLIneStrin.GetValue(5)
                    strAccount = strLIneStrin.GetValue(6)
                    strVatKey = strLIneStrin.GetValue(7)
                    strQuantity = strLIneStrin.GetValue(8)
                    strDelNotePrice = strLIneStrin.GetValue(9)
                    strSalesPriceNet = strLIneStrin.GetValue(10)
                    strSalesPriceVat = strLIneStrin.GetValue(11)
                    stroSalesPriceNet = strLIneStrin.GetValue(12)
                    strOSalesPriceVAT = strLIneStrin.GetValue(13)
                    strperrateofVAT = strLIneStrin.GetValue(14)

                    'strdate = strInvDate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtMfrDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strInvDate = ""
                    'End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@SURIVLG", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("SURIVLG")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Originator").Value = strOrigin
                    oUsertable.UserFields.Fields.Item("U_Z_InvoiceNo").Value = oApplication.Utilities.getDocumentQuantity(strInvoiceNo)
                    'If strInvDate <> "" Then
                    '    oUsertable.UserFields.Fields.Item("U_Z_InvDate").Value = dtMfrDate
                    'End If
                    oUsertable.UserFields.Fields.Item("U_Z_Supplier").Value = strSupDocNo
                    oUsertable.UserFields.Fields.Item("U_Z_Brand").Value = strBrand
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    oUsertable.UserFields.Fields.Item("U_Z_VatKey").Value = strVatKey

                    oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = oApplication.Utilities.getDocumentQuantity(strQuantity)
                    oUsertable.UserFields.Fields.Item("U_Z_DelNotePrice").Value = oApplication.Utilities.getDocumentQuantity(strDelNotePrice)

                    oUsertable.UserFields.Fields.Item("U_Z_SalePriceNet").Value = oApplication.Utilities.getDocumentQuantity(strSalesPriceNet)
                    oUsertable.UserFields.Fields.Item("U_Z_SalePriceVat").Value = oApplication.Utilities.getDocumentQuantity(strSalesPriceVat)
                    oUsertable.UserFields.Fields.Item("U_Z_OsalePriceNet").Value = oApplication.Utilities.getDocumentQuantity(stroSalesPriceNet)
                    oUsertable.UserFields.Fields.Item("U_Z_OSalePriceVat").Value = oApplication.Utilities.getDocumentQuantity(strOSalesPriceVAT)

                    oUsertable.UserFields.Fields.Item("U_Z_PerRateofVAT").Value = oApplication.Utilities.getDocumentQuantity(strperrateofVAT)
                     If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub ReadSURIVLC(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading SURIVLG files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strInvoiceNo, strInvDate, strDeliveryNote, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String
                    Dim strFrightPer, strFreightnet, strFreeightVat, strTransInsPer, strTransInsNet, strTransInsVat, strTransCostNet, strtranscostvat As String
                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strOrigin = strLIneStrin.GetValue(2)
                    strInvoiceNo = strLIneStrin.GetValue(3)
                    strFrightPer = strLIneStrin.GetValue(4)
                    strFreightnet = strLIneStrin.GetValue(5)
                    strFreeightVat = strLIneStrin.GetValue(6)
                    strTransInsPer = strLIneStrin.GetValue(7)
                    strTransInsNet = strLIneStrin.GetValue(8)
                    strTransInsVat = strLIneStrin.GetValue(9)
                    strtranscostvat = strLIneStrin.GetValue(10)
                    strtranscostvat = strLIneStrin.GetValue(11)
                    'strdate = strInvDate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtMfrDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strInvDate = ""
                    'End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@SURIVLC", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("SURIVLC")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Originator").Value = strOrigin
                    oUsertable.UserFields.Fields.Item("U_Z_InvoiceNo").Value = oApplication.Utilities.getDocumentQuantity(strInvoiceNo)
                    'If strInvDate <> "" Then
                    '    oUsertable.UserFields.Fields.Item("U_Z_InvDate").Value = dtMfrDate
                    'End If
                    oUsertable.UserFields.Fields.Item("U_Z_FreightNet").Value = oApplication.Utilities.getDocumentQuantity(strFreightnet)
                    oUsertable.UserFields.Fields.Item("U_Z_FreightPer").Value = oApplication.Utilities.getDocumentQuantity(strFrightPer)
                    oUsertable.UserFields.Fields.Item("U_Z_FreightVat").Value = oApplication.Utilities.getDocumentQuantity(strFreeightVat)
                
                    oUsertable.UserFields.Fields.Item("U_Z_TransInsPer").Value = oApplication.Utilities.getDocumentQuantity(strTransInsPer)
                    oUsertable.UserFields.Fields.Item("U_Z_TransInsNet").Value = oApplication.Utilities.getDocumentQuantity(strTransInsNet)

                    oUsertable.UserFields.Fields.Item("U_Z_TransInsVat").Value = oApplication.Utilities.getDocumentQuantity(strTransInsVat)
                    oUsertable.UserFields.Fields.Item("U_Z_TransCostNet").Value = oApplication.Utilities.getDocumentQuantity(strTransCostNet)
                    oUsertable.UserFields.Fields.Item("U_Z_TransCostVat").Value = oApplication.Utilities.getDocumentQuantity(strtranscostvat)
                    
                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub ReadGIDNCO(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading GIDNCO files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(vbTab)
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strAccount, strAccountTxt, strVattxt, strnetValue, strvatpercentage, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strsuppliertype As String

                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strOrigin = strLIneStrin.GetValue(2)
                    strgoodsno = strLIneStrin.GetValue(3)
                    strAccount = strLIneStrin.GetValue(4)
                    strAccountTxt = strLIneStrin.GetValue(5)
                    strVattxt = strLIneStrin.GetValue(6)
                    strnetValue = strLIneStrin.GetValue(7)
                    strvatpercentage = strLIneStrin.GetValue(8)
                    strsuppliertype = strLIneStrin.GetValue(9)
                    
                    'strdate = strGoodsDate
                    'If strdate <> "" Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtMfrDate = GetDateTimeValue(DATE1)
                    'End If
                    'strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from [@GIDNCO] where U_Z_Imported='Y' and U_Z_Originator='" & strOrigin & "' and  U_Z_GoodsNo=" & strgoodsno)
                    If oTest.RecordCount > 0 Then
                        Exit Sub
                    Else
                        oTest.DoQuery("Delete from [@GIDNCO] where U_Z_Imported<>'Y' and U_Z_Originator='" & strOrigin & "' and  U_Z_GoodsNo=" & strgoodsno)
                    End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@GIDNCO", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("GIDNCO")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Originator").Value = strOrigin
                    oUsertable.UserFields.Fields.Item("U_Z_GoodsNo").Value = strgoodsno
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    oUsertable.UserFields.Fields.Item("U_Z_Accounttxt").Value = strAccountTxt
                    oUsertable.UserFields.Fields.Item("U_Z_VatKey").Value = strVattxt
                    oUsertable.UserFields.Fields.Item("U_Z_NetValue").Value = oApplication.Utilities.getDocumentQuantity(strnetValue)
                    oUsertable.UserFields.Fields.Item("U_Z_VatPercentage").Value = strvatpercentage
                    oUsertable.UserFields.Fields.Item("U_Z_SupplierType").Value = strsuppliertype
                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub ReadGIIV(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading GIIV files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strAccount, strinvoiceno, strSupplierInvNo, strInvoicedate, strTotalGrAmt, strVattxt, strnetValue, strvatpercentage, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strsuppliertype As String

                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strOrigin = strLIneStrin.GetValue(2)
                    strinvoiceno = strLIneStrin.GetValue(3)
                    strSupplier = strLIneStrin.GetValue(4)
                    strBrand = strLIneStrin.GetValue(5)
                    strSupplierInvNo = strLIneStrin.GetValue(6)
                    strInvoicedate = strLIneStrin.GetValue(7)
                    strCurrency = strLIneStrin.GetValue(8)
                    strTotalGrAmt = strLIneStrin.GetValue(9)

                    strdate = strInvoicedate
                    If strdate <> "" Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    End If
                    'strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@GIIV", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("GIIV")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Originator").Value = strOrigin
                    oUsertable.UserFields.Fields.Item("U_Z_InvoiceNo").Value = strinvoiceno
                    oUsertable.UserFields.Fields.Item("U_Z_Supplier").Value = strSupplier
                    oUsertable.UserFields.Fields.Item("U_Z_Brand").Value = strBrand
                    oUsertable.UserFields.Fields.Item("U_Z_SupplierInvNo").Value = strSupplierInvNo
                    If strInvoicedate <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_Invoicedate").Value = dtMfrDate
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_TotalGrAmt").Value = oApplication.Utilities.getDocumentQuantity(strTotalGrAmt)
                    oUsertable.UserFields.Fields.Item("U_Z_Currency").Value = strCurrency

                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub ReadGIIVLG(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading GIIVLG files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strInvoiceNo, strLineNo, strGoodsBranch, strGoodsNo, strVatKey, strSupplierDocNo, strAccount, strAccountTxt, strVattxt, strnetValue, strvatpercentage, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strvalue, strPaydate, strPeriodPay, strsuppliertype As String

                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strOrigin = strLIneStrin.GetValue(2)
                    strInvoiceNo = strLIneStrin.GetValue(3)
                    strLineNo = strLIneStrin.GetValue(4)
                    strGoodsBranch = strLIneStrin.GetValue(5)
                    strgoodsno = strLIneStrin.GetValue(6)
                    strVatKey = strLIneStrin.GetValue(7)
                    strSupplierDocNo = strLIneStrin.GetValue(8)
                    strBrand = strLIneStrin.GetValue(9)
                    strAccount = strLIneStrin.GetValue(10)
                    strAccountTxt = strLIneStrin.GetValue(11)
                    strnetValue = strLIneStrin.GetValue(12)

                    'strdate = strGoodsDate
                    'If strdate <> "" Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtMfrDate = GetDateTimeValue(DATE1)
                    'End If
                    'strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@GIIVLG", "Code")
                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from [@Z_GIIVLG] where U_Z_InvoiceNo=" & strInvoiceNo & " and U_Z_LineNo=" & strLineNo)
                    If oTest.RecordCount > 0 Then
                        Exit Sub
                    End If
                    oUsertable = oApplication.Company.UserTables.Item("GIIVLG")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Originator").Value = strOrigin
                    oUsertable.UserFields.Fields.Item("U_Z_InvoiceNo").Value = strInvoiceNo
                    oUsertable.UserFields.Fields.Item("U_Z_LineNo").Value = strLineNo
                    oUsertable.UserFields.Fields.Item("U_Z_GoodsBranch").Value = strGoodsBranch
                    oUsertable.UserFields.Fields.Item("U_Z_GoodsNo").Value = strgoodsno
                    oUsertable.UserFields.Fields.Item("U_Z_VatKey").Value = strVatKey
                    oUsertable.UserFields.Fields.Item("U_Z_SupplierDocNo").Value = strSupDocNo
                    oUsertable.UserFields.Fields.Item("U_Z_Brand").Value = strBrand
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    oUsertable.UserFields.Fields.Item("U_Z_Accounttxt").Value = strAccountTxt
                    oUsertable.UserFields.Fields.Item("U_Z_NetValue").Value = oApplication.Utilities.getDocumentQuantity(strnetValue)

                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub ReadGIIVLC(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading GIIVLG files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strInvoiceNo, strLineNo, strGoodsBranch, strGoodsNo, strVatKey, strSupplierDocNo, strAccount, strAccountTxt, strVattxt, strnetValue, strvatpercentage, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strvalue, strPaydate, strPeriodPay, strsuppliertype As String

                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strOrigin = strLIneStrin.GetValue(2)
                    strInvoiceNo = strLIneStrin.GetValue(3)
                    strLineNo = strLIneStrin.GetValue(4)
                    strGoodsBranch = strLIneStrin.GetValue(5)
                    strgoodsno = strLIneStrin.GetValue(6)
                    strVatKey = strLIneStrin.GetValue(7)
                    strSupplierDocNo = strLIneStrin.GetValue(8)
                    strBrand = strLIneStrin.GetValue(9)
                    strAccount = strLIneStrin.GetValue(10)
                    strAccountTxt = strLIneStrin.GetValue(11)
                    strnetValue = strLIneStrin.GetValue(12)

                    'strdate = strGoodsDate
                    'If strdate <> "" Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtMfrDate = GetDateTimeValue(DATE1)
                    'End If
                    'strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@GIIVLC", "Code")
                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from [@Z_GIIVLC] where U_Z_InvoiceNo=" & strInvoiceNo & " and U_Z_LineNo=" & strLineNo)
                    If oTest.RecordCount > 0 Then
                        Exit Sub
                    End If
                    oUsertable = oApplication.Company.UserTables.Item("GIIVLC")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Originator").Value = strOrigin
                    oUsertable.UserFields.Fields.Item("U_Z_InvoiceNo").Value = strInvoiceNo
                    oUsertable.UserFields.Fields.Item("U_Z_LineNo").Value = strLineNo
                    oUsertable.UserFields.Fields.Item("U_Z_GoodsBranch").Value = strGoodsBranch
                    oUsertable.UserFields.Fields.Item("U_Z_GoodsNo").Value = strgoodsno
                    oUsertable.UserFields.Fields.Item("U_Z_VatKey").Value = strVatKey
                    oUsertable.UserFields.Fields.Item("U_Z_SupplierDocNo").Value = strSupDocNo
                    oUsertable.UserFields.Fields.Item("U_Z_Brand").Value = strBrand
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    oUsertable.UserFields.Fields.Item("U_Z_Accounttxt").Value = strAccountTxt
                    oUsertable.UserFields.Fields.Item("U_Z_NetValue").Value = oApplication.Utilities.getDocumentQuantity(strnetValue)
                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Sub ReadCURIV(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading CURIV files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String

                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strOrigin = strLIneStrin.GetValue(2)
                    strDeliveryNoteNo = strLIneStrin.GetValue(3)
                    strVatKey = strLIneStrin.GetValue(4)
                    strSupplier = strLIneStrin.GetValue(5)
                    strBrand = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strDeliveryNoteDt = strLIneStrin.GetValue(8)
                    strDeliveryNoteNo = strLIneStrin.GetValue(9)
                    strFromBranch = strLIneStrin.GetValue(10)
                    strToCustomer = strLIneStrin.GetValue(11)
                    strComments = strLIneStrin.GetValue(12)
                    strQuantity = strLIneStrin.GetValue(13)
                    strDelNotePrice = strLIneStrin.GetValue(14)
                    strSalesPriceNet = strLIneStrin.GetValue(15)
                    strSalesPriceVat = strLIneStrin.GetValue(16)
                    stroSalesPriceNet = strLIneStrin.GetValue(17)
                    strOSalesPriceVAT = strLIneStrin.GetValue(18)
                    strperrateofVAT = strLIneStrin.GetValue(19)
                    strsuppliertype = strLIneStrin.GetValue(10)
                    strdate = strDeliveryNoteDt
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDeliveryNoteDt = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@CUSIV", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("CUSIV")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Originator").Value = strOrigin
                    oUsertable.UserFields.Fields.Item("U_Z_InvoiceNo").Value = oApplication.Utilities.getDocumentQuantity(strDeliveryNoteNo)
                    oUsertable.UserFields.Fields.Item("U_Z_VatKey").Value = strVatKey
                    oUsertable.UserFields.Fields.Item("U_Z_Supplier").Value = strSupplier
                    oUsertable.UserFields.Fields.Item("U_Z_Brand").Value = strBrand

                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    If strDeliveryNoteDt <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_InvDate").Value = dtMfrDate
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_DeliveryNote").Value = strDeliveryNoteNo

                    oUsertable.UserFields.Fields.Item("U_Z_FromBranch").Value = strFromBranch
                    oUsertable.UserFields.Fields.Item("U_Z_ToCustomer").Value = strToCustomer
                    oUsertable.UserFields.Fields.Item("U_Z_Comment").Value = strComments
                    oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = oApplication.Utilities.getDocumentQuantity(strQuantity)
                    oUsertable.UserFields.Fields.Item("U_Z_DelNotePrice").Value = oApplication.Utilities.getDocumentQuantity(strDelNotePrice)

                    oUsertable.UserFields.Fields.Item("U_Z_SalePriceNet").Value = oApplication.Utilities.getDocumentQuantity(strSalesPriceNet)
                    oUsertable.UserFields.Fields.Item("U_Z_SalePriceVat").Value = oApplication.Utilities.getDocumentQuantity(strSalesPriceVat)
                    oUsertable.UserFields.Fields.Item("U_Z_OsalePriceNet").Value = oApplication.Utilities.getDocumentQuantity(stroSalesPriceNet)
                    oUsertable.UserFields.Fields.Item("U_Z_OSalePriceVat").Value = oApplication.Utilities.getDocumentQuantity(strOSalesPriceVAT)

                    oUsertable.UserFields.Fields.Item("U_Z_PerRateofVAT").Value = oApplication.Utilities.getDocumentQuantity(strperrateofVAT)
                    oUsertable.UserFields.Fields.Item("U_Z_CustomerType").Value = oApplication.Utilities.getDocumentQuantity(strsuppliertype)
                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readSTKCOR(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading STKCOR files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String
                    Dim strDocID, strBranch, strreason, strreasonTxt, strPurcostprice, strSalesPrice, strperratevat, strDateCorre As String
                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strBranch = strLIneStrin.GetValue(2)
                    strDocID = strLIneStrin.GetValue(3)
                    strVatKey = strLIneStrin.GetValue(4)
                    strSupplier = strLIneStrin.GetValue(5)
                    strBrand = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strQuantity = strLIneStrin.GetValue(8)
                    strreason = strLIneStrin.GetValue(9)
                    strreasonTxt = strLIneStrin.GetValue(10)
                    strPurcostprice = strLIneStrin.GetValue(11)
                    strSalesPrice = strLIneStrin.GetValue(12)
                    strperratevat = strLIneStrin.GetValue(13)
                    strDateCorre = strLIneStrin.GetValue(14)
                    
                    strdate = strDateCorre
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDateCorre = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@STKCOR", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("STKCOR")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Branch").Value = strBranch
                    oUsertable.UserFields.Fields.Item("U_Z_DocID").Value = strDocID
                    oUsertable.UserFields.Fields.Item("U_Z_VatKey").Value = strVatKey
                    oUsertable.UserFields.Fields.Item("U_Z_Supplier").Value = strSupplier
                    oUsertable.UserFields.Fields.Item("U_Z_Brand").Value = strBrand
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    If strDateCorre <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_DateCorre").Value = dtMfrDate
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_Quantity)").Value = oApplication.Utilities.getDocumentQuantity(strQuantity)
                    oUsertable.UserFields.Fields.Item("U_Z_Reason").Value = oApplication.Utilities.getDocumentQuantity(strreason)
                    oUsertable.UserFields.Fields.Item("U_Z_Reasontxt").Value = oApplication.Utilities.getDocumentQuantity(strreasonTxt)
                    oUsertable.UserFields.Fields.Item("U_Z_PurCostPrice").Value = oApplication.Utilities.getDocumentQuantity(strPurcostprice)
                    oUsertable.UserFields.Fields.Item("U_Z_SalesPrice").Value = oApplication.Utilities.getDocumentQuantity(strSalesPrice)
                    oUsertable.UserFields.Fields.Item("U_Z_PerRateVAT").Value = oApplication.Utilities.getDocumentQuantity(strperratevat)

                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readSTKTAK(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading STKTAK files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String
                    Dim strDocID, strInventoryNo, strText, strExpQuatnity, strInvDiff, strProQuantity, strBranch, strreason, strreasonTxt, strPurcostprice, strSalesPrice, strperratevat, strDateCorre As String
                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strInventoryNo = strLIneStrin.GetValue(2)
                    strVatKey = strLIneStrin.GetValue(3)
                    strBranch = strLIneStrin.GetValue(4)
                    strSupplier = strLIneStrin.GetValue(5)
                    strBrand = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strDateCorre = strLIneStrin.GetValue(8)
                    strText = strLIneStrin.GetValue(9)
                    strExpQuatnity = strLIneStrin.GetValue(10)
                    strInvDiff = strLIneStrin.GetValue(11)
                    strProQuantity = strLIneStrin.GetValue(12)
                    strPurcostprice = strLIneStrin.GetValue(13)
                    strSalesPrice = strLIneStrin.GetValue(14)
                    strperratevat = strLIneStrin.GetValue(15)
                    strdate = strDateCorre
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDateCorre = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@STKTAK", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("STKTAK")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_InventoryNo").Value = oApplication.Utilities.getDocumentQuantity(strInventoryNo)
                    oUsertable.UserFields.Fields.Item("U_Z_VatKey").Value = strVatKey
                    oUsertable.UserFields.Fields.Item("U_Z_Branch").Value = strBranch
                    oUsertable.UserFields.Fields.Item("U_Z_Supplier").Value = strSupplier
                    oUsertable.UserFields.Fields.Item("U_Z_Brand").Value = strBrand
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    If strDateCorre <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_Date").Value = dtMfrDate
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_Text").Value = strText
                    oUsertable.UserFields.Fields.Item("U_Z_ExpQuantity)").Value = oApplication.Utilities.getDocumentQuantity(strExpQuatnity)
                    oUsertable.UserFields.Fields.Item("U_Z_InvDiff").Value = oApplication.Utilities.getDocumentQuantity(strInvDiff)
                    oUsertable.UserFields.Fields.Item("U_Z_ProQuantity").Value = oApplication.Utilities.getDocumentQuantity(strProQuantity)
                    oUsertable.UserFields.Fields.Item("U_Z_PurCostPrice").Value = oApplication.Utilities.getDocumentQuantity(strPurcostprice)
                    oUsertable.UserFields.Fields.Item("U_Z_SalesPrice").Value = oApplication.Utilities.getDocumentQuantity(strSalesPrice)
                    oUsertable.UserFields.Fields.Item("U_Z_PerRateVAT").Value = oApplication.Utilities.getDocumentQuantity(strperratevat)

                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readSLSTRN(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading SLSTRN files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(vbTab)
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String
                    Dim strDocID, strInventoryNo, strText, strExpQuatnity, strInvDiff, strProQuantity, strBranch, strreason, strreasonTxt, strPurcostprice, strSalesPrice, strperratevat, strDateCorre As String
                    Dim strReportNo, strCashNo, strSales, strDiscount, strCostprice As String
                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strDateCorre = strLIneStrin.GetValue(2)
                    strBranch = strLIneStrin.GetValue(3)
                    strReportNo = strLIneStrin.GetValue(4)
                    strcashno = strLIneStrin.GetValue(5)
                    strVatKey = strLIneStrin.GetValue(6)
                    strSupplier = strLIneStrin.GetValue(7)
                    strBrand = strLIneStrin.GetValue(8)
                    strAccount = strLIneStrin.GetValue(9)
                    strSales = strLIneStrin.GetValue(10)
                    strDiscount = strLIneStrin.GetValue(11)
                    strvat = strLIneStrin.GetValue(12)
                    strCostprice = strLIneStrin.GetValue(13)
                    strQuantity = strLIneStrin.GetValue(14)
                    strperratevat = strLIneStrin.GetValue(15)
                    strCurrency = strLIneStrin.GetValue(16)
                    strdate = strDateCorre
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDateCorre = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from [@SLSTRN] where U_Z_Imported='Y' and U_Z_Branch='" & strBranch & "' and  U_Z_CompanyCode='" & strCompCode & "' and  U_Z_CashNo= " & strCashNo & " and  U_Z_ReportNo=" & strReportNo & " and U_Z_Brand='" & strBrand & "'")
                    If oTest.RecordCount > 0 Then
                        Exit Sub
                    Else
                        oTest.DoQuery("Delete from [@SLSTRN] where U_Z_Imported<>'Y' and U_Z_Branch='" & strBranch & "' and  U_Z_CompanyCode='" & strCompCode & "' and  U_Z_CashNo= " & strCashNo & " and  U_Z_ReportNo=" & strReportNo & " and U_Z_Brand='" & strBrand & "'")
                    End If 'End If

                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@SLSTRN", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("SLSTRN")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_ReportNo").Value = oApplication.Utilities.getDocumentQuantity(strReportNo)
                    oUsertable.UserFields.Fields.Item("U_Z_CashNo").Value = oApplication.Utilities.getDocumentQuantity(strCashNo)
                    oUsertable.UserFields.Fields.Item("U_Z_VatKey").Value = strVatKey
                    oUsertable.UserFields.Fields.Item("U_Z_Branch").Value = strBranch
                    oUsertable.UserFields.Fields.Item("U_Z_Supplier").Value = strSupplier
                    oUsertable.UserFields.Fields.Item("U_Z_Brand").Value = strBrand
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    If strDateCorre <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_Date").Value = dtMfrDate
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_Sales").Value = oApplication.Utilities.getDocumentQuantity(strSales)
                    oUsertable.UserFields.Fields.Item("U_Z_Discount").Value = oApplication.Utilities.getDocumentQuantity(strDiscount)
                    ' oUsertable.UserFields.Fields.Item("U_Z_VAT").Value = oApplication.Utilities.getDocumentQuantity(strVat)
                    '  oUsertable.UserFields.Fields.Item("U_Z_CostPrice").Value = oApplication.Utilities.getDocumentQuantity(strCostprice)
                    oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = oApplication.Utilities.getDocumentQuantity(strQuantity)
                    oUsertable.UserFields.Fields.Item("U_Z_Currency").Value = strCurrency
                    oUsertable.UserFields.Fields.Item("U_Z_PerRateVat").Value = oApplication.Utilities.getDocumentQuantity(strperratevat)

                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ' Return False
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readSLSPAY(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading SLSPAY files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(vbTab)
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String
                    Dim strDocID, strInventoryNo, strText, strExpQuatnity, strInvDiff, strProQuantity, strBranch, strreason, strreasonTxt, strPurcostprice, strSalesPrice, strperratevat, strDateCorre As String
                    Dim strReportNo, strCashNo, strpaytype, strTurnOver, strTurnFrgCur, strFrgCurrency, strSales, strDiscount, strCostprice As String
                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strDateCorre = strLIneStrin.GetValue(2)
                    strBranch = strLIneStrin.GetValue(3)
                    strReportNo = strLIneStrin.GetValue(4)
                    strCashNo = strLIneStrin.GetValue(5)
                    strPaytype = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strCurrency = strLIneStrin.GetValue(8)
                    strTurnOver = strLIneStrin.GetValue(9)
                    strTurnFrgCur = strLIneStrin.GetValue(10)
                    strFrgCurrency = strLIneStrin.GetValue(11)
                   
                    strdate = strDateCorre
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDateCorre = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from [@SLSPAY] where U_Z_Imported='Y' and U_Z_Branch='" & strBranch & "' and  U_Z_CompanyCode='" & strCompCode & "' and  U_Z_CashNo= " & strCashNo & " and  U_Z_ReportNo=" & strReportNo)
                    If oTest.RecordCount > 0 Then
                        Exit Sub
                    Else
                        oTest.DoQuery("Delete from [@SLSPAY] where U_Z_Imported<>'Y' and U_Z_Branch='" & strBranch & "' and  U_Z_CompanyCode='" & strCompCode & "' and  U_Z_CashNo= " & strCashNo & " and  U_Z_ReportNo=" & strReportNo)
                    End If 'End If


                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@SLSPAY", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("SLSPAY")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Branch").Value = strBranch
                    oUsertable.UserFields.Fields.Item("U_Z_ReportNo").Value = oApplication.Utilities.getDocumentQuantity(strReportNo)
                    oUsertable.UserFields.Fields.Item("U_Z_CashNo").Value = oApplication.Utilities.getDocumentQuantity(strCashNo)
                    oUsertable.UserFields.Fields.Item("U_Z_PayType").Value = strpaytype
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    oUsertable.UserFields.Fields.Item("U_Z_Currency").Value = strCurrency
                    oUsertable.UserFields.Fields.Item("U_Z_TurnOver").Value = oApplication.Utilities.getDocumentQuantity(strTurnOver)
                    If strDateCorre <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_Date").Value = dtMfrDate
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_TurnFrgCur").Value = oApplication.Utilities.getDocumentQuantity(strTurnFrgCur)
                    oUsertable.UserFields.Fields.Item("U_Z_FrgCurrency").Value = strFrgCurrency

                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Sub readSLSDIF(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading SLSDIF files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(vbTab)
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String
                    Dim strDocID, strInventoryNo, strText, strExpQuatnity, strInvDiff, strProQuantity, strBranch, strreason, strreasonTxt, strPurcostprice, strSalesPrice, strperratevat, strDateCorre As String
                    Dim strReportNo, strCashNo, strpaytype, strDifference, strDiffFrgCur, strTurnOver, strTurnFrgCur, strFrgCurrency, strSales, strDiscount, strCostprice As String
                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strDateCorre = strLIneStrin.GetValue(2)
                    strBranch = strLIneStrin.GetValue(3)
                    strReportNo = strLIneStrin.GetValue(4)
                    strCashNo = strLIneStrin.GetValue(5)
                    strpaytype = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strCurrency = strLIneStrin.GetValue(8)
                    strDifference = strLIneStrin.GetValue(9)
                    strDiffFrgCur = strLIneStrin.GetValue(10)
                    strFrgCurrency = strLIneStrin.GetValue(11)

                    strdate = strDateCorre
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDateCorre = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from [@SLSDIF] where U_Z_Imported='Y' and U_Z_Branch='" & strBranch & "' and  U_Z_CompanyCode='" & strCompCode & "' and  U_Z_CashNo= " & strCashNo & " and  U_Z_ReportNo=" & strReportNo)
                    If oTest.RecordCount > 0 Then
                        Exit Sub
                    Else
                        oTest.DoQuery("Delete from [@SLSDIF] where U_Z_Imported<>'Y' and U_Z_Branch='" & strBranch & "' and  U_Z_CompanyCode='" & strCompCode & "' and  U_Z_CashNo= " & strCashNo & " and  U_Z_ReportNo=" & strReportNo)
                    End If 'End If
                    'End If


                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@SLSDIF", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("SLSDIF")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Branch").Value = strBranch
                    oUsertable.UserFields.Fields.Item("U_Z_ReportNo").Value = oApplication.Utilities.getDocumentQuantity(strReportNo)
                    oUsertable.UserFields.Fields.Item("U_Z_CashNo").Value = oApplication.Utilities.getDocumentQuantity(strCashNo)
                    oUsertable.UserFields.Fields.Item("U_Z_PayType").Value = strpaytype
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    oUsertable.UserFields.Fields.Item("U_Z_Currency").Value = strCurrency
                    oUsertable.UserFields.Fields.Item("U_Z_Difference").Value = oApplication.Utilities.getDocumentQuantity(strDifference)
                    If strDateCorre <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_Date").Value = dtMfrDate
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_DiffFrgCur").Value = oApplication.Utilities.getDocumentQuantity(strDiffFrgCur)
                    oUsertable.UserFields.Fields.Item("U_Z_FrgCurrency").Value = strFrgCurrency

                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readSLSDRP(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading SLSDRP files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String
                    Dim strDocID, strInventoryNo, strText, strExpQuatnity, strInvDiff, strProQuantity, strBranch, strreason, strreasonTxt, strPurcostprice, strSalesPrice, strperratevat, strDateCorre As String
                    Dim strReportNo, strCashNo, strpaytype, strDifference, strDiffFrgCur, strTurnOver, strTurnFrgCur, strFrgCurrency, strSales, strDiscount, strCostprice As String
                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strDateCorre = strLIneStrin.GetValue(2)
                    strBranch = strLIneStrin.GetValue(3)
                    strReportNo = strLIneStrin.GetValue(4)
                    strCashNo = strLIneStrin.GetValue(5)
                    strpaytype = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strCurrency = strLIneStrin.GetValue(8)
                    strDifference = strLIneStrin.GetValue(9)
                    strDiffFrgCur = strLIneStrin.GetValue(10)
                    strFrgCurrency = strLIneStrin.GetValue(11)

                    strdate = strDateCorre
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDateCorre = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@SLSDRP", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("SLSDRP")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Branch").Value = strBranch
                    oUsertable.UserFields.Fields.Item("U_Z_ReportNo").Value = oApplication.Utilities.getDocumentQuantity(strReportNo)
                    oUsertable.UserFields.Fields.Item("U_Z_CashNo").Value = oApplication.Utilities.getDocumentQuantity(strCashNo)
                    oUsertable.UserFields.Fields.Item("U_Z_PayType").Value = strpaytype
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    oUsertable.UserFields.Fields.Item("U_Z_Currency").Value = strCurrency
                    oUsertable.UserFields.Fields.Item("U_Z_DropAmt").Value = oApplication.Utilities.getDocumentQuantity(strDifference)
                    If strDateCorre <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_Date").Value = dtMfrDate
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_DropFrgCur").Value = oApplication.Utilities.getDocumentQuantity(strDiffFrgCur)
                    oUsertable.UserFields.Fields.Item("U_Z_FrgCurrency").Value = strFrgCurrency

                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readSLSPUP(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading SLSPUP files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String
                    Dim strDocID, strInventoryNo, strText, strExpQuatnity, strInvDiff, strProQuantity, strBranch, strreason, strreasonTxt, strPurcostprice, strSalesPrice, strperratevat, strDateCorre As String
                    Dim strReportNo, strCashNo, strpaytype, strDifference, strDiffFrgCur, strTurnOver, strTurnFrgCur, strFrgCurrency, strSales, strDiscount, strCostprice As String
                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strDateCorre = strLIneStrin.GetValue(2)
                    strBranch = strLIneStrin.GetValue(3)
                    strReportNo = strLIneStrin.GetValue(4)
                    strCashNo = strLIneStrin.GetValue(5)
                    strpaytype = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strCurrency = strLIneStrin.GetValue(8)
                    strDifference = strLIneStrin.GetValue(9)
                    strDiffFrgCur = strLIneStrin.GetValue(10)
                    strFrgCurrency = strLIneStrin.GetValue(11)

                    strdate = strDateCorre
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDateCorre = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@SLSPUP", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("SLSPUP")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Branch").Value = strBranch
                    oUsertable.UserFields.Fields.Item("U_Z_ReportNo").Value = oApplication.Utilities.getDocumentQuantity(strReportNo)
                    oUsertable.UserFields.Fields.Item("U_Z_CashNo").Value = oApplication.Utilities.getDocumentQuantity(strCashNo)
                    oUsertable.UserFields.Fields.Item("U_Z_PayType").Value = strpaytype
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    oUsertable.UserFields.Fields.Item("U_Z_Currency").Value = strCurrency
                    oUsertable.UserFields.Fields.Item("U_Z_PickAmt").Value = oApplication.Utilities.getDocumentQuantity(strDifference)
                    If strDateCorre <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_Date").Value = dtMfrDate
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_PickFrgCur").Value = oApplication.Utilities.getDocumentQuantity(strDiffFrgCur)
                    oUsertable.UserFields.Fields.Item("U_Z_FrgCurrency").Value = strFrgCurrency

                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readSLSGFT(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading SLSGFT files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String
                    Dim strDocID, strInventoryNo, strText, strExpQuatnity, strInvDiff, strProQuantity, strBranch, strreason, strreasonTxt, strPurcostprice, strSalesPrice, strperratevat, strDateCorre As String
                    Dim strReportNo, strCashNo, strpaytype, strDifference, strDiffFrgCur, strTurnOver, strTurnFrgCur, strFrgCurrency, strSales, strDiscount, strCostprice As String
                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strDateCorre = strLIneStrin.GetValue(2)
                    strBranch = strLIneStrin.GetValue(3)
                    strReportNo = strLIneStrin.GetValue(4)
                    strCashNo = strLIneStrin.GetValue(5)
                    strpaytype = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strCurrency = strLIneStrin.GetValue(8)
                    strDifference = strLIneStrin.GetValue(9)
                    strDiffFrgCur = strLIneStrin.GetValue(10)
                    strFrgCurrency = strLIneStrin.GetValue(11)

                    strdate = strDateCorre
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDateCorre = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@SLSGFT", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("SLSGFT")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Branch").Value = strBranch
                    oUsertable.UserFields.Fields.Item("U_Z_ReportNo").Value = oApplication.Utilities.getDocumentQuantity(strReportNo)
                    oUsertable.UserFields.Fields.Item("U_Z_CashNo").Value = oApplication.Utilities.getDocumentQuantity(strCashNo)
                    oUsertable.UserFields.Fields.Item("U_Z_PayType").Value = strpaytype
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    oUsertable.UserFields.Fields.Item("U_Z_Currency").Value = strCurrency
                    oUsertable.UserFields.Fields.Item("U_Z_GiftAmt").Value = oApplication.Utilities.getDocumentQuantity(strDifference)
                    If strDateCorre <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_Date").Value = dtMfrDate
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_GiftFrgCur").Value = oApplication.Utilities.getDocumentQuantity(strDiffFrgCur)
                    oUsertable.UserFields.Fields.Item("U_Z_FrgCurrency").Value = strFrgCurrency

                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readSLSPRE(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading SLSPRE files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String
                    Dim strDocID, strInventoryNo, strText, strExpQuatnity, strInvDiff, strProQuantity, strBranch, strreason, strreasonTxt, strPurcostprice, strSalesPrice, strperratevat, strDateCorre As String
                    Dim strReportNo, strCashNo, strpaytype, strDifference, strDiffFrgCur, strTurnOver, strTurnFrgCur, strFrgCurrency, strSales, strDiscount, strCostprice As String
                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strDateCorre = strLIneStrin.GetValue(2)
                    strBranch = strLIneStrin.GetValue(3)
                    strReportNo = strLIneStrin.GetValue(4)
                    strCashNo = strLIneStrin.GetValue(5)
                    strpaytype = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strCurrency = strLIneStrin.GetValue(8)
                    strDifference = strLIneStrin.GetValue(9)
                    'strDiffFrgCur = strLIneStrin.GetValue(10)
                    'strFrgCurrency = strLIneStrin.GetValue(11)

                    strdate = strDateCorre
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDateCorre = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@SLSPRE", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("SLSPRE")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Branch").Value = strBranch
                    oUsertable.UserFields.Fields.Item("U_Z_ReportNo").Value = oApplication.Utilities.getDocumentQuantity(strReportNo)
                    oUsertable.UserFields.Fields.Item("U_Z_CashNo").Value = oApplication.Utilities.getDocumentQuantity(strCashNo)
                    oUsertable.UserFields.Fields.Item("U_Z_PayType").Value = strpaytype
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    oUsertable.UserFields.Fields.Item("U_Z_Currency").Value = strCurrency
                    oUsertable.UserFields.Fields.Item("U_Z_PrePayAmt").Value = oApplication.Utilities.getDocumentQuantity(strDifference)
                    If strDateCorre <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_Date").Value = dtMfrDate
                    End If
                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readSLSEXP(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading SLSEXP files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String
                    Dim strDocID, strInventoryNo, strText, strExpQuatnity, strInvDiff, strProQuantity, strBranch, strreason, strreasonTxt, strPurcostprice, strSalesPrice, strperratevat, strDateCorre As String
                    Dim strReportNo, strCashNo, strpaytype, strDifference, strDiffFrgCur, strTurnOver, strTurnFrgCur, strFrgCurrency, strSales, strDiscount, strCostprice As String
                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strDateCorre = strLIneStrin.GetValue(2)
                    strBranch = strLIneStrin.GetValue(3)
                    strReportNo = strLIneStrin.GetValue(4)
                    strCashNo = strLIneStrin.GetValue(5)
                    strpaytype = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strCurrency = strLIneStrin.GetValue(8)
                    strDifference = strLIneStrin.GetValue(9)
                    'strDiffFrgCur = strLIneStrin.GetValue(10)
                    'strFrgCurrency = strLIneStrin.GetValue(11)

                    strdate = strDateCorre
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDateCorre = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@SLSEXP", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("SLSEXP")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Branch").Value = strBranch
                    oUsertable.UserFields.Fields.Item("U_Z_ReportNo").Value = oApplication.Utilities.getDocumentQuantity(strReportNo)
                    oUsertable.UserFields.Fields.Item("U_Z_CashNo").Value = oApplication.Utilities.getDocumentQuantity(strCashNo)
                    oUsertable.UserFields.Fields.Item("U_Z_PayType").Value = strpaytype
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    oUsertable.UserFields.Fields.Item("U_Z_Currency").Value = strCurrency
                    oUsertable.UserFields.Fields.Item("U_Z_ExpAmt").Value = oApplication.Utilities.getDocumentQuantity(strDifference)
                    If strDateCorre <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_Date").Value = dtMfrDate
                    End If
                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub readCREBAL(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading CREBAL files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strgoodsno, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strGoodsBranch, strvalue, strPaydate, strPeriodPay, strvatpercentage, strsuppliertype As String
                    Dim strDeliveryNoteNo, strToCustomer, strperrateofVAT, strVatKey, strToSupplier, strAccount, strDeliveryNoteDt, strFromBranch, strComments, strQuantity, strDelNotePrice, strSalesPriceNet, strSalesPriceVat, stroSalesPriceNet, strOSalesPriceVAT As String
                    Dim strDocID, strInventoryNo, strText, strExpQuatnity, strInvDiff, strProQuantity, strBranch, strreason, strreasonTxt, strPurcostprice, strSalesPrice, strperratevat, strDateCorre As String
                    Dim strReportNo, strCashNo, strCustomer, strCustomerType, strRefDoc, strBalamt, strpaytype, strDifference, strDiffFrgCur, strTurnOver, strTurnFrgCur, strFrgCurrency, strSales, strDiscount, strCostprice As String
                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strDateCorre = strLIneStrin.GetValue(2)
                    strBranch = strLIneStrin.GetValue(3)
                    strReportNo = strLIneStrin.GetValue(4)
                    strCashNo = strLIneStrin.GetValue(5)
                    strpaytype = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strCustomer = strLIneStrin.GetValue(8)
                    strCustomerType = strLIneStrin.GetValue(9)
                    strRefDoc = strLIneStrin.GetValue(10)
                    strCurrency = strLIneStrin.GetValue(11)
                    strBalamt = strLIneStrin.GetValue(12)

                    strdate = strDateCorre
                    If strdate <> "" And strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDateCorre = ""
                    End If
                    '  strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@CREBAL", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("CREBAL")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Branch").Value = strBranch
                    oUsertable.UserFields.Fields.Item("U_Z_ReportNo").Value = oApplication.Utilities.getDocumentQuantity(strReportNo)
                    oUsertable.UserFields.Fields.Item("U_Z_CashNo").Value = oApplication.Utilities.getDocumentQuantity(strCashNo)
                    oUsertable.UserFields.Fields.Item("U_Z_PayType").Value = strpaytype
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    oUsertable.UserFields.Fields.Item("U_Z_Customer").Value = strCustomer
                    oUsertable.UserFields.Fields.Item("U_Z_CustomerType").Value = oApplication.Utilities.getDocumentQuantity(strCustomerType)
                    oUsertable.UserFields.Fields.Item("U_Z_RefDoc").Value = strRefDoc
                    oUsertable.UserFields.Fields.Item("U_Z_Currency").Value = strCurrency
                    oUsertable.UserFields.Fields.Item("U_Z_BalAmt").Value = oApplication.Utilities.getDocumentQuantity(strBalamt)
                    If strDateCorre <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_Date").Value = dtMfrDate
                    End If
                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                ' oDelrec.DoQuery("Update [@Z_XADJ] set Name=code where name like '%M' and U_Z_Filename='" & fi.Name & "'")
                '  sr.Close()
                '  File.Move(fi.FullName, strSuccessFile)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", sPath)
                '   oApplication.Utilities.WriteErrorlog("Reading Process Completed: Filename-->" & fi.Name & " Moved to success folder", strImportErrorLog)

                'Return True
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub ReadIBTDN(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading IBTDN files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strInvoiceNo, strLineNo, strGoodsBranch, strGoodsNo, strVatKey, strSupplierDocNo, strAccount, strAccountTxt, strVattxt, strnetValue, strvatpercentage, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strvalue, strPaydate, strPeriodPay, strsuppliertype As String
                    Dim strDeliverynoteNo, strFromBranch, strDeliveryNoteDt, strtoBranch, strcomment, strQuantity, strDelNotePrice As String

                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strOrigin = strLIneStrin.GetValue(2)
                    strDeliverynoteNo = strLIneStrin.GetValue(3)
                    strFromBranch = strLIneStrin.GetValue(4)
                    strSupplier = strLIneStrin.GetValue(5)
                    strBrand = strLIneStrin.GetValue(6)
                    strAccount = strLIneStrin.GetValue(7)
                    strDeliveryNoteDt = strLIneStrin.GetValue(8)
                    strtoBranch = strLIneStrin.GetValue(9)
                    strcomment = strLIneStrin.GetValue(10)
                    strQuantity = strLIneStrin.GetValue(11)
                    strDelNotePrice = strLIneStrin.GetValue(12)

                    strdate = strDeliveryNoteDt
                    If strdate <> "" Or strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDeliveryNoteDt = ""
                    End If
                    'strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@IBTDN", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("IBTDN")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Originator").Value = strOrigin
                    oUsertable.UserFields.Fields.Item("U_Z_DeliveryNoteNo").Value = oApplication.Utilities.getDocumentQuantity(strDeliverynoteNo)
                    oUsertable.UserFields.Fields.Item("U_Z_FromBranch").Value = strFromBranch
                    oUsertable.UserFields.Fields.Item("U_Z_Supplier").Value = strSupplier
                    oUsertable.UserFields.Fields.Item("U_Z_Brand").Value = strBrand
                    oUsertable.UserFields.Fields.Item("U_Z_Account").Value = strAccount
                    If strDeliveryNoteDt <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_DeliveryNoteDt").Value = dtMfrDate
                    End If

                    oUsertable.UserFields.Fields.Item("U_Z_ToBranch").Value = strtoBranch
                    oUsertable.UserFields.Fields.Item("U_Z_Comment").Value = strcomment
                    oUsertable.UserFields.Fields.Item("U_Z_Quantity").Value = oApplication.Utilities.getDocumentQuantity(strQuantity)
                    oUsertable.UserFields.Fields.Item("U_Z_DelNotePrice").Value = oApplication.Utilities.getDocumentQuantity(strnetValue)


                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub ReadIBTCO(ByVal aString As String)
        '  Dim di As New IO.DirectoryInfo(aFolderpath)
        '  Dim aryFi As IO.FileInfo() = di.GetFiles("*.csv")
        Dim fi As IO.FileInfo
        Dim strStorekey, strsokey, strwhs, strImpDocType, strSuccessFile, strErrorFile, strsuccessfolder, strErrorfolder, strremarks, strType, strdate, strOrderKey, strShipdate, strSKU, strQty, strbatch, strmfgdate, strexpdate, strSusr1, strSur2, strholdcode As String
        Dim dtShipdate, dtMfrDate, dtExpDate As Date
        Dim sr As IO.StreamReader
        Dim YEAR, MONTH, DAY, DATE1, strFilename, linje As String
        Dim oDelrec As SAPbobsCOM.Recordset
        Try
            oApplication.Utilities.Message("Reading IBTDN files...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'oApplication.Utilities.WriteErrorlog("Reading ADJ Files processing..", strImportErrorLog)
            Dim strLIneStrin As String()
            Try
                Dim oRec, oRecUpdate, oTemp As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                linje = aString
                strLIneStrin = linje.Split(";")
                If 1 = 1 Then
                    Dim strType1, strCompCode, strOrigin, strInvoiceNo, strLineNo, strGoodsBranch, strGoodsNo, strVatKey, strSupplierDocNo, strAccount, strAccountTxt, strVattxt, strnetValue, strvatpercentage, strVat, strSupplier, strBrand, strGoodsDate, strSupDocNo, strCurrency, strvalue, strPaydate, strPeriodPay, strsuppliertype As String
                    Dim strDeliverynoteNo, strCashNo, strDateofCon, strFromBranch, strDeliveryNoteDt, strtoBranch, strcomment, strQuantity, strDelNotePrice As String

                    strType1 = strLIneStrin.GetValue(0)
                    strCompCode = strLIneStrin.GetValue(1)
                    strOrigin = strLIneStrin.GetValue(2)
                    strCashNo = strLIneStrin.GetValue(3)
                    strDeliverynoteNo = strLIneStrin.GetValue(4)
                    strFromBranch = strLIneStrin.GetValue(5)
                    strtoBranch = strLIneStrin.GetValue(6)
                    strcomment = strLIneStrin.GetValue(7)
                    strDateofCon = strLIneStrin.GetValue(8)
                    
                    strdate = strDateofCon
                    If strdate <> "" Or strdate.Length = 8 Then
                        DAY = strdate.Substring(6, 2)
                        MONTH = strdate.Substring(4, 2)
                        YEAR = strdate.Substring(0, 4)
                        DATE1 = DAY & MONTH & YEAR
                        dtMfrDate = GetDateTimeValue(DATE1)
                    Else
                        strDeliveryNoteDt = ""
                    End If
                    'strdate = strPaydate
                    'If strdate <> "" And strdate.Length = 8 Then
                    '    DAY = strdate.Substring(6, 2)
                    '    MONTH = strdate.Substring(4, 2)
                    '    YEAR = strdate.Substring(0, 4)
                    '    DATE1 = DAY & MONTH & YEAR
                    '    dtExpDate = GetDateTimeValue(DATE1)
                    'Else
                    '    strPaydate = ""
                    'End If
                    Dim oUsertable As SAPbobsCOM.UserTable
                    Dim strsql As String
                    strsql = oApplication.Utilities.getMaxCode("@IBTCO", "Code")
                    oUsertable = oApplication.Company.UserTables.Item("IBTCO")
                    oUsertable.Code = strsql
                    oUsertable.Name = strsql & "M"
                    oUsertable.UserFields.Fields.Item("U_Z_Type").Value = strType1
                    oUsertable.UserFields.Fields.Item("U_Z_CompanyCode").Value = strCompCode
                    oUsertable.UserFields.Fields.Item("U_Z_Originator").Value = strOrigin
                    oUsertable.UserFields.Fields.Item("U_Z_CashNo").Value = oApplication.Utilities.getDocumentQuantity(strCashNo)
                    oUsertable.UserFields.Fields.Item("U_Z_DeliveryNoteNo").Value = oApplication.Utilities.getDocumentQuantity(strDeliverynoteNo)
                    oUsertable.UserFields.Fields.Item("U_Z_FromBranch").Value = strFromBranch
                    If strDateofCon <> "" Then
                        oUsertable.UserFields.Fields.Item("U_Z_DateofConf").Value = dtMfrDate
                    End If
                    oUsertable.UserFields.Fields.Item("U_Z_ToBranch").Value = strtoBranch
                    oUsertable.UserFields.Fields.Item("U_Z_Comment").Value = strcomment
                    If oUsertable.Add <> 0 Then
                        MsgBox(oApplication.Company.GetLastErrorDescription)
                        oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oDelrec.DoQuery("Delete from [@Z_XADJ] where Name like '%M' and U_Z_Filename='" & fi.Name & "'")
                    End If
                End If
                oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region
#End Region

#End Region

#Region "GetDatetimevalue"
    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

#End Region




#End Region

#Region "Databind"
    Private Sub Databind()

        Dim strqry As String
        oGrid = oForm.Items.Item("17").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_0")
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Originator] 'Originator', T0.[U_Z_GoodsNo] 'Goods In Number', T0.[U_Z_VatKey] 'Vat Code', T0.[U_Z_Supplier] 'Supplier', T0.[U_Z_Brand] 'Goods In Brand', T0.[U_Z_AcctNo] 'Account Code', T0.[U_Z_GoodsDate] 'Goods In Date', T0.[U_Z_SupplierDocNo] 'Document Number', T0.[U_Z_Currency] 'Currency', T0.[U_Z_GoodsBranch] 'Goods in Branch', T0.[U_Z_Value] 'Value (in Cost / Purchase Price)', T0.[U_Z_PayDate] 'Payment Date' , T0.[U_Z_PeriodPay] 'Period to Pay', T0.[U_Z_VatPercentage] 'Vat % rate', T0.[U_Z_SupplierType] 'Supplier Type' FROM [dbo].[@GIDN]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)

        oGrid = oForm.Items.Item("18").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_1")
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Originator] 'Originator', T0.[U_Z_GoodsNo] 'Goods in Number', T0.[U_Z_Account] 'Account', T0.[U_Z_Accounttxt] 'Account Text', T0.[U_Z_VatKey] 'VatKey', T0.[U_Z_NetValue] 'Net Value', T0.[U_Z_VatPercentage] 'Vat % Rate', T0.[U_Z_SupplierType] 'Supplier Type' FROM [dbo].[@GIDNCO]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("18").Enabled = False

        oGrid = oForm.Items.Item("31").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_14")
        '  strqry = "Select * from [@SLSTRN]"
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Date] 'Date', T0.[U_Z_Branch] 'Branch', T0.[U_Z_ReportNo] 'Report Number', T0.[U_Z_CashNo] 'Cash Number', T0.[U_Z_VatKey] 'VAT Key', T0.[U_Z_Supplier] 'Supplier', T0.[U_Z_Brand] 'External Item Attribute', T0.[U_Z_Account] 'Account', T0.[U_Z_Sales] 'Sales', T0.[U_Z_Discount] 'Discount', T0.[U_Z_Vat] 'VAT', T0.[U_Z_CostPrice] 'Cost Price', T0.[U_Z_Quantity] 'Quantity', T0.[U_Z_PerRateVat] 'Rate of VAT %', T0.[U_Z_Currency] 'Currency'  FROM [dbo].[@SLSTRN]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("31").Enabled = False

        oGrid = oForm.Items.Item("32").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_15")
        ' strqry = "Select * from [@SLSPAY]"
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Date] 'Date', T0.[U_Z_Branch] 'Branch', T0.[U_Z_ReportNo] 'Report Number', T0.[U_Z_CashNo] 'Cash Number',  T0.[U_Z_PayType] 'Payment Type', T0.[U_Z_Account] 'Account', T0.[U_Z_Currency] 'Currency', T0.[U_Z_TurnOver] 'TurnOver', T0.[U_Z_TurnFrgCur] 'Turn Over Foreign Currency', T0.[U_Z_FrgCurrency]  'Foreign Currency' FROM [dbo].[@SLSPAY]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("32").Enabled = False

        oGrid = oForm.Items.Item("33").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_16")
        'strqry = "Select * from [@SLSDIF]"
        strqry = "SELECT  T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Date] 'Date', T0.[U_Z_Branch] 'Branch', T0.[U_Z_ReportNo] 'Report Number', T0.[U_Z_CashNo] 'Cash Number',  T0.[U_Z_PayType] 'Payment Type', T0.[U_Z_Account] 'Account', T0.[U_Z_Currency] 'Currency', T0.[U_Z_Difference] 'Difference', T0.[U_Z_DiffFrgCur] 'Difference Foreign Currency', T0.[U_Z_FrgCurrency] 'Foreign Currency' FROM [dbo].[@SLSDIF]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("33").Enabled = False

        Exit Sub


        oGrid = oForm.Items.Item("19").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_2")
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Originator] 'Originator', T0.[U_Z_InvoiceNo] 'Invoice Number', T0.[U_Z_Supplier] 'Supplier', T0.[U_Z_Brand] 'External Item Attribute', T0.[U_Z_SupplierInvNo] 'Supplier Invoice Number', T0.[U_Z_InvoiceDate] 'Invoice Date', T0.[U_Z_Currency] 'Currency', T0.[U_Z_TotalGrAmt]  'Total Gross Amount' FROM [dbo].[@GIIV]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("19").Enabled = False
        oGrid = oForm.Items.Item("20").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_3")
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Originator] 'Originator', T0.[U_Z_InvoiceNo] 'Invoice Number', T0.[U_Z_LineNo] 'Line Number', T0.[U_Z_GoodsBranch] 'Goods in Branch', T0.[U_Z_GoodsNo] 'Goods In Number', T0.[U_Z_VatKey] 'Vat Key', T0.[U_Z_SupplierDocNo] 'Supplier Document Number', T0.[U_Z_Brand] 'External Item Attribute', T0.[U_Z_Account] 'Account', T0.[U_Z_Accounttxt] 'Account Text', T0.[U_Z_NetValue] 'Net Value' FROM [dbo].[@GIIVLC]  T0 where isnull(U_Z_Imported,'N')='N' Union All"
        strqry = strqry & "  SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Originator] 'Originator', T0.[U_Z_InvoiceNo] 'Invoice Number', T0.[U_Z_LineNo] 'Line Number', T0.[U_Z_GoodsBranch] 'Goods in Branch', T0.[U_Z_GoodsNo] 'Goods In Number', T0.[U_Z_VatKey] 'Vat Key', T0.[U_Z_SupplierDocNo] 'Supplier Document Number', T0.[U_Z_Brand] 'External Item Attribute', T0.[U_Z_Account] 'Account', T0.[U_Z_Accounttxt] 'Account Text', T0.[U_Z_NetValue] 'Net Value' FROM [dbo].[@GIIVLG]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("20").Enabled = False
        oGrid = oForm.Items.Item("21").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_4")
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Originator] 'Originator',T0.[U_Z_DeliveryNoteNo] 'Delivery Note Number', T0.[U_Z_FromBranch] 'From Branch', T0.[U_Z_Supplier] 'Supplier', T0.[U_Z_Brand]  'External Item Attribute', T0.[U_Z_Account] 'Account', T0.[U_Z_DeliveryNoteDt] 'Delivery Note date', T0.[U_Z_ToBranch] 'To Branch', T0.[U_Z_Comment] 'Comment', T0.[U_Z_Quantity] 'Quantity', T0.[U_Z_DelNotePrice] 'Purchase Price' FROM [dbo].[@IBTDN]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("21").Enabled = False
        oGrid = oForm.Items.Item("22").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_5")
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Originator] 'Originator', T0.[U_Z_CashNo] 'Cash Number' ,T0.[U_Z_DeliveryNoteNo] 'Delivery Note Number', T0.[U_Z_FromBranch] 'From Branch', T0.[U_Z_ToBranch] 'To Branch', T0.[U_Z_Comment] 'Comment', T0.[U_Z_DateofConf] 'Date of Confirmation' FROM [dbo].[@IBTCO]  T0"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("22").Enabled = False
        ''Details2
        oGrid = oForm.Items.Item("23").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_6")
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Originator] 'Originator', T0.[U_Z_DeliveryNoteNo] 'Delivery Note Number', T0.[U_Z_VatKey] 'VAT Key', T0.[U_Z_ToSupplier] 'To Supplier', T0.[U_Z_Brand] 'External Item Attribute', T0.[U_Z_Account] 'Account', T0.[U_Z_DeliveryNoteDt] 'Delivery Note Date', T0.[U_Z_FromBranch] 'From Branch', T0.[U_Z_Comment] 'Comment', T0.[U_Z_Quantity] 'Quantity', T0.[U_Z_DelNotePrice] 'Delivery Note Purchase Price', T0.[U_Z_SalePriceNet] 'Sales Price Net', T0.[U_Z_SalePriceVat] 'Sales Price Incl VAT', T0.[U_Z_OSalePriceNet] 'Original Sales price net', T0.[U_Z_OSalePriceVat] 'Origional Sales Price Inc VAT', T0.[U_Z_PerRateofVAT] 'Vat % Rate', T0.[U_Z_SupplierType]  'Supplier Type' FROM [dbo].[@SURDN]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("23").Enabled = False
        oGrid = oForm.Items.Item("24").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_7")
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Originator] 'Originator', T0.[U_Z_DeliveryNoteNo] 'Delivery Note Number', T0.[U_Z_VatKey] 'VAT Key', T0.[U_Z_Supplier] 'Supplier', T0.[U_Z_Brand] 'External Item Attribute', T0.[U_Z_Account] 'Account', T0.[U_Z_DeliveryNoteDt] 'Delivery Note date', T0.[U_Z_FromBranch] 'From Branch', T0.[U_Z_ToCustomer] 'To Customer', T0.[U_Z_Comment] 'Comments', T0.[U_Z_Quantity] 'Quantity', T0.[U_Z_DelNotePrice] 'Delivery Note Purchase Price', T0.[U_Z_SalePriceNet] 'Sales Price Net', T0.[U_Z_SalePriceVat] 'Sales Price Incl VAT', T0.[U_Z_OSalePriceNet] 'Original Sales price net', T0.[U_Z_OSalePriceVat] 'Origional Sales Price Inc VAT', T0.[U_Z_PerRateofVAT] 'Vat % Rate',T0.[U_Z_CustomerType] 'Customer Type' FROM [dbo].[@CUSDN]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("24").Enabled = False
        oGrid = oForm.Items.Item("25").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_8")
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Originator] 'Originator', T0.[U_Z_InvoiceNo] 'Invoice Number', T0.[U_Z_InvDate] 'Invoice Date', T0.[U_Z_DeliveryNote] 'Delivery Note', T0.[U_Z_FromBranch] 'From Branch', T0.[U_Z_ToSupplier] 'To Supplier', T0.[U_Z_Comment]  'Comment' FROM [dbo].[@SURIV]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("25").Enabled = False
        oGrid = oForm.Items.Item("26").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_9")
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Originator] 'Originator', T0.[U_Z_InvoiceNo] 'Invoice Number', T0.[U_Z_Supplier] 'Supplier', T0.[U_Z_Brand] 'External Item Attribute', T0.[U_Z_Account] 'Account',  T0.[U_Z_VatKey] 'VAT Key', T0.[U_Z_Quantity] 'Quantity',T0.[U_Z_DelNotePrice] 'Delivery Note Purchase Price', T0.[U_Z_SalePriceNet] 'Sales Price Net', T0.[U_Z_SalePriceVat] 'Sales Price Incl VAT', T0.[U_Z_OSalePriceNet] 'Original Sales price net', T0.[U_Z_OSalePriceVat] 'Origional Sales Price Inc VAT', T0.[U_Z_PerRateofVAT] 'Vat % Rate' FROM [dbo].[@SURIVLG]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("26").Enabled = False
        oGrid = oForm.Items.Item("27").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_10")
        'strqry = "Select * from [@SURIVLC]"
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Originator] 'Originator', T0.[U_Z_InvoiceNo], 'Invoice Number',  T0.[U_Z_FreightPer] 'Freight Percentage', T0.[U_Z_FreightNet] 'Freight Net', T0.[U_Z_FreightVat] 'Freight Net', T0.[U_Z_TransInsPer] 'Transport Insurance  %', T0.[U_Z_TransInsNet] 'Transport Insurance Net', T0.[U_Z_TransInsVat] 'Transport Insurance VAT', T0.[U_Z_TransCostNet] 'Transport Cost Net',  T0.[U_Z_TransCostVat]  'Transport Cost VAT' FROM [dbo].[@SURIVLC]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("27").Enabled = False
        oGrid = oForm.Items.Item("28").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_11")
        ' strqry = "Select * from [@CUSIV]"
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Originator] 'Originator', T0.[U_Z_InvoiceNo] 'Invoice Number', T0.[U_Z_VatKey] 'VAT Key', T0.[U_Z_Supplier] ' Supplier', T0.[U_Z_Brand] 'External Item Attribute', T0.[U_Z_Account] 'Account', T0.[U_Z_InvDate] 'Invoice Date', T0.[U_Z_DeliveryNote] 'Delivery Note',  T0.[U_Z_FromBranch] 'From Branch',"
        strqry = strqry & " T0.[U_Z_ToCustomer] 'To Customer', T0.[U_Z_Comment] 'Comment', T0.[U_Z_Quantity] 'Quantity', T0.[U_Z_DelNotePrice] 'Delivery Note Purchase Price', T0.[U_Z_SalePriceNet] 'Sales Price Net', T0.[U_Z_SalePriceVat] 'Sales Price Incl VAT', T0.[U_Z_OSalePriceNet] 'Original Sales price net', T0.[U_Z_OSalePriceVat] 'Origional Sales Price Inc VAT', T0.[U_Z_PerRateofVAT] 'Vat % Rate', T0.[U_Z_CustomerType] FROM [dbo].[@CUSIV]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("28").Enabled = False
        ''Details3
        oGrid = oForm.Items.Item("29").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_12")
        'strqry = "Select * from [@STKCOR]"
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Branch] 'Branch', T0.[U_Z_DocId] 'Document ID', T0.[U_Z_VatKey] 'VAT Key', T0.[U_Z_Supplier] 'Supplier', T0.[U_Z_Brand] 'External Item Attribute', T0.[U_Z_Account] 'Account', T0.[U_Z_Quantity] 'Quantity', T0.[U_Z_Reason] 'reason', T0.[U_Z_Reasontxt] 'Reason text', T0.[U_Z_PurCostPrice] 'Purchase Cost Price ', T0.[U_Z_SalesPrice] 'Sales Price', T0.[U_Z_PerRateVat] 'Rat of VAT %', T0.[U_Z_DateCorre] 'Date of Correction ' FROM [dbo].[@STKCOR]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("29").Enabled = False
        oGrid = oForm.Items.Item("30").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_13")
        'strqry = "Select * from [@STKTAK]"
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_InventoryNo] 'Invoice Number', T0.[U_Z_VatKey] 'VAT Key', T0.[U_Z_Branch] 'Branch', T0.[U_Z_Supplier] 'Supplier', T0.[U_Z_Brand] 'Brand', T0.[U_Z_Account] 'Account', T0.[U_Z_Date] 'Date', T0.[U_Z_Text] 'Text', T0.[U_Z_ExpQuantity] 'Expected Quantity', T0.[U_Z_InvDiff] 'Inventory Difference', T0.[U_Z_ProQuantity] 'Processed Quantity', T0.[U_Z_PurCostPrice] 'Purchase Cost', T0.[U_Z_SalesPrice] 'Sales Price', T0.[U_Z_PerRateVat]  'Rate of VAT %' FROM [dbo].[@STKTAK]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("30").Enabled = False
       
        oGrid = oForm.Items.Item("34").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_17")
        ' strqry = "Select * from [@SLSDRP]"
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Date] 'Date', T0.[U_Z_Branch] 'Branch', T0.[U_Z_ReportNo] 'Report Number', T0.[U_Z_CashNo] 'Cash Number',  T0.[U_Z_PayType] 'Payment Type', T0.[U_Z_Account] 'Account', T0.[U_Z_Currency] 'Currency',T0.[U_Z_DropAmt] 'Drop off Amount', T0.[U_Z_FrgCurrency] 'Foreign Currency', T0.[U_Z_DropFrgCur]  'Drop off Amt.Forg.Currency' FROM [dbo].[@SLSDRP]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("34").Enabled = False
        ''Details4
        oGrid = oForm.Items.Item("35").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_18")
        'strqry = "Select * from [@SLSPUP]"
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Date] 'Date', T0.[U_Z_Branch] 'Branch', T0.[U_Z_ReportNo] 'Report Number', T0.[U_Z_CashNo] 'Cash Number',  T0.[U_Z_PayType] 'Payment Type', T0.[U_Z_Account] 'Account', T0.[U_Z_Currency] 'Currency', T0.[U_Z_PickAmt] 'Pickup Amount', T0.[U_Z_PickFrgCur] 'Pickup Foreign Currency Amt', T0.[U_Z_FrgCurrency] 'Foreign Currency' FROM [dbo].[@SLSPUP]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("35").Enabled = False
        oGrid = oForm.Items.Item("36").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_19")
        ' strqry = "Select * from [@SLSGFT]"
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Date] 'Date', T0.[U_Z_Branch] 'Branch', T0.[U_Z_ReportNo] 'Report Number', T0.[U_Z_CashNo] 'Cash Number',  T0.[U_Z_PayType] 'Payment Type', T0.[U_Z_Account] 'Account', T0.[U_Z_Currency] 'Currency', T0.[U_Z_GiftAmt] 'Gift Voucher Amount', T0.[U_Z_GiftFrgCur] 'Gift Voucher Foreign Amount', T0.[U_Z_FrgCurrency] 'Foreign Currency' FROM [dbo].[@SLSGFT]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("36").Enabled = False
        oGrid = oForm.Items.Item("37").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_20")
        ' strqry = "Select * from [@SLSPRE]"
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Date] 'Date', T0.[U_Z_Branch] 'Branch', T0.[U_Z_ReportNo] 'Report Number', T0.[U_Z_CashNo] 'Cash Number',  T0.[U_Z_PayType] 'Payment Type', T0.[U_Z_Account] 'Account', T0.[U_Z_Currency] 'Currency', T0.[U_Z_PrePayAmt] 'PrePayment Amount' FROM [dbo].[@SLSPRE]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("37").Enabled = False
        oGrid = oForm.Items.Item("38").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_21")
        '  strqry = "Select * from [@SLSEXP]"
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Date] 'Date', T0.[U_Z_Branch] 'Branch', T0.[U_Z_ReportNo] 'Report Number', T0.[U_Z_CashNo] 'Cash Number',   T0.[U_Z_Account] 'Account', T0.[U_Z_Currency] 'Currency', T0.[U_Z_ExpAmt] 'Expenses Amount'  FROM [dbo].[@SLSEXP]  T0"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("38").Enabled = False
        oGrid = oForm.Items.Item("39").Specific
        oGrid.DataTable = oForm.DataSources.DataTables.Item("DT_22")
        'strqry = "Select * from [@CREBAL]"
        strqry = "SELECT T0.[Code], T0.[Name], T0.[U_Z_Type] 'Type', T0.[U_Z_CompanyCode] 'Company Code', T0.[U_Z_Date] 'Date', T0.[U_Z_Branch] 'Branch', T0.[U_Z_ReportNo] 'Report Number', T0.[U_Z_CashNo] 'Cash Number',  T0.[U_Z_PayType] 'Payment Type', T0.[U_Z_Account] 'Account', T0.[U_Z_Customer] 'Customer', T0.[U_Z_CustomerType] 'Customer Type', T0.[U_Z_RefDoc] 'Refering Document ', T0.[U_Z_Currency] 'Currency', T0.[U_Z_BalAmt]  'Balance Amount' FROM [dbo].[@CREBAL]  T0 where isnull(U_Z_Imported,'N')='N'"
        oGrid.DataTable.ExecuteQuery(strqry)
        oForm.Items.Item("39").Enabled = False
    End Sub
#End Region
#Region "FormatGrid"
    Private Sub FormatGrid(ByVal aGrid As SAPbouiCOM.Grid, ByVal aChoice As String)
        If aChoice = "GIDN" Then
            aGrid.Columns.Item("Code").TitleObject.Caption = "Code"
            aGrid.Columns.Item("U_Z_CompanyCode").TitleObject.Caption = "Company Code"
            aGrid.Columns.Item("U_Z_Originator").TitleObject.Caption = "Originator"
            aGrid.Columns.Item("U_Z_GoodsNo").TitleObject.Caption = "Goods Number"
            aGrid.Columns.Item("U_Z_VatKey").TitleObject.Caption = "Vat Key"
            aGrid.Columns.Item("U_Z_Supplier").TitleObject.Caption = "Supplier"
            aGrid.Columns.Item("U_Z_Brand").TitleObject.Caption = "Brand"
            aGrid.Columns.Item("U_Z_GoodsDate").TitleObject.Caption = "Goods Date"
            aGrid.Columns.Item("U_Z_SupplierDocNo").TitleObject.Caption = "Supplier Doc.No"
            aGrid.Columns.Item("U_Z_Currency").TitleObject.Caption = "Currency"
            aGrid.Columns.Item("U_Z_GoodsBranch").TitleObject.Caption = "Goods Branch"
            aGrid.Columns.Item("U_Z_Value").TitleObject.Caption = "Value"
            aGrid.Columns.Item("U_Z_PayDate").TitleObject.Caption = "Payment Date"
            aGrid.Columns.Item("U_Z_PeriodPay").TitleObject.Caption = "Period Payment"
            aGrid.Columns.Item("U_Z_VatPercentage").TitleObject.Caption = "Vat Percentage"
            aGrid.Columns.Item("U_Z_SupplierType").TitleObject.Caption = "Supplier Type"
      
        End If
        aGrid.AutoResizeColumns()
        aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
#End Region

    Private Function DocumentCreation() As Boolean
        Dim strPath As String
        strFileName = Now.ToLongDateString
        strPath = System.Windows.Forms.Application.StartupPath
        strPath = strPath & "\ImportLog_" & strFileName & ".txt"
        strImportErrorLog = strPath

        ' strImportErrorLog = System.Windows.Forms.Application.StartupPath & "\ImportLog_Invoice.txt"

        If File.Exists(strImportErrorLog) Then
            '   File.Delete(strImportErrorLog)
        End If

        blnErrorflag = False

        WriteErrorlog("Document Creation started....", strImportErrorLog)
        WriteErrorlog("Creating GIDN Financial Entries", strImportErrorLog)

        If oApplication.Utilities.CreateGIDN() = False Then
            '  Return False
            blnErrorflag = False
        End If

        WriteErrorlog("GIDN Financial Entries Completed", strImportErrorLog)

        WriteErrorlog("Creating SLSTRN Documents", strImportErrorLog)
        If oApplication.Utilities.CreateSLSTRN() = False Then
            ' Return False
            blnErrorflag = False
        End If

        WriteErrorlog("SLSTRN Documents Process Completed", strImportErrorLog)

        WriteErrorlog("Creating SLSPAY Documents", strImportErrorLog)
        If oApplication.Utilities.CreateSLSPAY() = False Then
            'Return False
            blnErrorflag = False
        End If
        WriteErrorlog("SLSPAY Documents Process Completed", strImportErrorLog)

        WriteErrorlog("Creating SLSDIFF Documents", strImportErrorLog)
        If oApplication.Utilities.CreateSLSDIF() = False Then
            'Return False
            blnErrorflag = False
        End If
        WriteErrorlog("SLSDIFF Documents Process Completed", strImportErrorLog)
        'If blnErrorflag = False Then
        '    WriteErrorlog("Document Creation Completed with errors, please try again....", strImportErrorLog)
        '    Return False
        'End If
        WriteErrorlog("Document Creation Completed....", strImportErrorLog)

        Return True
    End Function

    Private Function CreateDocuments(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            oApplication.Company.StartTransaction()
            If DocumentCreation() = True Then
                If oApplication.Company.InTransaction() Then
                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                End If
            Else
                If oApplication.Company.InTransaction() Then
                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End If
            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False

        End Try
    End Function
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Import Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "15" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to import the details into SAP?", , "Yes", "No") = 2 Then
                                        Exit Sub
                                    End If
                                    If CreateDocuments(oForm) = True Then
                                        oApplication.Utilities.Message("Documents created successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        Databind()
                                    Else
                                        oApplication.Utilities.Message("Document creations failed ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If

                                    Dim x As System.Diagnostics.ProcessStartInfo
                                    x = New System.Diagnostics.ProcessStartInfo
                                    x.UseShellExecute = True
                                    sPath = strImportErrorLog 'System.Windows.Forms.Application.StartupPath & "\FImportLog_Invoice.txt"
                                    x.FileName = sPath
                                    System.Diagnostics.Process.Start(x)
                                    x = Nothing

                                End If

                                If pVal.ItemUID = "12" Then
                                    fillopen()
                                    oEditText = oForm.Items.Item("6").Specific
                                    oEditText.String = strSelectedFilepath
                                ElseIf pVal.ItemUID = "3" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to import the  into UDT?", , "Yes", "No") = 2 Then
                                        Exit Sub
                                    End If
                                    'ReadImportFiles(oForm)
                                    'Import(oForm)
                                    readFiles(oForm, strSelectedFilepath)
                                ElseIf pVal.ItemUID = "8" Then
                                    'If oApplication.SBO_Application.MessageBox("Do you want to import the documents?", , "Yes", "No") = 2 Then
                                    '    Exit Sub
                                    'End If
                                    'Import(oForm)
                                ElseIf pVal.ItemUID = "1000001" Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 2
                                    Databind()
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "10" Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 2
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "11" Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 8
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "1000002" Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 14
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "13" Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 20
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "14" Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 1
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "Folder1" Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 2
                                    oForm.Freeze(False)
                                    oForm.Update()
                                ElseIf pVal.ItemUID = "Folder2" Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 3
                                    oForm.Freeze(False)
                                    oForm.Update()
                                ElseIf pVal.ItemUID = "Folder3" Then
                                    oForm.PaneLevel = 4
                                ElseIf pVal.ItemUID = "Folder4" Then
                                    oForm.PaneLevel = 5
                                ElseIf pVal.ItemUID = "Folder5" Then
                                    oForm.PaneLevel = 6
                                ElseIf pVal.ItemUID = "Folder6" Then
                                    oForm.PaneLevel = 7
                                ElseIf pVal.ItemUID = "Finance1" Then
                                    oForm.PaneLevel = 8
                                ElseIf pVal.ItemUID = "Finance2" Then
                                    oForm.PaneLevel = 9
                                ElseIf pVal.ItemUID = "Finance3" Then
                                    oForm.PaneLevel = 10
                                ElseIf pVal.ItemUID = "Finance4" Then
                                    oForm.PaneLevel = 11
                                ElseIf pVal.ItemUID = "Finance5" Then
                                    oForm.PaneLevel = 12
                                ElseIf pVal.ItemUID = "Finance6" Then
                                    oForm.PaneLevel = 13
                                ElseIf pVal.ItemUID = "Sales1" Then
                                    oForm.PaneLevel = 14
                                ElseIf pVal.ItemUID = "Sales2" Then
                                    oForm.PaneLevel = 15
                                ElseIf pVal.ItemUID = "Sales3" Then
                                    oForm.PaneLevel = 16
                                ElseIf pVal.ItemUID = "Sales4" Then
                                    oForm.PaneLevel = 17
                                ElseIf pVal.ItemUID = "Sales5" Then
                                    oForm.PaneLevel = 18
                                ElseIf pVal.ItemUID = "Sales6" Then
                                    oForm.PaneLevel = 19
                                ElseIf pVal.ItemUID = "Payments1" Then
                                    oForm.PaneLevel = 20
                                ElseIf pVal.ItemUID = "Payments2" Then
                                    oForm.PaneLevel = 21
                                    oNewItem = oForm.Items.Item(pVal.ItemUID)
                                    oFolder = oNewItem.Specific
                                    oFolder.Select()
                                ElseIf pVal.ItemUID = "Payments3" Then
                                    oForm.PaneLevel = 22
                                    oNewItem = oForm.Items.Item(pVal.ItemUID)
                                    oFolder = oNewItem.Specific
                                    oFolder.Select()
                                ElseIf pVal.ItemUID = "Payments4" Then
                                    oForm.PaneLevel = 23
                                    oNewItem = oForm.Items.Item(pVal.ItemUID)
                                    oFolder = oNewItem.Specific
                                    oFolder.Select()
                                ElseIf pVal.ItemUID = "Payments5" Then
                                    oForm.PaneLevel = 24
                                    oNewItem = oForm.Items.Item(pVal.ItemUID)
                                    oFolder = oNewItem.Specific
                                    oFolder.Select()
                                End If

                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Import
                    If pVal.BeforeAction = False Then
                        'oApplication.Utilities.Message("Import process under development", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'Exit Sub
                        LoadForm()
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
