Public Class clsBrandMapping
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()
        oApplication.Utilities.LoadForm(xml_BrandMapping, frm_BrandMapping)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oGrid = oForm.Items.Item("1").Specific
        FormatGrid(oGrid)
    End Sub
    Private Sub FormatGrid(ByVal aGrid As SAPbouiCOM.Grid)
        oGrid.DataTable.ExecuteQuery("select PrcCode,PrcName,U_Z_Brand,U_Z_Branch,U_Z_CardCode,U_Z_warehouse  from OPRC where DimCode=2 and locked='N' order by Prccode")
        oGrid.Columns.Item("PrcCode").TitleObject.Caption = "SAP Branch Code"
        oGrid.Columns.Item("PrcCode").Editable = False
        oGrid.Columns.Item("PrcName").TitleObject.Caption = "SAP Branch Name"
        oGrid.Columns.Item("PrcName").Editable = False
        oGrid.Columns.Item("U_Z_Brand").TitleObject.Caption = "Futura Branch Code"
        oGrid.Columns.Item("U_Z_Branch").TitleObject.Caption = "Futura Brand Code"
        oGrid.Columns.Item("U_Z_Branch").Visible = False
        oGrid.Columns.Item("U_Z_CardCode").TitleObject.Caption = "SAP Customer Code"
        oGrid.Columns.Item("U_Z_CardCode").Visible = False
        oGrid.Columns.Item("U_Z_warehouse").TitleObject.Caption = "SAP Warehosue Code"
        oGrid.Columns.Item("U_Z_warehouse").Visible = False
        oEditTextColumn = oGrid.Columns.Item("U_Z_CardCode")
        oEditTextColumn.ChooseFromListUID = "CFL_2"
        oEditTextColumn.ChooseFromListAlias = "CardCode"
        oEditTextColumn.LinkedObjectType = "2"
        oEditTextColumn = oGrid.Columns.Item("U_Z_warehouse")
        oEditTextColumn.ChooseFromListUID = "CFL_3"
        oEditTextColumn.ChooseFromListAlias = "WhsCode"
        oEditTextColumn.LinkedObjectType = "64"
        oGrid.AutoResizeColumns()
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None

    End Sub

    Private Sub addtoUDT(ByVal aform As SAPbouiCOM.Form)
        Dim oRec As SAPbobsCOM.Recordset
        oGrid = aform.Items.Item("1").Specific
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oRec.DoQuery("Update OPRC set U_Z_Brand='" & oGrid.DataTable.GetValue("U_Z_Brand", intRow) & "',U_Z_Branch='" & oGrid.DataTable.GetValue("U_Z_Branch", intRow) & "',U_Z_CardCode='" & oGrid.DataTable.GetValue("U_Z_CardCode", intRow) & "',U_Z_warehouse='" & oGrid.DataTable.GetValue("U_Z_warehouse", intRow) & "' where PrcCode='" & oGrid.DataTable.GetValue("PrcCode", intRow) & "'")
        Next
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BrandMapping Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" Then
                                    oGrid = oForm.Items.Item("1").Specific
                                    addtoUDT(oForm)
                                    FormatGrid(oGrid)
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        oGrid = oForm.Items.Item("1").Specific
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If ((pVal.ItemUID = "11" And (pVal.ColUID = "ItemCode" Or pVal.ColUID = "Dscription" Or pVal.ColUID = "CodeBars"))) Then
                                        ElseIf pVal.ItemUID = "1" And (pVal.ColUID = "U_Z_warehouse" Or pVal.ColUID = "U_Z_CardCode") Then
                                            If pVal.ColUID = "U_Z_CardCode" Then
                                                val = oDataTable.GetValue("CardCode", 0)
                                            Else
                                                val = oDataTable.GetValue("WhsCode", 0)
                                            End If

                                            oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)

                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                End Try

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
                Case mnu_Brand
                    If pVal.BeforeAction = False Then
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
