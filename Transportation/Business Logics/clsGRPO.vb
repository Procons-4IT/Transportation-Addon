Public Class clsGRPO
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
    Private oItems As SAPbouiCOM.Item
    Private oBP As SAPbobsCOM.BusinessPartners
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub AddControl(ByVal aForm As SAPbouiCOM.Form)
        aForm.Freeze(True)
        oApplication.Utilities.AddControls(aForm, "stIns", "86", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , , , "Include Transportation", 140)
        oApplication.Utilities.AddControls(aForm, "cmbIns", "46", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, "DOWN", , )

        oCombobox = aForm.Items.Item("cmbIns").Specific
        oCombobox.ValidValues.Add("Y", "Yes")
        oCombobox.ValidValues.Add("N", "No")
        oCombobox.DataBind.SetBound(True, "OPDN", "U_Z_Trans")

        oItems = aForm.Items.Item("stIns")
        oItems.LinkTo = "cmbIns"

        aForm.Items.Item("cmbIns").DisplayDesc = True
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

        oApplication.Utilities.AddControls(aForm, "stType", "stIns", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , , , "Transportation Type", 140)
        oApplication.Utilities.AddControls(aForm, "cmbType", "cmbIns", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, "DOWN", , )
        oItems = aForm.Items.Item("stType")
        oItems.LinkTo = "cmbType"
        oCombobox = aForm.Items.Item("cmbType").Specific
        oCombobox.DataBind.SetBound(True, "OPDN", "U_Z_TrnsCode")
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("Select U_Z_Code,U_Z_Name from [@Z_OTRANS] order by Convert(Numeric,Code)")
        For intRow As Integer = 0 To oTest.RecordCount - 1
            oCombobox.ValidValues.Add(oTest.Fields.Item(0).Value, oTest.Fields.Item(1).Value)
            oTest.MoveNext()
        Next
        aForm.Items.Item("cmbType").DisplayDesc = True
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

        oApplication.Utilities.AddControls(aForm, "stAmt", "stType", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , , , "Transportation Amount", )
        oApplication.Utilities.AddControls(aForm, "edAmt", "cmbType", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", , )
        oItems = aForm.Items.Item("stAmt")
        oItems.LinkTo = "edAmt"
        oEditText = aForm.Items.Item("edAmt").Specific
        oEditText.DataBind.SetBound(True, "OPDN", "U_Z_Amount")
        AddChooseFromList(aForm)

        oApplication.Utilities.AddControls(aForm, "stCardCode", "70", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , , , "Transporter  Code", )
        oApplication.Utilities.AddControls(aForm, "edCardCode", "stCardCode", SAPbouiCOM.BoFormItemTypes.it_EDIT, "RIGHT", , , , , 120)
        oItems = aForm.Items.Item("stCardCode")
        oItems.LinkTo = "edCardCode"
        oEditText = aForm.Items.Item("edCardCode").Specific
        oEditText.DataBind.SetBound(True, "OPDN", "U_Z_CardCode")
        oEditText.ChooseFromListUID = "CFL_1"
        oEditText.ChooseFromListAlias = "CardCode"
        aForm.Freeze(False)
    End Sub

    Private Function createDocument(ByVal aDocEntry As Integer) As Boolean
        Dim oDoc As SAPbobsCOM.Documents
        Dim oStockTransfer As SAPbobsCOM.Documents
        Dim otest, otest1 As SAPbobsCOM.Recordset
        oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        oStockTransfer = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
        If oStockTransfer.GetByKey(aDocEntry) Then
            If oStockTransfer.UserFields.Fields.Item("U_Z_Trans").Value = "Y" And oStockTransfer.UserFields.Fields.Item("U_Z_TrnsCode").Value <> "" And oStockTransfer.UserFields.Fields.Item("U_Z_Amount").Value > 0 Then
                If oStockTransfer.UserFields.Fields.Item("U_Z_Invoicde").Value = "Y" Then
                    Return True
                End If
                oDoc.DocDate = oStockTransfer.DocDate
                oDoc.DocDueDate = oStockTransfer.DocDueDate
                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                oDoc.CardCode = oStockTransfer.UserFields.Fields.Item("U_Z_CardCode").Value ' oStockTransfer.CardCode
                otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otest.DoQuery("Select * from [@Z_OTRANS] where U_Z_Code='" & oStockTransfer.UserFields.Fields.Item("U_Z_TrnsCode").Value & "'")
                otest1.DoQuery("Select * from OACT where FormatCode='" & otest.Fields.Item("U_Z_GLACC").Value & "'")
                oDoc.UserFields.Fields.Item("U_Z_StockNo").Value = oStockTransfer.DocNum.ToString
                oDoc.Comments = "Created based on Goods Receipt PO Number : " & oStockTransfer.DocNum & " and transportation Type : " & otest.Fields.Item("U_Z_Name").Value
                oDoc.Lines.ItemDescription = otest.Fields.Item("U_Z_Name").Value
                oDoc.Lines.AccountCode = otest1.Fields.Item("AcctCode").Value
                oDoc.Lines.LineTotal = oStockTransfer.UserFields.Fields.Item("U_Z_Amount").Value
                If oDoc.Add <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    Dim strdoc As String
                    oApplication.Company.GetNewObjectCode(strdoc)
                    oStockTransfer.UserFields.Fields.Item("U_Z_Invoicde").Value = "Y"
                    oStockTransfer.UserFields.Fields.Item("U_Z_Ref").Value = strdoc
                    oStockTransfer.Update()
                End If
            End If
        End If
        Return True
    End Function
#Region "Add Choose From List"
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition


            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False




            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL_1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            ' Adding Conditions to CFL2

            '' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()



        Catch
            MsgBox(Err.Description)
        End Try
    End Sub

#End Region


    Private Function Validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        oCombobox = aform.Items.Item("cmbIns").Specific
        If oCombobox.Selected.Value = "Y" Then
            oCombobox = aform.Items.Item("cmbType").Specific
            Try
                If oCombobox.Selected.Value = "" Then
                    oApplication.Utilities.Message("Transportation type is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Catch ex As Exception
                oApplication.Utilities.Message("Transportation type is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try
            If oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aform, "edAmt")) = 0 Then
                ' oApplication.Utilities.Message("Transportation Amount is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                ' Return False
            End If
            If oApplication.Utilities.getEdittextvalue(aform, "edCardCode") = "" Then
                oApplication.Utilities.Message("Transporter  details is missing..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                Dim otest As SAPbobsCOM.Recordset
                otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otest.DoQuery("Select * from OCRD where CardType='C' and CardCode='" & oApplication.Utilities.getEdittextvalue(aform, "edCardCode") & "'")
                If otest.RecordCount > 0 Then
                    oApplication.Utilities.Message("Only Vendors should be selected if the transportation is incldued ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
        End If

        Return True
    End Function
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_GRPO Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "cmbType" Then
                                    oCombobox = oForm.Items.Item("cmbIns").Specific
                                    Try
                                        If oCombobox.Selected.Value = "N" Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Catch ex As Exception
                                        BubbleEvent = False
                                        Exit Sub
                                    End Try

                                End If

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "edAmt" And pVal.CharPressed <> 9 Then
                                    oCombobox = oForm.Items.Item("cmbIns").Specific
                                    Try
                                        If oCombobox.Selected.Value = "N" Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Catch ex As Exception
                                        BubbleEvent = False
                                        Exit Sub
                                    End Try

                                End If

                            Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "edAmt" And pVal.CharPressed <> 9 Then
                                    oCombobox = oForm.Items.Item("cmbIns").Specific
                                    Try
                                        If oCombobox.Selected.Value = "N" Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Catch ex As Exception
                                        BubbleEvent = False
                                        Exit Sub
                                    End Try

                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "edAmt" And pVal.ItemUID = "cmbType" And pVal.CharPressed <> 9 Then
                                    oCombobox = oForm.Items.Item("cmbIns").Specific
                                    Try
                                        If oCombobox.Selected.Value = "N" Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Catch ex As Exception
                                        BubbleEvent = False
                                        Exit Sub
                                    End Try

                                End If

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "cmbType" Then
                                    oCombobox = oForm.Items.Item("cmbIns").Specific
                                    Try
                                        If oCombobox.Selected.Value = "N" Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Catch ex As Exception
                                        BubbleEvent = False
                                        Exit Sub
                                    End Try

                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                AddControl(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "cmbType" Then
                                    oForm.Freeze(True)
                                    Dim otest As SAPbobsCOM.Recordset
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oCombobox = oForm.Items.Item("cmbType").Specific
                                    otest.DoQuery("Select * from [@Z_OTRANS] where U_Z_Code='" & oCombobox.Selected.Value & "'")
                                    oApplication.Utilities.setEdittextvalue(oForm, "edAmt", otest.Fields.Item("U_Z_Amount").Value)
                                    oForm.Freeze(False)
                                End If

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
                                        ' oGrid = oForm.Items.Item("11").Specific
                                        intChoice = 0
                                        oForm.Freeze(True)

                                        If pVal.ItemUID = "edCardCode" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            Try

                                                oApplication.Utilities.setEdittextvalue(oForm, "edCardCode", val)

                                            Catch ex As Exception

                                            End Try
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
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_InvSO
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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If oForm.TypeEx = frm_GRPO Then
                    oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    Dim oSt As SAPbobsCOM.Documents
                    oSt = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                    If oSt.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        If createDocument(oSt.DocEntry) = True Then
                            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    End If
                End If
            End If
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If oForm.TypeEx = frm_GRPO Then
                    oForm.Items.Item("cmbType").Enabled = False
                    oForm.Items.Item("edAmt").Enabled = False
                    oForm.Items.Item("cmbIns").Enabled = False
                    oForm.Items.Item("edCardCode").Enabled = False
                End If
            End If
            'If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
            '    oForm = oApplication.SBO_Application.Forms.ActiveForm()
            '    If oForm.TypeEx = frm_InventoryTransfer Then
            '        oForm.Items.Item("cmbType").Enabled = True
            '        oForm.Items.Item("edAmt").Enabled = True
            '        oForm.Items.Item("cmbIns").Enabled = True
            '    End If
            'End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

End Class
