Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase

Namespace DataLog
    <FormAttribute("DataLog.MonthlyDataLog", "SBOForms/MonthlyDataLog.b1f")>
    Friend Class MonthlyDataLog
        Inherits UserFormBaseClass

        Public Sub New()


        End Sub

      
        Private Sub txtInv_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtInv.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_INVMDLHDR")
            oDataSource.SetValue("U_DocNum", 0, val)
            oDataSource = oForm.DataSources.DBDataSources.Item("OINV")
            If getOffset(CInt(val), "DocEntry", oDataSource) Then
                oDataSource.Offset = 0
            End If

            ' Put form in UPDATE Mode when in OK Mode
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

        End Sub

        Private Sub txtGLAcct_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtGLAcct.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_INVMDLHDR")
            oDataSource.SetValue("U_GlAcct", 0, val)
            oDataSource = oForm.DataSources.DBDataSources.Item("OACT")
            If getOffset(val, "AcctCode", oDataSource) Then
                oDataSource.Offset = 0
            End If

            ' Put form in UPDATE Mode when in OK Mode
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.ComboBox0 = CType(Me.GetItem("txtMonth").Specific, SAPbouiCOM.ComboBox)
            Me.txtInv = CType(Me.GetItem("txtInv").Specific, SAPbouiCOM.EditText)
            Me.txtMthName = CType(Me.GetItem("txtMthName").Specific, SAPbouiCOM.EditText)
            Me.txtACQ = CType(Me.GetItem("txtACQ").Specific, SAPbouiCOM.EditText)
            Me.txtYear = CType(Me.GetItem("txtYear").Specific, SAPbouiCOM.EditText)
            Me.txtDocNo = CType(Me.GetItem("txtDocEnt").Specific, SAPbouiCOM.EditText)
            Me.txtQtyCons = CType(Me.GetItem("txtQtyCons").Specific, SAPbouiCOM.EditText)
            Me.txtMUG = CType(Me.GetItem("txtMUG").Specific, SAPbouiCOM.EditText)
            Me.txtTOPQty = CType(Me.GetItem("txtTOPQty").Specific, SAPbouiCOM.EditText)
            Me.btnCalcQty = CType(Me.GetItem("btnCalcQty").Specific, SAPbouiCOM.Button)
            Me.btnGenInvoice = CType(Me.GetItem("btnGenInv").Specific, SAPbouiCOM.Button)
            Me.txtQtyBilled = CType(Me.GetItem("txtQtyBill").Specific, SAPbouiCOM.EditText)
            Me.txtUPrice = CType(Me.GetItem("txtUPrice").Specific, SAPbouiCOM.EditText)
            Me.txtAmt = CType(Me.GetItem("txtAmt").Specific, SAPbouiCOM.EditText)
            Me.StaticText0 = CType(Me.GetItem("Item_8").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("Item_36").Specific, SAPbouiCOM.EditText)
            Me.txtCust = CType(Me.GetItem("txtCust").Specific, SAPbouiCOM.EditText)
            Me.Folder0 = CType(Me.GetItem("Item_43").Specific, SAPbouiCOM.Folder)
            Me.Matrix0 = CType(Me.GetItem("Item_44").Specific, SAPbouiCOM.Matrix)
            Me.txtGLAcct = CType(Me.GetItem("glacct").Specific, SAPbouiCOM.EditText)
            Me.txtDesc = CType(Me.GetItem("Desc").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub


        Private Sub OnCustomInitialize()

        End Sub

        Private WithEvents txtCust As SAPbouiCOM.EditText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents txtInv As SAPbouiCOM.EditText
        Private WithEvents txtMthName As SAPbouiCOM.EditText
        Private WithEvents txtACQ As SAPbouiCOM.EditText
        Private WithEvents txtYear As SAPbouiCOM.EditText
        Private WithEvents txtDocNo As SAPbouiCOM.EditText
        Private WithEvents txtQtyCons As SAPbouiCOM.EditText
        Private WithEvents txtMUG As SAPbouiCOM.EditText
        Private WithEvents txtTOPQty As SAPbouiCOM.EditText
        Private WithEvents btnCalcQty As SAPbouiCOM.Button
        Private WithEvents btnGenInvoice As SAPbouiCOM.Button
        Private WithEvents txtAmt As SAPbouiCOM.EditText
        Private WithEvents txtUPrice As SAPbouiCOM.EditText
        Private WithEvents txtQtyBilled As SAPbouiCOM.EditText
        Private WithEvents txtGLAcct As SAPbouiCOM.EditText
        Private WithEvents txtDesc As SAPbouiCOM.EditText

        Private Sub ComboBox0_ComboSelectAfter(ByVal sboObject As System.Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles ComboBox0.ComboSelectAfter

            txtMthName = oForm.Items.Item("txtMthName").Specific

            Select Case pVal.PopUpIndicator
                Case 0 : txtMthName.Value = "January"
                Case 1 : txtMthName.Value = "February"
                Case 2 : txtMthName.Value = "March"
                Case 3 : txtMthName.Value = "April"
                Case 4 : txtMthName.Value = "May"
                Case 5 : txtMthName.Value = "June"
                Case 6 : txtMthName.Value = "July"
                Case 7 : txtMthName.Value = "August"
                Case 8 : txtMthName.Value = "September"
                Case 9 : txtMthName.Value = "October"
                Case 10 : txtMthName.Value = "November"
                Case 11 : txtMthName.Value = "December"

            End Select

        End Sub


        Private Function GetSCFValue() As Decimal

            Dim oRec As SAPbobsCOM.Recordset, sQuery As String, dSCF As Decimal

            'get the following values
            txtCust = oForm.Items.Item("txtCust").Specific

            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_INVMDLHDR")

            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'use the value of customer to get the value of SCF
            sQuery = "Select * From [OCRD] Where CardCode='" & txtCust.Value & "'"
            oRec.DoQuery(sQuery)

            If oRec.RecordCount > 0 Then
                dSCF = oRec.Fields.Item("U_SCFValue").Value
            Else
                dSCF = 35.29
            End If

            Return dSCF

        End Function
        Private Sub txtYear_ValidateBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles txtYear.ValidateBefore
            Try

                Dim oRec As SAPbobsCOM.Recordset, sQuery As String, dSCF As Decimal

                'get the following values
                txtCust = oForm.Items.Item("txtCust").Specific
                txtACQ = oForm.Items.Item("txtACQ").Specific
                txtYear = oForm.Items.Item("txtYear").Specific

                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_INVMDLHDR")

                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
 
                dSCF = GetSCFValue()

                'use the value of year and customer to get the ACQ for that year
                sQuery = "Select * From [@OWA_INVACQLINES] Where Code='" & txtCust.Value & "' "
                sQuery &= "And U_Year='" & txtYear.Value & "'"

                oRec.DoQuery(sQuery)

                If oRec.RecordCount > 0 Then
                    oDataSource.SetValue("U_ACQ", 0, oRec.Fields.Item("U_ACQ").Value)
                    oDataSource.SetValue("U_TOPQty", 0, oRec.Fields.Item("U_TOP").Value)
                    oDataSource.SetValue("U_TOPQtySCF", 0, oRec.Fields.Item("U_TOP").Value * dSCF)
                    oDataSource.SetValue("U_MthTOPQty", 0, oRec.Fields.Item("U_TOP").Value / 12)
                    oDataSource.SetValue("U_MthTOPQtySCF", 0, oRec.Fields.Item("U_TOP").Value * dSCF / 12)
                    oDataSource.SetValue("U_UnitPrice", 0, oRec.Fields.Item("U_UnitPrice").Value)
                Else
                    oDataSource.SetValue("U_ACQ", 0, 0)
                    oDataSource.SetValue("U_TOPQty", 0, 0)
                    oDataSource.SetValue("U_TOPQtySCF", 0, 0)
                    oDataSource.SetValue("U_MthTOPQty", 0, 0)
                    oDataSource.SetValue("U_MthTOPQtySCF", 0, 0)
                    oDataSource.SetValue("U_UnitPrice", 0, 0)
                End If



            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub


        Private Sub btnCalcQty_ClickBefore(ByVal sboObject As System.Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles btnCalcQty.ClickBefore

            Try
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_INVMDLHDR")

                If SBO_Application.MessageBox("Do you want to calculate off-take?", 2, "Yes", "No", "") = 2 Then
                    Exit Sub
                Else
                    If Trim(oDataSource.GetValue("U_OffTakeCalc", 0)) = "Y" Then
                        SBO_Application.MessageBox("Off Take already calculated")
                        Exit Sub
                    End If

                    Dim oRec As SAPbobsCOM.Recordset, oRec2 As SAPbobsCOM.Recordset, sQuery As String, QtyCons As Decimal, BillQty As Decimal, dSCF As Decimal

                    'get the following values
                    txtDocNo = oForm.Items.Item("txtDocEnt").Specific
                    txtCust = oForm.Items.Item("txtCust").Specific

                    If String.IsNullOrEmpty(Trim(txtCust.Value)) Then
                        SBO_Application.MessageBox("Customer Code must be supplied")
                        Exit Sub
                    End If

                    If String.IsNullOrEmpty(Trim(oDataSource.GetValue("U_Year", 0))) Then
                        SBO_Application.MessageBox("Year Log must be supplied")
                        Exit Sub
                    End If

                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    dSCF = GetSCFValue()

                    'get the following values
                    txtDocNo = oForm.Items.Item("txtDocEnt").Specific
                    txtCust = oForm.Items.Item("txtCust").Specific
                    txtYear = oForm.Items.Item("txtYear").Specific

                    Dim strSQL As String

                    strSQL = Me.GetSQLString("owa_getConsQty", txtDocNo.Value, txtCust.Value, txtYear.Value)
                    oDataTable = oForm.DataSources.DataTables.Item("DT_Base")
                    oDataTable.ExecuteQuery(strSQL)
                    If oDataTable.Rows.Count = 1 Then
                        QtyCons = oDataTable.GetValue("ConsQty", 0)
                        BillQty = oDataTable.GetValue("BilledQty", 0)
                    Else
                        QtyCons = 0
                        BillQty = 0
                    End If

                    oDataSource.SetValue("U_QtyCons", 0, QtyCons)
                    oDataSource.SetValue("U_QtyBilled", 0, BillQty)


                    txtTOPQty = oForm.Items.Item("txtTOPQty").Specific
                    txtQtyCons = oForm.Items.Item("txtQtyCons").Specific
                    txtQtyBilled = oForm.Items.Item("txtQtyBill").Specific
                    txtUPrice = oForm.Items.Item("txtUPrice").Specific

                    Dim mtrRental As String, mtrRentalQty As Decimal
                    'use the customer to get the meter rental for the customer
                    sQuery = "Select * From [OCRD] Where CardCode='" & Trim(txtCust.Value) & "'"
                    oRec2.DoQuery(sQuery)

                    mtrRental = oRec2.Fields.Item("U_MtrRental").Value

                    'if meter rental = yes, then calculate meter rental
                    If mtrRental = "Y" Then
                        mtrRentalQty = GetMtrRentalQty(CDbl(Trim(oDataSource.GetValue("U_QtyConsSCF", 0))) / 12)
                    Else
                        mtrRentalQty = 0
                    End If

                    'get the MUG Brought Forward.
                    'subtract 1 from the current month to get the previous month

                    Dim mugBF As Decimal = 0
                    sQuery = "Select [U_MUGbf] From [@OWA_INVMDLHDR] Where U_BPCode='" & Trim(txtCust.Value) & _
                             "' and U_Month=" & CInt(Trim(oDataSource.GetValue("U_Month", 0))) - 1
                    oRec2.DoQuery(sQuery)

                    If oRec2.RecordCount = 1 Then
                        mugBF = CDec(Trim(oRec2.Fields.Item("U_MUGbf").Value))
                    Else
                        mugBF = 0
                    End If

                    If txtTOPQty.Value > txtQtyCons.Value Then
                        oDataSource.SetValue("U_QtyBilled", 0, txtTOPQty.Value)
                        oDataSource.SetValue("U_UnitPrice", 0, txtUPrice.Value)
                        oDataSource.SetValue("U_Amount", 0, txtQtyBilled.Value * txtUPrice.Value)
                        oDataSource.SetValue("U_MUG", 0, txtTOPQty.Value - txtQtyCons.Value)
                        oDataSource.SetValue("U_MUGbf", 0, mugBF)
                        oDataSource.SetValue("U_MUGcf", 0, mugBF + txtTOPQty.Value - txtQtyCons.Value)
                    Else
                        oDataSource.SetValue("U_QtyBilled", 0, txtQtyCons.Value)
                        oDataSource.SetValue("U_UnitPrice", 0, txtUPrice.Value)
                        oDataSource.SetValue("U_Amount", 0, txtQtyBilled.Value * txtUPrice.Value)
                        oDataSource.SetValue("U_MUG", 0, 0)
                        oDataSource.SetValue("U_MUGbf", 0, mugBF)
                        oDataSource.SetValue("U_MUGcf", 0, mugBF)
                    End If

                    oDataSource.SetValue("U_MtrRentalQty", 0, mtrRentalQty)
                    oDataSource.SetValue("U_OffTakeCalc", 0, "Y")

                    ' Put form in UPDATE Mode when in OK Mode
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                End If

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Function GetMtrRentalQty(ByVal offTakeValue As Decimal) As Decimal
            Dim rentalMtrQty As Decimal

            If offTakeValue > 2000000 Then
                rentalMtrQty = offTakeValue * 0.15 / 100
            ElseIf offTakeValue < 1990000 And offTakeValue >= 1500000 Then
                rentalMtrQty = offTakeValue * 0.25 / 100
            ElseIf offTakeValue < 1490000 And offTakeValue >= 500000 Then
                rentalMtrQty = offTakeValue * 0.35 / 100
            ElseIf offTakeValue < 490000 And offTakeValue >= 200000 Then
                rentalMtrQty = offTakeValue * 0.35 / 100
            End If
                

            Return rentalMtrQty
        End Function

       

        Private Sub btnGenInvoice_ClickBefore(ByVal sboObject As System.Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles btnGenInvoice.ClickBefore

            Try
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_INVMDLHDR")

                If SBO_Application.MessageBox("Do you want to post Invoice?", 2, "Yes", "No", "") = 2 Then
                    Exit Sub
                Else
                    'tests if invoice already posted
                    If Trim(oDataSource.GetValue("U_InvStatus", 0)) = "Posted" Then
                        SBO_Application.MessageBox("Invoice already posted")
                        Exit Sub
                    End If

                    'tests if offtake already calculated
                    If Trim(oDataSource.GetValue("U_OffTakeCalc", 0)) = "N" Then
                        SBO_Application.MessageBox("Off-Take not yet calculated")
                        Exit Sub
                    End If

                    'generate invoice automatically
                    Dim invHeader As SAPbobsCOM.Documents = Nothing, invLines As SAPbobsCOM.Document_Lines = Nothing, sDate As String

                    invHeader = DirectCast(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts), SAPbobsCOM.Documents)
                    invHeader.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices

                    invLines = invHeader.Lines

                    oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_INVMDLHDR")
                    oDataSource.Offset = 0

                    sDate = oDataSource.GetValue("U_DocDate", 0)

                    With invHeader
                        .CardCode = oDataSource.GetValue("U_BPCode", 0)
                        .CardName = oDataSource.GetValue("U_BPName", 0)
                        .DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                        .DocDate = sDate.Substring(6, 2) & "." & sDate.Substring(4, 2) & "." & sDate.Substring(0, 4)
                        .NumAtCard = oDataSource.GetValue("U_SerNum", 0)
                        .JournalMemo = "A/R Invoices - " & oDataSource.GetValue("U_BPCode", 0)
                        .Comments = oDataSource.GetValue("U_SerNum", 0)

                        'user fields
                        .UserFields.Fields.Item("U_MUGMTH").Value = oDataSource.GetValue("U_MUG", 0)
                        .UserFields.Fields.Item("U_MUGBF").Value = oDataSource.GetValue("U_MUGbf", 0)
                        .UserFields.Fields.Item("U_MUGCF").Value = oDataSource.GetValue("U_MUGcf", 0)
                        .UserFields.Fields.Item("U_MUGUtil").Value = oDataSource.GetValue("U_MUGRecouped", 0)

                    End With

                    With invLines
                        .SetCurrentLine(0)
                        .AccountCode = txtGLAcct.Value         ' Account Code
                        .ItemDescription = txtDesc.Value    '  Description
                        .Quantity = oDataSource.GetValue("U_QtyBilled", 0)              ' Quantity Sold
                        .UnitPrice = oDataSource.GetValue("U_UnitPrice", 0)     ' Price
                        .LineTotal = oDataSource.GetValue("U_Amount", 0)

                        'user fields
                        .UserFields.Fields.Item("U_TopQty").Value = oDataSource.GetValue("U_TOPQty", 0)
                        .UserFields.Fields.Item("U_QtyCnsmd").Value = oDataSource.GetValue("U_QtyCons", 0)
                        .UserFields.Fields.Item("U_QtyBilled").Value = oDataSource.GetValue("U_QtyBilled", 0)
                        .UserFields.Fields.Item("U_QtyCnsmd").Value = oDataSource.GetValue("U_QtyCons", 0)
                        .UserFields.Fields.Item("U_MtrRntl").Value = oDataSource.GetValue("U_MtrRentalQty", 0)
                        .UserFields.Fields.Item("U_UnitPr").Value = oDataSource.GetValue("U_UnitPrice", 0)
                        .UserFields.Fields.Item("U_Amt").Value = oDataSource.GetValue("U_Amount", 0)
                    End With

                    invHeader.Add()

                    'set invoice status to posted
                    oDataSource.SetValue("U_InvStatus", 0, "Posted")

                    'commit status value
                    Dim oRec As SAPbobsCOM.Recordset, sQuery As String
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    txtDocNo = oForm.Items.Item("txtDocEnt").Specific
                    sQuery = "update [@OWA_INVMDLHDR] set U_InvStatus='Posted' Where DocEntry=" & txtDocNo.Value
                    oRec.DoQuery(sQuery)

                    SBO_Application.MessageBox("Invoice successfully posted")

                End If

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try



        End Sub
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText


       
        Private Sub txtCust_ChooseFromListAfter(ByVal sboObject As System.Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtCust.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_INVMDLHDR")
            oDataSource.SetValue("U_BPCode", 0, val)
            oDataSource = oForm.DataSources.DBDataSources.Item("OCRD")
            If getOffset(val, "CardCode", oDataSource) Then
                oDataSource.Offset = 0
                getOffset(val, "CardCode", oDataSource)
                oDataSource2 = oForm.DataSources.DBDataSources.Item("@OWA_INVMDLHDR")
                oDataSource2.SetValue("U_BPName", 0, oDataSource.GetValue("CardName", 0))
            End If

        End Sub
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
    End Class
End Namespace
