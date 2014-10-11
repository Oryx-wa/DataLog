Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase

Namespace DataLog
    <FormAttribute("DataLog.ACQ", "SBOForms/ACQ.b1f")>
    Friend Class ACQ
        Inherits UserFormBaseClass

        Public Sub New()
        End Sub


        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent

            oForm = SBO_Application.Forms.ActiveForm

            If oForm.TypeEx = "DataLog.ACQ" Then
                If (pVal.BeforeAction) Then
                    oDataSource = oForm.DataSources.DBDataSources.Item("OCRD")
                    txtCust = oForm.Items.Item("txtBPCode").Specific
                    If getOffset(txtCust.Value, "CardCode", oDataSource) Then
                        oDataSource.Offset = 0                      
                    End If
                End If
            End If
        End Sub
        Private Sub txtCust_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtCust.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_INVACQHDR")
            oDataSource.SetValue("Code", 0, val)
            oDataSource = oForm.DataSources.DBDataSources.Item("OCRD")
            If getOffset(val, "CardCode", oDataSource) Then
                oDataSource.Offset = 0
                getOffset(val, "CardCode", oDataSource)
                oDataSource2 = oForm.DataSources.DBDataSources.Item("@OWA_INVACQHDR")
                oDataSource2.SetValue("Name", 0, oDataSource.GetValue("CardName", 0))
            End If

        End Sub

        Private Sub matACQ_PressedBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles matACQ.PressedBefore
            Try
                oForm.DataSources.DBDataSources.Item("@OWA_INVACQLINES").Clear()
                matACQ = oForm.Items.Item("matACQ").Specific

                If pVal.Row = matACQ.RowCount + 1 Then
                    If pVal.Row = 1 Then
                        matACQ.AddRow(1)
                    Else
                        matACQ.AddRow(1, matACQ.RowCount)
                    End If
                    matACQ.Columns.Item(1).Cells.Item(pVal.Row).Click()
                End If
            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub colACQ_ValidateBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles colACQ.ValidateBefore
            Try
                Dim oEditACQ As SAPbouiCOM.EditText    ' Item Quantity
                Dim oEditTOP As SAPbouiCOM.EditText   ' Total = Item Price * Item Quantity

                ' Get the items from the matrix
                oEditACQ = colACQ.Cells.Item(pVal.Row).Specific
                oEditTOP = colTOP.Cells.Item(pVal.Row).Specific

                colTOP.Editable = False
                oEditTOP.Value = If(oEditACQ.Value.Length = 0, 0, oEditACQ.Value) * 0.9
            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Public Overrides Sub OnInitializeComponent()


            Me.txtCust = CType(Me.GetItem("txtBPCode").Specific, SAPbouiCOM.EditText)
            Me.matACQ = CType(Me.GetItem("matACQ").Specific, SAPbouiCOM.Matrix)
            Me.colACQ = CType(Me.GetItem("matACQ").Specific, SAPbouiCOM.Matrix).Columns.Item("colacq")
            Me.colTOP = CType(Me.GetItem("matACQ").Specific, SAPbouiCOM.Matrix).Columns.Item("coltop")

            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

        Private Sub OnCustomInitialize()

        End Sub

        Private WithEvents txtCust As SAPbouiCOM.EditText
        Private WithEvents matACQ As SAPbouiCOM.Matrix
        Private WithEvents colACQ As SAPbouiCOM.Column
        Private WithEvents colTOP As SAPbouiCOM.Column
       
       
    End Class
End Namespace
