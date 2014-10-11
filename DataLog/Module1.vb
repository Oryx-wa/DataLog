Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase
Namespace DataLog
    Module Module1

        Private Addon As SboAddon
        <STAThread()>
        Public Sub Main()
            Dim dbo_RunApplication As Boolean

            Addon = New DataLogAddon(Windows.Forms.Application.StartupPath, "Oryx Addons", dbo_RunApplication)

            If dbo_RunApplication = True Then
                System.Windows.Forms.Application.Run()
            Else
                System.Windows.Forms.Application.Exit()
            End If
        End Sub



    End Module
End Namespace