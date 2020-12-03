Imports Autodesk.Revit.UI
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.Attributes
<Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)>
Public Class clsFormTest
    Implements IExternalCommand
    Public Function Execute(commandData As ExternalCommandData, ByRef message As String, elements As ElementSet) As Result Implements IExternalCommand.Execute 'Phương thức execute chưa rõ
        'TaskDialog.Show("RevitAPI", "HelloWorld")
        Dim Frm As New Valhalla
        Frm.ShowDialog()
        Return Result.Succeeded 'phần này chưa rõ, khi thay đổi các kiểu khác thì vẫn ra 1 kết quả như vậy.
    End Function
End Class
