Imports System.Collections.Generic
Imports System.Text
Imports System.Linq
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.UI.Selection
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.DB.Structure
<Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)>
Public Class clsDemDam
    Implements IExternalCommand
    Dim doc As Document
    Dim app As Application
    Public Function Execute(commandData As ExternalCommandData, ByRef message As String, elements As ElementSet) As Result Implements IExternalCommand.Execute 'Phương thức execute chưa rõ
        'Dim Frm As New Valhalla
        'Frm.ShowDialog()
        Dim UiApp As UIApplication = commandData.Application
        Dim Uidoc As UIDocument = UiApp.ActiveUIDocument
        doc = Uidoc.Document
        app = UiApp.Application

        Dim strcl As New List(Of Element)
        Dim ide As New List(Of ElementId)
        Dim pickedElements As IList(Of Element) = Uidoc.Selection.PickElementsByRectangle("Chọn một khoảng") '' Lấy 1 list các thành phần 
        If pickedElements.Count > 0 Then ' nếu chọn được các thành phần 
            Dim SoLuongColumns As Integer
            Dim idsToSelect As IList(Of ElementId) = New List(Of ElementId)(pickedElements.Count)
            For Each element As Element In pickedElements
                If element.Category.Name() = "Structural Columns" Then
                    SoLuongColumns = SoLuongColumns + 1
                    strcl.Add(element)
                    ide.Add(element.Id)
                End If
                idsToSelect.Add(element.Id)
            Next
            Uidoc.Selection.SetElementIds(idsToSelect)
            TaskDialog.Show("Số lượng cột trong khoảng chọn", SoLuongColumns)
        End If
        For Each id As ElementId In ide
            Dim defaultType As FamilyInstance = TryCast(doc.GetElement(id), FamilyInstance)
            'TaskDialog.Show("a", GetFamilyName(defaultType)) 'Lấy ra Family type
        Next
        Return Result.Succeeded
    End Function
    Public Shared Sub GetInfoForSymbols(family As Family)
        Dim message As New StringBuilder("Selected element's family name is : " & Convert.ToString(family.Name))
        Dim familySymbolIds As ISet(Of ElementId) = family.GetFamilySymbolIds()

        If familySymbolIds.Count = 0 Then
            message.AppendLine("Contains no family symbols.")
        Else
            message.AppendLine("The family symbols contained in this family are : ")

            ' Get family symbols which is contained in this family
            For Each id As ElementId In familySymbolIds
                Dim familySymbol As FamilySymbol = TryCast(family.Document.GetElement(id), FamilySymbol)
                ' Get family symbol name

                message.AppendLine(vbLf & "Name: " + familySymbol.Name)
                For Each materialId As ElementId In familySymbol.GetMaterialIds(False)
                    Dim material As Material = TryCast(familySymbol.Document.GetElement(materialId), Material)
                    message.AppendLine(vbLf & "Material : " + material.Name)
                Next
            Next
        End If

        TaskDialog.Show("Revit", message.ToString())
    End Sub
    Private Sub AssignDefaultTypeToColumn(document As Document, column As FamilyInstance)
        Dim defaultTypeId As ElementId = document.GetDefaultFamilyTypeId(New ElementId(BuiltInCategory.OST_StructuralColumns))
        If defaultTypeId <> ElementId.InvalidElementId Then
            Dim defaultType As FamilySymbol = TryCast(document.GetElement(defaultTypeId), FamilySymbol)
            If defaultType IsNot Nothing Then
                column.Symbol = defaultType
            End If
        End If
    End Sub
    Public Shared Function GetFamilyName(ByVal e As Element) As String
        Dim eId = e?.GetTypeId()
        If eId Is Nothing Then Return ""
        Dim elementType = TryCast(e.Document.GetElement(eId), ElementType)
        Return If(elementType?.FamilyName, "")
    End Function
End Class
