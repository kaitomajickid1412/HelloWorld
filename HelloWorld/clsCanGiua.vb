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
Public Class clsCanGiua
    Implements IExternalCommand
    Dim doc As Document
    Dim app As Application
    Public Function Execute(commandData As ExternalCommandData, ByRef message As String, elements As ElementSet) As Result Implements IExternalCommand.Execute
        Dim UiApp As UIApplication = commandData.Application
        Dim Uidoc As UIDocument = UiApp.ActiveUIDocument
        doc = Uidoc.Document
        app = UiApp.Application
        Try
            Dim TagLamChuan As Element = Nothing
            Dim ref As Reference = Uidoc.Selection.PickObject(ObjectType.Element, "Select a target tag object.")
            Dim elePick As Element = doc.GetElement(ref)
            If TypeOf elePick Is MultiReferenceAnnotation Then
                Dim MTag As MultiReferenceAnnotation = TryCast(elePick, MultiReferenceAnnotation)
                TagLamChuan = TryCast(doc.GetElement(MTag.TagId), IndependentTag)
            ElseIf TypeOf elePick Is IndependentTag Then
                TagLamChuan = TryCast(elePick, IndependentTag)
            End If
            If TagLamChuan IsNot Nothing Then
                Dim ListTagSelect_Multi As New List(Of Element)()
                Dim ListTagSelect_Indep As New List(Of Element)()
                Dim ListSelect As New List(Of Element)
                Dim ListReference As List(Of Reference) = Uidoc.Selection.PickObjects(ObjectType.Element, "Select other tag objects to align.")
                For Each refele As Reference In ListReference
                    ListSelect.Add(doc.GetElement(refele))
                Next

                For Each elem As Autodesk.Revit.DB.Element In ListSelect
                    If TypeOf elem Is MultiReferenceAnnotation Then
                        Dim MTag As MultiReferenceAnnotation = TryCast(elem, MultiReferenceAnnotation)

                        ListTagSelect_Multi.Add(doc.GetElement(MTag.TagId))
                    ElseIf TypeOf elem Is IndependentTag Then
                        ListTagSelect_Indep.Add(elem)
                        '' thuộc tính leader của independentTags

                    End If
                Next

                '20-06-19 Sua AlignTags (cho truong hop dung leader )
                ' can tag voi tags la Multi

                If ListTagSelect_Indep.Count <> 0 Then
                    Dim AcView As View = doc.ActiveView
                    Dim ToaDoChuan As XYZ = XacDinhDiemTag_Giua(TagLamChuan, AcView)

                    Using t As Transaction = New Transaction(doc, "CanTag")
                        t.Start()
                        For i As Integer = 0 To ListTagSelect_Indep.Count - 1
                            Dim Tag As Element = ListTagSelect_Indep(i)

                            ' kiem tra xem tags co dung paramater Leader hay k
                            If Tag.LookupParameter("Leader Line").AsInteger = 0 Then
                                Dim ToaDoCu As XYZ = XacDinhDiemTag_Giua(Tag, AcView)

                                If Math.Round(AcView.UpDirection.X, 2) = 0 And Math.Round(AcView.UpDirection.Z, 2) = 0 Then
                                    'Truc Y Thang dung, Xác định tọa độ cùng phương X
                                    Dim NewPos As XYZ = New XYZ(ToaDoChuan.X, ToaDoCu.Y, ToaDoCu.Z)
                                    Dim Vector As XYZ = New XYZ(NewPos.X - ToaDoCu.X, NewPos.Y - ToaDoCu.Y, NewPos.Z - ToaDoCu.Z)
                                    Tag.Location.Move(Vector)
                                ElseIf Math.Round(AcView.UpDirection.Y, 2) = 0 And Math.Round(AcView.UpDirection.Z, 2) = 0 Then
                                    Dim NewPos As XYZ = New XYZ(ToaDoCu.X, ToaDoChuan.Y, ToaDoCu.Z)
                                    Dim Vector As XYZ = New XYZ(NewPos.X - ToaDoCu.X, NewPos.Y - ToaDoCu.Y, NewPos.Z - ToaDoCu.Z)
                                    Tag.Location.Move(Vector)
                                    'Trục X thẳng đứng, Xác định tọa độ cùng phương Y
                                    'ElseIf Math.Round(AcView.RightDirection.X, 2) <> 0 Then
                                ElseIf Math.Round(AcView.UpDirection.X, 2) = 0 And Math.Round(AcView.UpDirection.Y, 2) = 0 Then
                                    'Truc Z thẳng đứng chia 2 trường hợp xem trục Y nằm ngàng hay trục X nằm ngang\
                                    If Math.Round(AcView.RightDirection.X, 2) = 0 Then
                                        Dim NewPos As XYZ = New XYZ(ToaDoCu.X, ToaDoChuan.Y, ToaDoCu.Z)
                                        Dim Vector As XYZ = New XYZ(NewPos.X - ToaDoCu.X, NewPos.Y - ToaDoCu.Y, NewPos.Z - ToaDoCu.Z)
                                        Tag.Location.Move(Vector)
                                    Else
                                        Dim NewPos As XYZ = New XYZ(ToaDoChuan.X, ToaDoCu.Y, ToaDoCu.Z)
                                        Dim Vector As XYZ = New XYZ(NewPos.X - ToaDoCu.X, NewPos.Y - ToaDoCu.Y, NewPos.Z - ToaDoCu.Z)
                                        Tag.Location.Move(Vector)
                                    End If
                                End If

                            Else

                                ' cần bổ leader mới cantag được

                            End If
                        Next
                        t.Commit()

                    End Using
                End If
                ' Can Tags voi tag la Indep
                If ListTagSelect_Multi.Count <> 0 Then
                    Dim AcView As View = doc.ActiveView
                    Dim ToaDoChuan As XYZ = XacDinhDiemTag_Giua(TagLamChuan, AcView)


                    Using t As Transaction = New Transaction(doc, "CanTag")
                        t.Start()
                        For i As Integer = 0 To ListTagSelect_Multi.Count - 1
                            Dim Tag As Element = ListTagSelect_Multi(i)
                            Dim ToaDoCu As XYZ = XacDinhDiemTag_Giua(Tag, AcView)

                            If Math.Round(AcView.UpDirection.X, 2) = 0 And Math.Round(AcView.UpDirection.Z, 2) = 0 Then
                                'Truc Y Thang dung, Xác định tọa độ cùng phương X
                                Dim NewPos As XYZ = New XYZ(ToaDoChuan.X, ToaDoCu.Y, ToaDoCu.Z)
                                Dim Vector As XYZ = New XYZ(NewPos.X - ToaDoCu.X, NewPos.Y - ToaDoCu.Y, NewPos.Z - ToaDoCu.Z)
                                Tag.Location.Move(Vector)
                            ElseIf Math.Round(AcView.UpDirection.Y, 2) = 0 And Math.Round(AcView.UpDirection.Z, 2) = 0 Then
                                Dim NewPos As XYZ = New XYZ(ToaDoCu.X, ToaDoChuan.Y, ToaDoCu.Z)
                                Dim Vector As XYZ = New XYZ(NewPos.X - ToaDoCu.X, NewPos.Y - ToaDoCu.Y, NewPos.Z - ToaDoCu.Z)
                                Tag.Location.Move(Vector)
                                'Trục X thẳng đứng, Xác định tọa độ cùng phương Y
                                'ElseIf Math.Round(AcView.RightDirection.X, 2) <> 0 Then
                            ElseIf Math.Round(AcView.UpDirection.X, 2) = 0 And Math.Round(AcView.UpDirection.Y, 2) = 0 Then
                                'Truc Z thẳng đứng chia 2 trường hợp xem trục Y nằm ngàng hay trục X nằm ngang\
                                If Math.Round(AcView.RightDirection.X, 2) = 0 Then
                                    Dim NewPos As XYZ = New XYZ(ToaDoCu.X, ToaDoChuan.Y, ToaDoCu.Z)
                                    Dim Vector As XYZ = New XYZ(NewPos.X - ToaDoCu.X, NewPos.Y - ToaDoCu.Y, NewPos.Z - ToaDoCu.Z)
                                    Tag.Location.Move(Vector)
                                Else
                                    Dim NewPos As XYZ = New XYZ(ToaDoChuan.X, ToaDoCu.Y, ToaDoCu.Z)
                                    Dim Vector As XYZ = New XYZ(NewPos.X - ToaDoCu.X, NewPos.Y - ToaDoCu.Y, NewPos.Z - ToaDoCu.Z)
                                    Tag.Location.Move(Vector)
                                End If
                            End If
                        Next
                        t.Commit()

                    End Using
                End If



            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return Result.Succeeded
    End Function
    Public Function XacDinhDiemTag_Giua(ByVal Tag As Element, ByVal AcView As View) As XYZ
        Dim Boundingbox As BoundingBoxXYZ = Tag.BoundingBox(AcView)
        Dim Min As XYZ = Boundingbox.Min
        Dim Max As XYZ = Boundingbox.Max
        XacDinhDiemTag_Giua = New XYZ((Max.X + Min.X) / 2, (Max.Y + Min.Y) / 2, (Max.Z + Min.Z) / 2)
    End Function
End Class
