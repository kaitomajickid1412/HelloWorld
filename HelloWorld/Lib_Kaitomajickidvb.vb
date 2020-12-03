Imports System.Collections.Generic
Imports System.Text
Imports System.Linq
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.UI.Selection
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.DB.Structure
Imports Autodesk.Revit.DB.Architecture
Imports Autodesk.Revit.Creation.ItemFactoryBase

<Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)>
Public Class Lib_Kaitomajickidvb
    Implements IExternalCommand
    Dim doc As Document
    Dim app As Application
    Public Function Execute(ByVal commandData As ExternalCommandData, ByRef message As String, ByVal elements As ElementSet) As Result Implements IExternalCommand.Execute
        Dim geo As New Autodesk.Revit.DB.XYZ
        Dim doc As Document = commandData.Application.ActiveUIDocument.Document
        Dim UiApp As UIApplication = commandData.Application
        Dim Uidoc As UIDocument = UiApp.ActiveUIDocument
        doc = Uidoc.Document
        app = UiApp.Application
        'CreateColumnsTructure(doc, PickPoint(Uidoc), "Level 1") 'done
        'CreateWalls(doc, "Level 1", PickPoint(Uidoc), PickPoint(Uidoc), True) 'done, Struc = True, Arc = False
        'Getinfo_Level(doc)
        'CreateBeam(doc, "Level 1", PickPoint(Uidoc), PickPoint(Uidoc))
        'CreateFloor(doc, "Level 1", PickPoint(Uidoc), PickPoint(Uidoc), True) 'done, Struc = True, Arc  = False
        'SetParameter(Uidoc, "Top Offset")
        'MoveElement(Uidoc)
        'RotateElement(Uidoc, 45)
        'Getgeometry(Uidoc) 'Lấy thông tin hình học của vật thể
        'GetdeliverySolid(Uidoc) ' tìm khối giao nhau
        'GetPointintersect(Uidoc, New XYZ(0, 0, 1)) 'Tìm giao của 2 vật thể và khoảng cách
        'CreateViewPlan(Uidoc, "Level 1") ' Tạo lkhung nhìn 
        'CreateFilter(Uidoc, "Section 1") ' tạo filter cho section 
        'Tagelement(Uidoc) 'tạo tag cho element  
        'CreateSheet(Uidoc, "kaitomajickid1412", "1412-01") ' tạo sheet khung bản vẽ
        'CreateView(Uidoc, "kaitomajickid1412", "Level 1") ' tạo viewport trong sheet
        'CreateReBarCoLumns(Uidoc, "M_00", "22M", 50) ' Đặt thép cho cột thẳng đứng
        'CreateRooms(doc)
        'CreateRoom(doc, "Level 1")
        BoundaryExtrude(Uidoc, doc, 0, 10)

        Return Result.Succeeded
    End Function
    Public Function PickPoint(uidoc As UIDocument) As XYZ
        Dim snapTypes As ObjectSnapTypes = ObjectSnapTypes.Endpoints Or ObjectSnapTypes.Intersections
        Dim point As XYZ = uidoc.Selection.PickPoint(snapTypes, "Select an end point or intersection")
        Dim strCoords As String = "Selected point is " & point.ToString()
        'TaskDialog.Show("Revit", strCoords)
        Return point
    End Function
    Public Function FindColumnSymbol(ByVal doc As Document) As Autodesk.Revit.DB.FamilySymbol
        Dim collector As FilteredElementCollector = New FilteredElementCollector(doc)
        collector.OfCategory(BuiltInCategory.OST_Columns).OfClass(GetType(FamilySymbol))
        Dim symbol As FamilySymbol = Nothing
        If (collector.ToElementIds().Count > 0) Then
            symbol = collector.FirstElement()
        End If
        Return symbol
    End Function
    Public Function FindBeamSymbol(ByVal doc As Document) As Autodesk.Revit.DB.FamilySymbol
        Dim collector As FilteredElementCollector = New FilteredElementCollector(doc)
        collector.OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(GetType(FamilySymbol))
        Dim symbol As FamilySymbol = Nothing
        If (collector.ToElementIds().Count > 0) Then
            symbol = collector.FirstElement()
        End If
        Return symbol
    End Function
    Public Function FindWallSymbol(ByVal doc As Document) As Autodesk.Revit.DB.FamilySymbol
        Dim collector As FilteredElementCollector = New FilteredElementCollector(doc)
        collector.OfCategory(BuiltInCategory.OST_WallsStructure).OfClass(GetType(FamilySymbol))
        Dim symbol As FamilySymbol = Nothing
        If (collector.ToElementIds().Count > 0) Then
            symbol = collector.FirstElement()
        End If
        Return symbol
    End Function
    Public Function FindFloorType(ByVal doc As Document) As FloorType
        Dim collector As New FilteredElementCollector(doc)
        collector.OfClass(GetType(FloorType))
        Dim floorType As FloorType = Nothing
        If (collector.ToElementIds().Count > 0) Then
            floorType = TryCast(collector.FirstElement(), FloorType)
        End If
        Return floorType
    End Function
    Public Function FindLevel(ByVal doc As Document, TenLevel As String) As Autodesk.Revit.DB.Level
        Dim collector As FilteredElementCollector = New FilteredElementCollector(doc)
        Dim level As Level = collector.OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().Cast(Of Level).First(Function(x) x.Name = TenLevel)
        Return level
    End Function
    Public Sub CreateColumnsTructure(ByVal doc As Document, ToaDo As XYZ, TenLevel As String)
        Dim sym As Autodesk.Revit.DB.FamilySymbol = FindColumnSymbol(doc)
        Dim level As Autodesk.Revit.DB.Level = FindLevel(doc, TenLevel)
        Try
            Using tr As New Transaction(doc, "Columns")
                tr.Start()
                sym.Activate()
                Dim col As Autodesk.Revit.DB.FamilyInstance = doc.Create.NewFamilyInstance(ToaDo, sym, level, [Structure].StructuralType.Column)
                If col Is Nothing Then
                    MsgBox("Không thể tạo được cột")
                End If
                tr.Commit()
            End Using
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
    End Sub
    Public Sub CreateWalls(ByVal doc As Document, TenLevel As String, ToaDo1 As XYZ, ToaDo2 As XYZ, StrorArc As Boolean)
        Dim level As Autodesk.Revit.DB.Level = FindLevel(doc, TenLevel)
        Dim BeamLine As Line = Line.CreateBound(ToaDo1, ToaDo2)
        Try
            Using tr As New Transaction(doc, "place family")
                tr.Start()
                Dim col As Wall = Wall.Create(doc, BeamLine, level.Id, StrorArc)
                If col Is Nothing Then
                    MsgBox("Không thể tạo được dầm")
                End If
                tr.Commit()
            End Using
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
    End Sub
    Public Sub CreateBeam(ByVal doc As Document, TenLevel As String, ToaDo1 As XYZ, ToaDo2 As XYZ)
        Dim sym As Autodesk.Revit.DB.FamilySymbol = FindBeamSymbol(doc)
        Dim level As Autodesk.Revit.DB.Level = FindLevel(doc, TenLevel)
        Dim line As Line = Line.CreateBound(ToaDo1, ToaDo2)
        Try
            Using tr As New Transaction(doc, "Beams")
                tr.Start()
                sym.Activate()
                Dim col As Autodesk.Revit.DB.FamilyInstance = doc.Create.NewFamilyInstance(line, sym, level, [Structure].StructuralType.Beam)
                If col Is Nothing Then
                    MsgBox("Không thể tạo được dầm")
                End If
                tr.Commit()
            End Using
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
    End Sub
    Public Sub CreateFloor(ByVal doc As Document, TenLevel As String, ToaDo1 As XYZ, ToaDo3 As XYZ, StrucorArc As Boolean)
        Dim FloorT As FloorType = FindFloorType(doc)
        Dim level As Autodesk.Revit.DB.Level = FindLevel(doc, TenLevel)
        Dim ToaDo2 As New XYZ(ToaDo3.X, ToaDo1.Y, 0)
        Dim ToaDo4 As New XYZ(ToaDo1.X, ToaDo3.Y, 0)
        Dim profile As New CurveArray()
        profile.Append(Line.CreateBound(ToaDo1, ToaDo2))
        profile.Append(Line.CreateBound(ToaDo2, ToaDo3))
        profile.Append(Line.CreateBound(ToaDo3, ToaDo4))
        profile.Append(Line.CreateBound(ToaDo4, ToaDo1))
        Dim normal As XYZ = XYZ.BasisZ
        Try
            Using tr As New Transaction(doc, "Floor")
                tr.Start()
                doc.Create.NewFloor(profile, FloorT, level, StrucorArc, normal)
                tr.Commit()
            End Using
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
    End Sub
    Public Function GetInfoParameter(ByVal Uidoc As UIDocument, r As Reference, TenParameter As String)
        Dim doc As Document = Uidoc.Document
        Dim kq As BuiltInParameter
        Try
            If r IsNot Nothing Then
                Dim elementid As ElementId = r.ElementId
                Dim element As Element = doc.GetElement(elementid)

                Dim para As Parameter = element.LookupParameter(TenParameter)
                Dim def As InternalDefinition = para.Definition
                kq = def.BuiltInParameter

            Else

            End If
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
        Return kq
    End Function
    Public Function GetElement(ByVal Uidoc As UIDocument, r As Reference) As Element
        Dim elementid As ElementId = r.ElementId
        Dim kq As Element = doc.GetElement(elementid)
        Return kq
    End Function
    Public Sub SetParameter(ByVal UIDoc As UIDocument, TenParameter As String)
        Dim doc As Document = UIDoc.Document
        Dim r As Reference = GetRefPickObject(UIDoc)
        Dim elementid As ElementId = r.ElementId
        Dim element As Element = doc.GetElement(elementid)
        Dim BuiltPara As BuiltInParameter = GetInfoParameter(UIDoc, r, TenParameter)
        Try
            Dim para As Parameter = element.Parameter(BuiltPara)
            Using tr As New Transaction(doc, "set para")
                tr.Start()
                para.Set(6.5)
                tr.Commit()
            End Using
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try

    End Sub
    Public Shared Function GetRefPickObject(ByVal Uidoc As UIDocument) As Reference
        Dim r As Reference
        Try
            r = Uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element)
        Catch ex As Exception
            r = Nothing
        End Try
        Return r
    End Function
    Public Sub MoveElement(UIdoc As UIDocument)
        Dim doc As Document = UIdoc.Document
        Dim r As Reference = GetRefPickObject(UIdoc)
        Dim elementid As ElementId = r.ElementId
        Dim element As Element = doc.GetElement(elementid)
        If r IsNot Nothing Then
            Try

                Using tr As New Transaction(doc, "move element")
                    tr.Start()
                    Dim VecToMove As XYZ = PickPoint(UIdoc)
                    ElementTransformUtils.MoveElement(doc, elementid, VecToMove)
                    tr.Commit()
                End Using
            Catch ex As Exception
                Dim [error] As String = ex.Message
            End Try
        End If
    End Sub
    Public Sub RotateElement(UIdoc As UIDocument, GocXoay As Double)
        Dim doc As Document = UIdoc.Document
        Dim r As Reference = GetRefPickObject(UIdoc)
        Dim elementid As ElementId = r.ElementId
        Dim element As Element = doc.GetElement(elementid)
        If r IsNot Nothing Then
            Try

                Using tr As New Transaction(doc, "rotation element")
                    tr.Start()
                    Dim loc As LocationPoint = element.Location
                    Dim P1 As XYZ = loc.Point
                    Dim P2 As XYZ = New XYZ(P1.X, P1.Y, P1.Z + 1)
                    Dim axis As Line = Line.CreateBound(P1, P2)
                    Dim Angle As Double = GocXoay * (Math.PI / 180)
                    ElementTransformUtils.RotateElement(doc, elementid, axis, Angle)
                    tr.Commit()
                End Using

            Catch ex As Exception
                Dim [error] As String = ex.Message
            End Try
        End If
    End Sub
    Public Sub DeleteElementUseid(document As Autodesk.Revit.DB.Document, ele As Element)
        Dim elementId As Autodesk.Revit.DB.ElementId = ele.Id
        Dim deletedIdSet As ICollection(Of Autodesk.Revit.DB.ElementId) = document.Delete(elementId)
        MsgBox(1)
    End Sub
    Public Sub DeleteEleSelected(ByVal UIdoc As UIDocument, ByVal lstDelete As List(Of ElementId))
        Using t As Transaction = New Transaction(UIdoc.Document, "Delete")
            t.Start()
            UIdoc.Document.Delete(lstDelete)
            t.Commit()
        End Using
    End Sub
    Public Sub Getgeometry(UIdoc As UIDocument)
        Dim doc As Document = UIdoc.Document
        Dim r As Reference = GetRefPickObject(UIdoc)

        If r IsNot Nothing Then
            Try

                Using tr As New Transaction(doc, "Get geometry")
                    tr.Start()
                    Dim elementid As ElementId = r.ElementId
                    Dim element As Element = doc.GetElement(elementid)
                    Dim opt As Options = New Options()
                    opt.DetailLevel = ViewDetailLevel.Fine
                    Dim GeoEle As GeometryElement = element.Geometry(opt)
                    For Each obj As GeometryObject In GeoEle
                        Dim Solid As Solid = TryCast(obj, Solid)
                        Dim Face As Integer = 0
                        Dim area As Double = 0.0
                        For Each f As Face In Solid.Faces
                            area += f.Area
                            Face += 1

                        Next
                        TaskDialog.Show("Geometry", String.Format("Vật thể đã chọn có số mặt là {0}, tổng diện tích các mặt là {1}", Face, UnitUtils.ConvertFromInternalUnits(area, DisplayUnitType.DUT_SQUARE_METERS)))
                    Next

                    tr.Commit()
                End Using

            Catch ex As Exception
                Dim [error] As String = ex.Message
            End Try
        End If
    End Sub
    Public Sub GetdeliverySolid(UIdoc As UIDocument)
        Dim doc As Document = UIdoc.Document
        Dim r As Reference = GetRefPickObject(UIdoc)

        If r IsNot Nothing Then
            Try

                Using tr As New Transaction(doc, "Get geometry delivery")
                    tr.Start()
                    Dim elementid As ElementId = r.ElementId
                    Dim element As Element = doc.GetElement(elementid)
                    Dim opt As Options = New Options()
                    opt.DetailLevel = ViewDetailLevel.Fine
                    opt.ComputeReferences = True
                    Dim Solid As Solid = Nothing
                    Dim listSolid As List(Of Solid) = GetElementSolids(element, opt)
                    Solid = listSolid(0)
                    Dim colletor As FilteredElementCollector = New FilteredElementCollector(doc)
                    Dim filter As ElementIntersectsSolidFilter = New ElementIntersectsSolidFilter(Solid)
                    Dim intersection As ICollection(Of ElementId) = colletor.OfCategory(BuiltInCategory.OST_Roofs).WherePasses(filter).ToElementIds() 'càn thay đổi OST_roofs
                    UIdoc.Selection.SetElementIds(intersection)
                    tr.Commit()
                End Using
            Catch ex As Exception
                Dim [error] As String = ex.Message
            End Try
        End If
    End Sub
    Public Shared Function GetElementSolids(ByVal elem As Element, ByVal Optional opt As Options = Nothing, ByVal Optional useOriginGeom4FamilyInstance As Boolean = False) As List(Of Solid)
        If elem Is Nothing Then
            Return Nothing
        End If

        If opt Is Nothing Then opt = New Options()
        Dim solids As New List(Of Solid)
        Dim gElem As GeometryElement = Nothing

        Try

            If useOriginGeom4FamilyInstance AndAlso TypeOf elem Is FamilyInstance Then
                Dim fInst As FamilyInstance = TryCast(elem, FamilyInstance)
                gElem = fInst.GetOriginalGeometry(opt)
                Dim trf As Transform = fInst.GetTransform()
                If Not trf.IsIdentity Then gElem = gElem.GetTransformed(trf)
            Else
                gElem = elem.Geometry(opt)
            End If

            If gElem Is Nothing Then
                Return Nothing
            End If

            Dim gIter As IEnumerator(Of GeometryObject) = gElem.GetEnumerator()
            gIter.Reset()

            While gIter.MoveNext()
                solids.AddRange(GetSolids(gIter.Current))
            End While

        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
        Return solids
    End Function
    Public Shared Function GetSolids(ByVal gObj As GeometryObject) As List(Of Solid)
        Dim solids As List(Of Solid) = New List(Of Solid)()

        If TypeOf gObj Is Solid Then
            Dim solid As Solid = TryCast(gObj, Solid)
            If solid.Faces.Size > 0 AndAlso Math.Abs(solid.Volume) > 0 Then solids.Add(TryCast(gObj, Solid))
        ElseIf TypeOf gObj Is GeometryInstance Then
            Dim gIter2 As IEnumerator(Of GeometryObject) = (TryCast(gObj, GeometryInstance)).GetInstanceGeometry().GetEnumerator()
            gIter2.Reset()

            While gIter2.MoveNext()
                solids.AddRange(GetSolids(gIter2.Current))
            End While
        ElseIf TypeOf gObj Is GeometryElement Then
            Dim gIter2 As IEnumerator(Of GeometryObject) = (TryCast(gObj, GeometryElement)).GetEnumerator()
            gIter2.Reset()

            While gIter2.MoveNext()
                solids.AddRange(GetSolids(gIter2.Current))
            End While
        End If
        Return solids
    End Function
    Public Shared Sub GetPointintersect(Uidoc As UIDocument, TiaRay As XYZ)
        Dim doc As Document = Uidoc.Document
        Dim r As Reference = GetRefPickObject(Uidoc)

        If r IsNot Nothing Then
            Try
                Using tr As New Transaction(doc, "Get point intersect")
                    tr.Start()
                    Dim elementid As ElementId = r.ElementId
                    Dim element As Element = doc.GetElement(elementid)
                    Dim ray As XYZ = TiaRay 'tạo tia ray
                    'xác định gốc vector
                    Dim lockpoint As LocationPoint = TryCast(element.Location, LocationPoint)
                    Dim projectRay = lockpoint.Point
                    Dim filter As ElementCategoryFilter = New ElementCategoryFilter(BuiltInCategory.OST_Roofs)
                    Dim refinter As ReferenceIntersector = New ReferenceIntersector(filter, FindReferenceTarget.Face, CType(doc.ActiveView, View3D))
                    Dim refcontect As ReferenceWithContext = refinter.FindNearest(projectRay, ray)
                    Dim ref As Reference = refcontect.GetReference()
                    Dim intpoint As XYZ = ref.GlobalPoint
                    Dim distance As Double = projectRay.DistanceTo(intpoint) ' tính khoảng cách từ điểm gốc của element đến giao điểm 
                    TaskDialog.Show("Intersection", String.Format("Điểm cắt từ cột vào mái là {0}, khoảng cách giữa location point của cột và điểm giao là " + "{1}", intpoint, distance))
                    tr.Commit()
                End Using
            Catch ex As Exception
                Dim [error] As String = ex.Message
            End Try
        End If
    End Sub
    Public Sub CreateViewPlan(UIdoc As UIDocument, TenLevel As String)
        Dim doc As Document = UIdoc.Document
        Dim level As Autodesk.Revit.DB.Level = FindLevel(doc, TenLevel)
        ' lấy family 
        Dim VFamily As ViewFamilyType = New FilteredElementCollector(doc).OfClass(GetType(ViewFamilyType)).Cast(Of ViewFamilyType)().First(Function(x) x.ViewFamily = ViewFamily.FloorPlan)


        Try
            Using tr As New Transaction(doc, "View Plan")
                tr.Start()
                Dim ViewP = ViewPlan.Create(doc, VFamily.Id, level.Id)
                ViewP.Name = "Khung nhìn mặt sàn từ revit API"
                tr.Commit()
            End Using
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
    End Sub
    Public Sub CreateFilter(UIdoc As UIDocument, Name As String)
        Dim doc As Document = UIdoc.Document
        ' tạo categories
        Dim cast As New List(Of ElementId)
        cast.Add(New ElementId(BuiltInCategory.OST_Sections))
        ' tạo element filter
        Dim elefilter As ElementParameterFilter = New ElementParameterFilter(ParameterFilterRuleFactory.CreateContainsRule(New ElementId(BuiltInParameter.VIEW_NAME), Name, False))
        Try
            Using tr As New Transaction(doc, "Create Filter")
                tr.Start()
                Dim filterele As ParameterFilterElement = ParameterFilterElement.Create(doc, "My Filter", cast, elefilter)
                doc.ActiveView.AddFilter(filterele.Id)
                doc.ActiveView.SetFilterVisibility(filterele.Id, False)
                tr.Commit()
            End Using
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
    End Sub
    Public Sub Tagelement(Uidoc As UIDocument)
        Dim doc As Document = Uidoc.Document
        Dim tMode As TagMode = TagMode.TM_ADDBY_CATEGORY
        Dim tOrientation As TagOrientation = TagOrientation.Horizontal
        ' tạo các category
        Dim cats As New List(Of BuiltInCategory)
        cats.Add(BuiltInCategory.OST_Windows)
        cats.Add(BuiltInCategory.OST_Doors)

        Dim EleMucateFilter As ElementMulticategoryFilter = New ElementMulticategoryFilter(cats)
        Dim list As List(Of Element) = New FilteredElementCollector(doc, doc.ActiveView.Id).WherePasses(EleMucateFilter).WhereElementIsNotElementType().ToElements()

        Try
            Using tr As New Transaction(doc, "Create TagElement")
                tr.Start()
                For Each e As Element In list
                    Dim ref As Reference = New Reference(e)
                    Dim loc As LocationPoint = TryCast(e.Location, LocationPoint)
                    Dim pos As XYZ = loc.Point
                    Dim tag As IndependentTag = IndependentTag.Create(doc, doc.ActiveView.Id, ref, True, tMode, tOrientation, pos)

                Next
                tr.Commit()
            End Using
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
    End Sub
    Public Sub CreateSheet(Uidoc As UIDocument, Name As String, SheetNumber As String)
        Dim doc As Document = Uidoc.Document
        ' lấy family tile block
        Dim TitleBlock As FamilySymbol = New FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().Cast(Of FamilySymbol).First()
        Try
            Using tr As New Transaction(doc, "Create Sheet")
                tr.Start()

                Dim sheet As ViewSheet = ViewSheet.Create(doc, TitleBlock.Id)
                sheet.Name = Name
                sheet.SheetNumber = SheetNumber


                tr.Commit()
            End Using
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
    End Sub
    Public Sub CreateView(UIdoc As UIDocument, Name As String, TenLevel As String)
        Dim doc As Document = UIdoc.Document
        ' Lấy Viewsheet
        Dim sheet As ViewSheet = New FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().Cast(Of ViewSheet).First(Function(x) x.Name = Name)
        'Lấy view
        Dim view As Element = New FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().Cast(Of View).First(Function(x) x.Name = TenLevel)
        'Lấy điểm đặt của view
        Dim uv As BoundingBoxUV = sheet.Outline
        Dim ux As Double = (uv.Max.U + uv.Min.U) / 2
        Dim uy As Double = (uv.Max.V + uv.Min.V) / 2
        Dim point As XYZ = New XYZ(ux, uy, 0)
        Try
            Using tr As New Transaction(doc, "Create ViewPort")
                tr.Start()
                Dim ViewPorrt As Viewport = Viewport.Create(doc, sheet.Id, view.Id, point)

                tr.Commit()
            End Using
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
    End Sub
    Public Sub CreateReBarCoLumns(Uidoc As UIDocument, TenThep As String, KieuThep As String, LopBaoVe As Double)
        Dim doc As Document = Uidoc.Document
        'lấy rebar 
        Dim rbShape As RebarShape = New FilteredElementCollector(doc).OfClass(GetType(RebarShape)).Cast(Of RebarShape).First(Function(x) x.Name = TenThep)

        'lấy rebartype 
        Dim barType As RebarBarType = New FilteredElementCollector(doc).OfClass(GetType(RebarBarType)).Cast(Of RebarBarType).First(Function(x) x.Name = KieuThep)
        'lấy đối tượng
        Dim r As Reference = GetRefPickObject(Uidoc)
        Dim ele As Element = doc.GetElement(r)
        ' tạo lớp bảo vệ rebar
        Dim cover As Double = LopBaoVe
        Dim boudingbox As BoundingBoxXYZ = ele.BoundingBox(Nothing)
        MsgBox(UnitUtils.ConvertToInternalUnits(boudingbox.Max.X, DisplayUnitType.DUT_MILLIMETERS))
        MsgBox(UnitUtils.ConvertToInternalUnits(boudingbox.Max.Y, DisplayUnitType.DUT_MILLIMETERS))
        MsgBox(UnitUtils.ConvertToInternalUnits(boudingbox.Min.X, DisplayUnitType.DUT_MILLIMETERS))
        MsgBox(UnitUtils.ConvertToInternalUnits(boudingbox.Min.Y, DisplayUnitType.DUT_MILLIMETERS))
        Dim origin_1st As XYZ = New XYZ(boudingbox.Min.X + UnitUtils.ConvertToInternalUnits(cover, DisplayUnitType.DUT_MILLIMETERS),
                                        boudingbox.Min.Y + UnitUtils.ConvertToInternalUnits(cover, DisplayUnitType.DUT_MILLIMETERS),
                                        boudingbox.Min.Z)
        Dim origin_2st As XYZ = New XYZ(boudingbox.Max.X - UnitUtils.ConvertToInternalUnits(cover, DisplayUnitType.DUT_MILLIMETERS),
                                        boudingbox.Max.Y - UnitUtils.ConvertToInternalUnits(cover, DisplayUnitType.DUT_MILLIMETERS),
                                        boudingbox.Min.Z)
        Dim origin_3st As XYZ = New XYZ(boudingbox.Max.X - UnitUtils.ConvertToInternalUnits(cover, DisplayUnitType.DUT_MILLIMETERS),
                                        boudingbox.Min.Y + UnitUtils.ConvertToInternalUnits(cover, DisplayUnitType.DUT_MILLIMETERS),
                                        boudingbox.Min.Z)
        Dim origin_4st As XYZ = New XYZ(boudingbox.Min.X + UnitUtils.ConvertToInternalUnits(cover, DisplayUnitType.DUT_MILLIMETERS),
                                        boudingbox.Max.Y - UnitUtils.ConvertToInternalUnits(cover, DisplayUnitType.DUT_MILLIMETERS),
                                        boudingbox.Min.Z)

        'Vecto đặt thép
        Dim xVec As XYZ = New XYZ(0, 0, 1)
        Dim YVec As XYZ = New XYZ(0, 1, 0)

        Try
        Using tr As New Transaction(doc, "Create Rebar Columns")
            tr.Start()

                Rebar.CreateFromRebarShape(doc, rbShape, barType, ele, origin_1st, xVec, YVec)
                Rebar.CreateFromRebarShape(doc, rbShape, barType, ele, origin_2st, xVec, YVec)
                Rebar.CreateFromRebarShape(doc, rbShape, barType, ele, origin_3st, xVec, YVec)
                Rebar.CreateFromRebarShape(doc, rbShape, barType, ele, origin_4st, xVec, YVec)
                tr.Commit()
        End Using
        Catch ex As Exception
        Dim [error] As String = ex.Message
        End Try
    End Sub
    Private Sub Getinfo_Level(document As Document)
        Dim levelInformation As New StringBuilder()
        Dim levelNumber As Integer = 0
        Dim collector As New FilteredElementCollector(document)
        Dim collection As ICollection(Of Element) = collector.OfClass(GetType(Level)).ToElements()
        For Each e As Element In collection
            Dim level As Level = TryCast(e, Level)

            If level IsNot Nothing Then
                ' keep track of number of levels
                levelNumber += 1

                'get the name of the level
                levelInformation.Append(vbLf & "Level Name: " + level.Name)

                'get the elevation of the level
                levelInformation.Append(vbLf & vbTab & "Elevation: " + Convert.ToString(level.Elevation))

                ' get the project elevation of the level
                levelInformation.Append(vbLf & vbTab & "Project Elevation: " + Convert.ToString(level.ProjectElevation))
            End If
        Next

        'number of total levels in current document
        levelInformation.Append(vbLf & vbLf & " There are " & levelNumber & " levels in the document!")

        'show the level information in the messagebox
        TaskDialog.Show("Revit", levelInformation.ToString())
    End Sub
    Private Sub CreateBoudary(UIdoc As UIDocument)
        Dim doc As Document = UIdoc.Document
        Try
            Using tr As New Transaction(doc, "Create Rebar Columns")
                tr.Start()


                tr.Commit()
            End Using
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
    End Sub
    Private Sub PoLylineDragon(UIdoc As UIDocument)
        Dim doc As Document = UIdoc.Document
        Dim Goc As Double = Math.PI / 2
        Dim r As Reference = GetRefPickObject(UIdoc)

        Try
            Using tr As New Transaction(doc, "Create Rebar Columns")
                tr.Start()


                tr.Commit()
            End Using
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
    End Sub
    Private Sub BoundaryExtrude(UIdoc As UIDocument, document As Autodesk.Revit.DB.Document, z1 As Double, z2 As Double)
        Dim doc As Document = UIdoc.Document
        Dim ListStartPointModelLine As New List(Of XYZ)
        Dim ListEndPointModelLine As New List(Of XYZ)
        Dim ListDoanThangPoint As New List(Of List(Of XYZ))

        Dim lstDoanThangab As List(Of List(Of Double))

        Dim lstPointIntersect As New List(Of XYZ)
        Dim lst4PointTong As New List(Of List(Of XYZ))

        Dim z As Double
        Try
            Dim pickedElements As IList(Of Element) = UIdoc.Selection.PickElementsByRectangle("Chọn một khoảng")
            If pickedElements.Count > 0 Then ' nếu chọn được các thành phần 
                Dim idsToSelect As IList(Of ElementId) = New List(Of ElementId)(pickedElements.Count)

                For Each element As Element In pickedElements
                    If element.Category.Name() = "Lines" Then
                        idsToSelect.Add(element.Id)
                        Dim lineModel As ModelLine = TryCast(element, ModelLine)
                        Dim Line As Line = TryCast(lineModel.GeometryCurve, Line)
                        sketplane = lineModel.SketchPlane
                        Dim StartPoint As XYZ = Line.GetEndPoint(0)
                        Dim EndPoint As XYZ = Line.GetEndPoint(1)
                        z = StartPoint.Z
                        'MsgBox(StartPoint.X + " " + StartPoint.Y + " " + StartPoint.Z)
                        'MsgBox(EndPoint.X + " " + EndPoint.Y + " " + EndPoint.Z)
                        'ListStartPointModelLine.Add(StartPoint)
                        'ListEndPointModelLine.Add(EndPoint)
                        Dim ListTowPoint As New List(Of XYZ)
                        ListTowPoint.Add(StartPoint)
                        ListTowPoint.Add(EndPoint)
                        ListDoanThangPoint.Add(ListTowPoint)
                    End If
                Next
                UIdoc.Selection.SetElementIds(idsToSelect)

                DeleteEleSelected(UIdoc, idsToSelect)
                lstDoanThangab = Lineequations(ListDoanThangPoint)

                lstPointIntersect = Findintersectline(lstDoanThangab, z, ListDoanThangPoint)


                lst4PointTong = lstPointDetermind2(lstPointIntersect)

                Using tr As New Transaction(doc, "Extrude Boundary")
                    tr.Start()
                    'CreateModelLine(UIdoc)
                    For i = 0 To lst4PointTong.Count - 1
                        Dim lst4Point As List(Of XYZ) = lst4PointTong(i)
                        CreateExtrusion(doc, sketplane, lst4Point, z1, z2)
                    Next

                    lstCheckPoint.Clear()
                    tr.Commit()
                End Using
            End If

        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
    End Sub
    Public Sub CreateModelLine(UIdoc As UIDocument)
        Dim doc As Document = UIdoc.Document
        ' get handle to application from doc
        Dim application As Autodesk.Revit.ApplicationServices.Application = doc.Application

        ' Create a geometry line in Revit application
        Dim startPoint As New XYZ(0, 0, 0)
        Dim endPoint As New XYZ(10, 10, 0)
        Dim geomLine As Line = Line.CreateBound(startPoint, endPoint)

        ' Create a geometry arc in Revit application
        Dim end0 As New XYZ(1, 0, 0)
        Dim end1 As New XYZ(10, 10, 10)
        Dim pointOnCurve As New XYZ(10, 0, 0)
        Dim geomArc As Arc = Arc.Create(end0, end1, pointOnCurve)

        ' Create a geometry plane in Revit application
        Dim origin As New XYZ(0, 0, 0)
        Dim normal As New XYZ(1, 1, 0)
        Dim geomPlane As Plane = Plane.CreateByNormalAndOrigin(normal, origin)

        ' Create a sketch plane in current document
        Dim sketch As SketchPlane = SketchPlane.Create(doc, geomPlane)

        ' Create a ModelLine element using the created geometry line and sketch plane


        Using tr As New Transaction(doc, "Model")
            tr.Start()
            Dim line__1 As ModelLine = TryCast(doc.Create.NewModelCurve(geomLine, sketch), ModelLine)
            tr.Commit()
        End Using
    End Sub
    Public Function lstPointDetermind1(ByVal lstPoint As List(Of XYZ)) As List(Of List(Of XYZ))
        Dim kqTong As New List(Of List(Of XYZ))
        Dim lstPointold As New List(Of XYZ)
        Dim lstPointoldTong As New List(Of List(Of XYZ))
        For k = 0 To lstPoint.Count - 1
            'Dim kq As List(Of XYZ)
            Dim ListRanger As New List(Of Double)
            Dim lstPointNew As New List(Of XYZ)

            Dim lstPointCheck As New List(Of XYZ)
            Dim Check4Point As New List(Of XYZ)
            Dim DiemGoc As XYZ = lstPoint(k)
            'tạo list khoảng cách
            For i = 0 To lstPoint.Count - 1
                Dim a = DetermindRange(DiemGoc, lstPoint(i))
                ListRanger.Add(a)
            Next
            'Sắp xếp lại listpoint theo listkc tăng dần
            For i = 0 To lstPoint.Count - 2
                For j = i + 1 To lstPoint.Count - 1
                    If ListRanger(i) > ListRanger(j) Then
                        Dim tgkc As Double = ListRanger(i)
                        Dim tgXYZ As XYZ = lstPoint(i)
                        ListRanger(i) = ListRanger(j)
                        lstPoint(i) = lstPoint(j)
                        ListRanger(j) = tgkc
                        lstPoint(j) = tgXYZ
                    End If
                Next
            Next
            lstPointCheck.Add(DiemGoc)
            lstPointCheck.Add(lstPoint(1))
            Dim Skt As Double = Edge(lstPointCheck(0), lstPointCheck(1))
            If Skt > 0 Then
                Dim lstRanger As New List(Of Double)
                Dim LstPointClone As New List(Of XYZ)
                lstPointNew.Add(lstPoint(0))
                lstPointNew.Add(lstPoint(1))
                Dim lstCheck As New List(Of XYZ)
                For i = 0 To lstPoint.Count - 1
                    lstCheck.Add(lstPoint(i))
                Next
                lstCheck.RemoveAt(0)
                lstCheck.RemoveAt(0)
                Dim S As Double = 0
                For i = 2 To lstPoint.Count - 1
                    If lstPointNew.Count = 4 Then
                        Exit For
                    Else

                        For j = 0 To lstCheck.Count - 1
                            'If lstPoint(j).X = lstCheck(j).X And lstPoint(j).Y = lstCheck(j).Y Then
                            'Else
                            S = S + Edge(lstPointNew(lstPointNew.Count - 1), lstCheck(j))
                            Dim a = DetermindRange(lstPointNew(lstPointNew.Count - 1), lstCheck(j))
                            ' End If
                            If S < 0 Then
                                S = 0
                            Else
                                If LstPointClone.Count = 2 Then

                                    Exit For
                                Else
                                    lstRanger.Add(a)
                                    LstPointClone.Add(lstCheck(j))

                                End If
                            End If
                        Next
                        For h = 0 To LstPointClone.Count - 2
                            For q = h + 1 To LstPointClone.Count - 1
                                If lstRanger(h) > lstRanger(q) Then
                                    Dim tgkc As Double = lstRanger(h)
                                    Dim tgXYZ As XYZ = LstPointClone(h)
                                    lstRanger(h) = lstRanger(q)
                                    LstPointClone(h) = LstPointClone(q)
                                    lstRanger(q) = tgkc
                                    LstPointClone(q) = tgXYZ
                                End If
                            Next
                        Next
                        lstPointNew.Add(LstPointClone(0))
                        lstCheck.RemoveAt(RemovePoint(LstPointClone(0), lstCheck))
                        lstRanger.Clear()
                        LstPointClone.Clear()
                    End If
                Next

                If ChecklistPointTriangle(lstPointNew) Then
                    kqTong.Add(lstPointNew)
                Else

                End If
            ElseIf Skt < 0 Then
                Dim lstRanger As New List(Of Double)
                Dim LstPointClone As New List(Of XYZ)
                lstPointNew.Add(lstPoint(0))
                lstPointNew.Add(lstPoint(1))
                Dim lstCheck As New List(Of XYZ)
                For i = 0 To lstPoint.Count - 1
                    lstCheck.Add(lstPoint(i))
                Next
                lstCheck.RemoveAt(0)
                lstCheck.RemoveAt(0)
                Dim S As Double = 0
                For i = 2 To lstPoint.Count - 1
                    If lstPointNew.Count = 4 Then
                        Exit For
                    Else

                        For j = 0 To lstCheck.Count - 1
                            'If lstPoint(j).X = lstCheck(j).X And lstPoint(j).Y = lstCheck(j).Y Then
                            'Else
                            S = S + Edge(lstPointNew(lstPointNew.Count - 1), lstCheck(j))
                            Dim a = DetermindRange(lstPointNew(lstPointNew.Count - 1), lstCheck(j))
                            ' End If
                            If S > 0 Then
                                S = 0
                            Else
                                If LstPointClone.Count = 2 Then

                                    Exit For
                                Else
                                    lstRanger.Add(a)
                                    LstPointClone.Add(lstCheck(j))

                                End If
                            End If
                        Next
                        For h = 0 To LstPointClone.Count - 2
                            For q = h + 1 To LstPointClone.Count - 1
                                If lstRanger(h) > lstRanger(q) Then
                                    Dim tgkc As Double = lstRanger(h)
                                    Dim tgXYZ As XYZ = LstPointClone(h)
                                    lstRanger(h) = lstRanger(q)
                                    LstPointClone(h) = LstPointClone(q)
                                    lstRanger(q) = tgkc
                                    LstPointClone(q) = tgXYZ
                                End If
                            Next
                        Next
                        lstPointNew.Add(LstPointClone(0))
                        lstCheck.RemoveAt(RemovePoint(LstPointClone(0), lstCheck))
                        lstRanger.Clear()
                        LstPointClone.Clear()
                    End If
                Next


                If ChecklistPointTriangle(lstPointNew) Then
                    kqTong.Add(lstPointNew)
                Else

                End If
            ElseIf Skt = 0 Then
                lstPointCheck.Add(lstPoint(2))
                Dim Skt2 As Double = Edge(lstPointCheck(1), lstPointCheck(2))
                If Skt2 > 0 Then
                    Dim lstRanger As New List(Of Double)
                    Dim LstPointClone As New List(Of XYZ)
                    lstPointNew.Add(lstPoint(0))
                    lstPointNew.Add(lstPoint(1))
                    Dim lstCheck As New List(Of XYZ)
                    For i = 0 To lstPoint.Count - 1
                        lstCheck.Add(lstPoint(i))
                    Next
                    lstCheck.RemoveAt(0)
                    lstCheck.RemoveAt(0)
                    Dim S As Double = 0
                    For i = 2 To lstPoint.Count - 1
                        If lstPointNew.Count = 4 Then
                            Exit For
                        Else

                            For j = 0 To lstCheck.Count - 1
                                'If lstPoint(j).X = lstCheck(j).X And lstPoint(j).Y = lstCheck(j).Y Then
                                'Else
                                S = S + Edge(lstPointNew(lstPointNew.Count - 1), lstCheck(j))
                                Dim a = DetermindRange(lstPointNew(lstPointNew.Count - 1), lstCheck(j))
                                ' End If
                                If S < 0 Then
                                    S = 0
                                Else
                                    If LstPointClone.Count = 2 Then

                                        Exit For
                                    Else
                                        lstRanger.Add(a)
                                        LstPointClone.Add(lstCheck(j))

                                    End If
                                End If
                            Next
                            For h = 0 To LstPointClone.Count - 2
                                For q = h + 1 To LstPointClone.Count - 1
                                    If lstRanger(h) > lstRanger(q) Then
                                        Dim tgkc As Double = lstRanger(h)
                                        Dim tgXYZ As XYZ = LstPointClone(h)
                                        lstRanger(h) = lstRanger(q)
                                        LstPointClone(h) = LstPointClone(q)
                                        lstRanger(q) = tgkc
                                        LstPointClone(q) = tgXYZ
                                    End If
                                Next
                            Next
                            lstPointNew.Add(LstPointClone(0))
                            lstCheck.RemoveAt(RemovePoint(LstPointClone(0), lstCheck))
                            lstRanger.Clear()
                            LstPointClone.Clear()
                        End If
                    Next
                    If ChecklistPointTriangle(lstPointNew) Then
                        kqTong.Add(lstPointNew)
                    Else

                    End If
                ElseIf Skt2 < 0 Then
                    Dim lstRanger As New List(Of Double)
                    Dim LstPointClone As New List(Of XYZ)
                    lstPointNew.Add(lstPoint(0))
                    lstPointNew.Add(lstPoint(1))
                    Dim lstCheck As New List(Of XYZ)
                    For i = 0 To lstPoint.Count - 1
                        lstCheck.Add(lstPoint(i))
                    Next
                    lstCheck.RemoveAt(0)
                    lstCheck.RemoveAt(0)
                    Dim S As Double = 0
                    For i = 2 To lstPoint.Count - 1
                        If lstPointNew.Count = 4 Then
                            Exit For
                        Else

                            For j = 0 To lstCheck.Count - 1
                                'If lstPoint(j).X = lstCheck(j).X And lstPoint(j).Y = lstCheck(j).Y Then
                                'Else
                                S = S + Edge(lstPointNew(lstPointNew.Count - 1), lstCheck(j))
                                Dim a = DetermindRange(lstPointNew(lstPointNew.Count - 1), lstCheck(j))
                                ' End If
                                If S > 0 Then
                                    S = 0
                                Else
                                    If LstPointClone.Count = 2 Then

                                        Exit For
                                    Else
                                        lstRanger.Add(a)
                                        LstPointClone.Add(lstCheck(j))

                                    End If
                                End If
                            Next
                            For h = 0 To LstPointClone.Count - 2
                                For q = h + 1 To LstPointClone.Count - 1
                                    If lstRanger(h) > lstRanger(q) Then
                                        Dim tgkc As Double = lstRanger(h)
                                        Dim tgXYZ As XYZ = LstPointClone(h)
                                        lstRanger(h) = lstRanger(q)
                                        LstPointClone(h) = LstPointClone(q)
                                        lstRanger(q) = tgkc
                                        LstPointClone(q) = tgXYZ
                                    End If
                                Next
                            Next
                            lstPointNew.Add(LstPointClone(0))
                            lstCheck.RemoveAt(RemovePoint(LstPointClone(0), lstCheck))
                            lstRanger.Clear()
                            LstPointClone.Clear()
                        End If
                    Next
                    If ChecklistPointTriangle(lstPointNew) Then
                        kqTong.Add(lstPointNew)
                    Else

                    End If

                End If
                End If

            'lstPointold.Add(lstPoint(i))
            'lstPointoldTong.Add(lstPointold)
            ''lstCheckPoint.Add(Check4Point)
            'kq = lstPointNew
            'kqTong.Add(kq)

            'DeletePoint(lstPoint, DiemGoc)
        Next
        Return kqTong
    End Function
    Public Function RemovePoint(ByVal point As XYZ, ByVal lst As List(Of XYZ)) As Integer
        Dim kq As Integer
        For i = 0 To lst.Count - 1
            If (point.X = lst(i).X) And (point.Y = lst(i).Y) Then

                kq = i
                Exit For
            End If
        Next
        Return kq
    End Function
    Public Function lstPointDetermind2(ByVal lstPoint As List(Of XYZ)) As List(Of List(Of XYZ))
        Dim kqTong As New List(Of List(Of XYZ))
        Dim lstPointold As New List(Of XYZ)
        Dim lstPointoldTong As New List(Of List(Of XYZ))
        For k = 0 To lstPoint.Count - 1
            'Dim kq As List(Of XYZ)
            Dim ListRanger As New List(Of Double)
            Dim lstPointNew1 As New List(Of XYZ)
            Dim lstPointNew2 As New List(Of XYZ)
            Dim lstPointCheckBanDau As New List(Of XYZ)
            Dim lstPointCheck1 As New List(Of XYZ)
            Dim lstPointCheck2 As New List(Of XYZ)
            Dim Check4Point As New List(Of XYZ)
            Dim DiemGoc As XYZ = lstPoint(k)
            'tạo list khoảng cách
            For i = 0 To lstPoint.Count - 1
                Dim a = DetermindRange(DiemGoc, lstPoint(i))
                ListRanger.Add(a)
            Next
            'Sắp xếp lại listpoint theo listkc tăng dần
            For i = 0 To lstPoint.Count - 2
                For j = i + 1 To lstPoint.Count - 1
                    If ListRanger(i) > ListRanger(j) Then
                        Dim tgkc As Double = ListRanger(i)
                        Dim tgXYZ As XYZ = lstPoint(i)
                        ListRanger(i) = ListRanger(j)
                        lstPoint(i) = lstPoint(j)
                        ListRanger(j) = tgkc
                        lstPoint(j) = tgXYZ
                    End If
                Next
            Next
            lstPointCheckBanDau.Add(DiemGoc)
            lstPointCheckBanDau.Add(lstPoint(1))
            Dim Skt As Double = Edge(lstPointCheckBanDau(0), lstPointCheckBanDau(1))
            If Skt = 0 Then

#Region "Ngược chiều"
                lstPointCheck1.Add(DiemGoc)
                lstPointCheck1.Add(lstPoint(1))
                Dim lstPointClone01 As New List(Of XYZ)
                Dim Ranger01 As New List(Of Double)
                Dim Skt1 As Double = 0
                For i = 0 To lstPoint.Count - 1
                    Skt1 = Edge(lstPointCheck1(1), lstPoint(i))
                    Dim a01 As Double = Edge(lstPointCheck1(1), lstPoint(i))
                    If Skt1 > 0 Then
                        Dim kt As Boolean = False

                        If Math.Round(lstPoint(i).X, 1) = Math.Round(lstPointCheck1(0).X, 1) Or Math.Round(lstPoint(i).Y, 1) = Math.Round(lstPointCheck1(0).Y, 1) Or Math.Round(lstPoint(i).X, 1) = Math.Round(lstPointCheck1(1).X, 1) Or Math.Round(lstPoint(i).Y, 1) = Math.Round(lstPointCheck1(1).Y, 1) Then
                            kt = True
                        End If
                        If kt = True Then
                            lstPointClone01.Add(lstPoint(i))
                            Ranger01.Add(a01)
                        End If

                    End If
                Next
                If lstPointClone01.Count >= 2 Then
                        For h = 0 To lstPointClone01.Count - 2
                            For q = h + 1 To lstPointClone01.Count - 1
                                If Ranger01(h) > Ranger01(q) Then
                                    Dim tgkc As Double = Ranger01(h)
                                    Dim tgXYZ As XYZ = lstPointClone01(h)
                                    Ranger01(h) = Ranger01(q)
                                    lstPointClone01(h) = lstPointClone01(q)
                                    Ranger01(q) = tgkc
                                    lstPointClone01(q) = tgXYZ
                                End If
                            Next
                        Next

                        lstPointCheck1.Add(lstPointClone01(0))

                Else
                        lstPointCheck1.Add(lstPointClone01(0))

                End If


                    If lstPointCheck1.Count = 2 Then
                    'MsgBox("1")
                Else
                    Dim lstRanger1 As New List(Of Double)
                    Dim LstPointClone1 As New List(Of XYZ)
                    lstPointNew1.Add(lstPointCheck1(0))
                    lstPointNew1.Add(lstPointCheck1(1))
                    lstPointNew1.Add(lstPointCheck1(2))
                    Dim lstCheck1 As New List(Of XYZ)
                    For i = 0 To lstPoint.Count - 1
                        lstCheck1.Add(lstPoint(i))
                    Next
                    lstCheck1.RemoveAt(RemovePoint(lstPointCheck1(0), lstCheck1))
                    lstCheck1.RemoveAt(RemovePoint(lstPointCheck1(1), lstCheck1))
                    lstCheck1.RemoveAt(RemovePoint(lstPointCheck1(2), lstCheck1))
                    Dim S1 As Double = 0
                    For i = 2 To lstPoint.Count - 1
                        If lstPointNew1.Count = 4 Then
                            Exit For
                        Else

                            For j = 0 To lstCheck1.Count - 1
                                'If lstPoint(j).X = lstCheck(j).X And lstPoint(j).Y = lstCheck(j).Y Then
                                'Else
                                S1 = S1 + Edge(lstPointNew1(lstPointNew1.Count - 1), lstCheck1(j))

                                Dim a = DetermindRange(lstPointNew1(0), lstCheck1(j))
                                ' End If
                                If S1 > 0 Then
                                    S1 = 0
                                Else
                                    If LstPointClone1.Count = 1 Then

                                        Exit For
                                    Else
                                        lstRanger1.Add(a)
                                        LstPointClone1.Add(lstCheck1(j))

                                    End If
                                End If
                            Next
                            For h = 0 To LstPointClone1.Count - 2
                                For q = h + 1 To LstPointClone1.Count - 1
                                    If lstRanger1(h) > lstRanger1(q) Then
                                        Dim tgkc As Double = lstRanger1(h)
                                        Dim tgXYZ As XYZ = LstPointClone1(h)
                                        lstRanger1(h) = lstRanger1(q)
                                        LstPointClone1(h) = LstPointClone1(q)
                                        lstRanger1(q) = tgkc
                                        LstPointClone1(q) = tgXYZ
                                    End If
                                Next
                            Next
                            lstPointNew1.Add(LstPointClone1(0))
                            lstCheck1.RemoveAt(RemovePoint(LstPointClone1(0), lstCheck1))
                            lstRanger1.Clear()
                            LstPointClone1.Clear()
                        End If
                    Next
                    If ChecklistPointTriangle(lstPointNew1) Then
                        kqTong.Add(lstPointNew1)
                    Else

                    End If
                End If

#End Region
#Region "Cùng  chiều"
                Dim Skt2 As Double = 0
                lstPointCheck2.Add(DiemGoc)
                lstPointCheck2.Add(lstPoint(1))
                Dim lstPointClone02 As New List(Of XYZ)
                Dim Ranger02 As New List(Of Double)
                For i = 0 To lstPoint.Count - 1
                    Skt2 = Edge(lstPointCheck2(1), lstPoint(i))
                    Dim a02 As Double = Edge(lstPointCheck2(1), lstPoint(i))
                    If Skt2 < 0 Then
                        Dim kt As Boolean = False
                        If Math.Round(lstPoint(i).X, 1) = Math.Round(lstPointCheck2(0).X, 1) Or Math.Round(lstPoint(i).Y, 1) = Math.Round(lstPointCheck2(0).Y, 1) Or Math.Round(lstPoint(i).X, 1) = Math.Round(lstPointCheck2(1).X, 1) Or Math.Round(lstPoint(i).Y, 1) = Math.Round(lstPointCheck2(1).Y, 1) Then
                            kt = True
                        End If
                        If kt = True Then
                            lstPointClone02.Add(lstPoint(i))
                            Ranger02.Add(a02)
                        End If
                    End If
                Next
                If lstPointCheck2.Count >= 2 Then
                        For h = 0 To lstPointClone02.Count - 2
                            For q = h + 1 To lstPointClone02.Count - 1
                                If Ranger02(h) > Ranger02(q) Then
                                    Dim tgkc As Double = Ranger01(h)
                                    Dim tgXYZ As XYZ = lstPointClone02(h)
                                    Ranger02(h) = Ranger02(q)
                                    lstPointClone02(h) = lstPointClone02(q)
                                    Ranger02(q) = tgkc
                                    lstPointClone02(q) = tgXYZ
                                End If
                            Next
                        Next
                        lstPointCheck2.Add(lstPointClone02(0))

                    Else
                        lstPointCheck2.Add(lstPointClone02(0))

                    End If

                    If lstPointCheck2.Count = 2 Then
                    'MsgBox("2")
                Else
                    Dim lstRanger2 As New List(Of Double)
                    Dim LstPointClone2 As New List(Of XYZ)
                    lstPointNew2.Add(lstPointCheck2(0))
                    lstPointNew2.Add(lstPointCheck2(1))
                    lstPointNew2.Add(lstPointCheck2(2))
                    Dim lstCheck2 As New List(Of XYZ)
                    For i = 0 To lstPoint.Count - 1
                        lstCheck2.Add(lstPoint(i))
                    Next
                    lstCheck2.RemoveAt(RemovePoint(lstPointCheck2(0), lstCheck2))
                    lstCheck2.RemoveAt(RemovePoint(lstPointCheck2(1), lstCheck2))
                    lstCheck2.RemoveAt(RemovePoint(lstPointCheck2(2), lstCheck2))

                    Dim S2 As Double = 0
                    For i = 2 To lstPoint.Count - 1
                        If lstPointNew2.Count = 4 Then
                            Exit For
                        Else

                            For j = 0 To lstCheck2.Count - 1
                                'If lstPoint(j).X = lstCheck(j).X And lstPoint(j).Y = lstCheck(j).Y Then
                                'Else
                                S2 = S2 + Edge(lstPointNew2(lstPointNew2.Count - 1), lstCheck2(j))
                                Dim a = DetermindRange(lstPointNew2(0), lstCheck2(j))
                                ' End If
                                If S2 < 0 Then
                                    S2 = 0
                                Else
                                    If LstPointClone2.Count = 1 Then

                                        Exit For
                                    Else
                                        lstRanger2.Add(a)
                                        LstPointClone2.Add(lstCheck2(j))

                                    End If
                                End If
                            Next
                            For h = 0 To LstPointClone2.Count - 2
                                For q = h + 1 To LstPointClone2.Count - 1
                                    If lstRanger2(h) > lstRanger2(q) Then
                                        Dim tgkc As Double = lstRanger2(h)
                                        Dim tgXYZ As XYZ = LstPointClone2(h)
                                        lstRanger2(h) = lstRanger2(q)
                                        LstPointClone2(h) = LstPointClone2(q)
                                        lstRanger2(q) = tgkc
                                        LstPointClone2(q) = tgXYZ
                                    End If
                                Next
                            Next

                            lstPointNew2.Add(LstPointClone2(0))
                            lstCheck2.RemoveAt(RemovePoint(LstPointClone2(0), lstCheck2))
                            lstRanger2.Clear()
                            LstPointClone2.Clear()
                        End If
                    Next
                    If ChecklistPointTriangle(lstPointNew2) Then
                        kqTong.Add(lstPointNew2)
                    Else

                    End If
                End If
#End Region
            Else
#Region "Ngược chiều"

                Dim Skt1 As Double = 0
                lstPointCheck1.Add(DiemGoc)
                For i = 0 To lstPoint.Count - 1
                    Skt1 = Edge(lstPointCheck1(0), lstPoint(i))
                    If Skt1 < 0 Then
                        lstPointCheck1.Add(lstPoint(i))
                        If lstPointCheck1.Count = 2 Then
                            Exit For
                        End If
                    End If
                Next
                If lstPointCheck1.Count = 1 Then
                    'MsgBox("3")
                Else
                    Dim lstRanger1 As New List(Of Double)
                    Dim LstPointClone1 As New List(Of XYZ)
                    lstPointNew1.Add(lstPointCheck1(0))
                    lstPointNew1.Add(lstPointCheck1(1))
                    Dim lstCheck1 As New List(Of XYZ)
                    For i = 0 To lstPoint.Count - 1
                        lstCheck1.Add(lstPoint(i))
                    Next
                    lstCheck1.RemoveAt(RemovePoint(lstPointCheck1(0), lstCheck1))
                    lstCheck1.RemoveAt(RemovePoint(lstPointCheck1(1), lstCheck1))
                    Dim S1 As Double = 0
                    For i = 2 To lstPoint.Count - 1
                        If lstPointNew1.Count = 4 Then
                            Exit For
                        Else

                            For j = 0 To lstCheck1.Count - 1
                                'If lstPoint(j).X = lstCheck(j).X And lstPoint(j).Y = lstCheck(j).Y Then
                                'Else
                                S1 = Edge(lstPointNew1(lstPointNew1.Count - 1), lstCheck1(j))
                                Dim a = DetermindRange(lstPointNew1(lstPointNew1.Count - 1), lstCheck1(j))
                                ' End If
                                If S1 > 0 Then
                                    S1 = 0
                                Else
                                    If lstPointNew1.Count = 3 Then
                                        If Edge(lstPointNew1(lstPointNew1.Count - 1), lstCheck1(j)) = 0 Then

                                        Else
                                            lstRanger1.Add(a)
                                            LstPointClone1.Add(lstCheck1(j))
                                        End If
                                    Else
                                        'If LstPointClone1.Count = 2 Then

                                        '    Exit For
                                        'Else
                                        lstRanger1.Add(a)
                                            LstPointClone1.Add(lstCheck1(j))

                                        'End If
                                    End If
                                End If
                            Next
                            For h = 0 To LstPointClone1.Count - 2
                                For q = h + 1 To LstPointClone1.Count - 1
                                    If lstRanger1(h) > lstRanger1(q) Then
                                        Dim tgkc As Double = lstRanger1(h)
                                        Dim tgXYZ As XYZ = LstPointClone1(h)
                                        lstRanger1(h) = lstRanger1(q)
                                        LstPointClone1(h) = LstPointClone1(q)
                                        lstRanger1(q) = tgkc
                                        LstPointClone1(q) = tgXYZ
                                    End If
                                Next
                            Next
                            Dim lstPointFinal As New List(Of XYZ)
                            For y = 0 To LstPointClone1.Count - 1
                                If Edge(lstPointNew1(0), LstPointClone1(y)) >= 0 Then
                                    lstPointFinal.Add(LstPointClone1(y))
                                End If
                            Next
                            If lstPointFinal.Count = 0 Then
                                lstPointNew1.Add(LstPointClone1(0))
                                lstCheck1.RemoveAt(RemovePoint(LstPointClone1(0), lstCheck1))
                                Dim lstPointClone11 As New List(Of XYZ)
                                Dim lstRanger11 As New List(Of Double)
                                For t = 0 To lstCheck1.Count - 1
                                    Dim S3 As Double
                                    S3 = DetermindRange(lstPointNew1(lstPointNew1.Count - 1), lstCheck1(t))
                                    Dim b = DetermindRange(lstPointNew1(lstPointNew1.Count - 1), lstCheck1(t))
                                    If S3 > 0 Then
                                        If lstPointNew1.Count = 3 Then
                                            If Edge(lstPointNew1(lstPointNew1.Count - 1), lstCheck1(t)) = 0 Then

                                            Else
                                                lstRanger11.Add(b)
                                                lstPointClone11.Add(lstCheck1(t))
                                            End If
                                        Else
                                            'If LstPointClone1.Count = 2 Then

                                            '    Exit For
                                            'Else
                                            lstRanger11.Add(b)
                                            lstPointClone11.Add(lstCheck1(t))

                                            'End If
                                        End If
                                    Else
                                        S3 = 0
                                    End If
                                Next
                                For h = 0 To lstPointClone11.Count - 2
                                    For q = h + 1 To LstPointClone1.Count - 1
                                        If lstRanger11(h) > lstRanger11(q) Then
                                            Dim tgkc As Double = lstRanger11(h)
                                            Dim tgXYZ As XYZ = lstPointClone11(h)
                                            lstRanger11(h) = lstRanger11(q)
                                            lstPointClone11(h) = lstPointClone11(q)
                                            lstRanger11(q) = tgkc
                                            lstPointClone11(q) = tgXYZ
                                        End If
                                    Next
                                Next
                                lstPointNew1.Add(lstPointClone11(0))
                                lstRanger11.Clear()
                                lstPointClone11.Clear()
                            Else
                                lstPointNew1.Add(lstPointFinal(0))
                                lstCheck1.RemoveAt(RemovePoint(lstPointFinal(0), lstCheck1))
                            End If

                            lstRanger1.Clear()
                            LstPointClone1.Clear()
                        End If
                    Next
                    If ChecklistPointTriangle(lstPointNew1) Then
                        kqTong.Add(lstPointNew1)
                    Else

                    End If
                End If

#End Region
#Region "Cùng  chiều"
                Dim Skt2 As Double = 0
                lstPointCheck2.Add(DiemGoc)
                For i = 0 To lstPoint.Count - 1
                    Skt2 = Edge(lstPointCheck2(0), lstPoint(i))
                    If Skt2 > 0 Then
                        lstPointCheck2.Add(lstPoint(i))
                        If lstPointCheck2.Count = 2 Then
                            Exit For


                        End If
                    End If
                Next
                If lstPointCheck2.Count = 1 Then
                    'MsgBox("4")
                Else
                    Dim lstRanger2 As New List(Of Double)
                    Dim LstPointClone2 As New List(Of XYZ)
                    lstPointNew2.Add(lstPointCheck2(0))
                    lstPointNew2.Add(lstPointCheck2(1))
                    Dim lstCheck2 As New List(Of XYZ)
                    For i = 0 To lstPoint.Count - 1
                        lstCheck2.Add(lstPoint(i))
                    Next
                    lstCheck2.RemoveAt(RemovePoint(lstPointCheck2(0), lstCheck2))
                    lstCheck2.RemoveAt(RemovePoint(lstPointCheck2(1), lstCheck2))
                    Dim S2 As Double = 0
                    For i = 2 To lstPoint.Count - 1
                        If lstPointNew2.Count = 4 Then
                            Exit For
                        Else

                            For j = 0 To lstCheck2.Count - 1
                                'If lstPoint(j).X = lstCheck(j).X And lstPoint(j).Y = lstCheck(j).Y Then
                                'Else
                                S2 = Edge(lstPointNew2(lstPointNew2.Count - 1), lstCheck2(j))
                                Dim a = DetermindRange(lstPointNew2(lstPointNew2.Count - 1), lstCheck2(j))
                                ' End If
                                If S2 < 0 Then
                                    S2 = 0
                                Else
                                    If lstPointNew2.Count = 3 Then
                                        If Edge(lstPointNew2(lstPointNew2.Count - 1), lstCheck2(j)) = 0 Then

                                        Else
                                            lstRanger2.Add(a)
                                            LstPointClone2.Add(lstCheck2(j))
                                        End If
                                    Else
                                        'If LstPointClone2.Count = 2 Then

                                        ' Exit For
                                        'Else

                                        lstRanger2.Add(a)
                                        LstPointClone2.Add(lstCheck2(j))

                                        'End If
                                    End If

                                End If
                            Next
                            For h = 0 To LstPointClone2.Count - 2
                                For q = h + 1 To LstPointClone2.Count - 1
                                    If lstRanger2(h) > lstRanger2(q) Then
                                        Dim tgkc As Double = lstRanger2(h)
                                        Dim tgXYZ As XYZ = LstPointClone2(h)
                                        lstRanger2(h) = lstRanger2(q)
                                        LstPointClone2(h) = LstPointClone2(q)
                                        lstRanger2(q) = tgkc
                                        LstPointClone2(q) = tgXYZ
                                    End If
                                Next
                            Next
                            Dim lstPointFinal As New List(Of XYZ)
                            For y = 0 To LstPointClone2.Count - 1
                                If Edge(lstPointNew2(0), LstPointClone2(y)) <= 0 Then
                                    lstPointFinal.Add(LstPointClone2(y))
                                End If
                            Next
                            If lstPointFinal.Count = 0 Then
                                lstPointNew2.Add(LstPointClone2(0))
                                lstCheck2.RemoveAt(RemovePoint(LstPointClone2(0), lstCheck2))
                                Dim lstPointClone22 As New List(Of XYZ)
                                Dim lstRanger22 As New List(Of Double)
                                For t = 0 To lstCheck2.Count - 1
                                    Dim S3 As Double
                                    S3 = DetermindRange(lstPointNew2(lstPointNew2.Count - 1), lstCheck2(t))
                                    Dim b = DetermindRange(lstPointNew2(lstPointNew2.Count - 1), lstCheck2(t))
                                    If S3 > 0 Then
                                        If lstPointNew2.Count = 3 Then
                                            If Edge(lstPointNew2(lstPointNew2.Count - 1), lstCheck2(t)) = 0 Then

                                            Else
                                                lstRanger22.Add(b)
                                                lstPointClone22.Add(lstCheck2(t))
                                            End If
                                        Else
                                            'If LstPointClone1.Count = 2 Then

                                            '    Exit For
                                            'Else
                                            lstRanger22.Add(b)
                                            lstPointClone22.Add(lstCheck2(t))

                                            'End If
                                        End If
                                    Else
                                        S3 = 0
                                    End If
                                Next
                                For h = 0 To lstPointClone22.Count - 2
                                    For q = h + 1 To LstPointClone2.Count - 1
                                        If lstRanger22(h) > lstRanger22(q) Then
                                            Dim tgkc As Double = lstRanger22(h)
                                            Dim tgXYZ As XYZ = lstPointClone22(h)
                                            lstRanger22(h) = lstRanger22(q)
                                            lstPointClone22(h) = lstPointClone22(q)
                                            lstRanger22(q) = tgkc
                                            lstPointClone22(q) = tgXYZ
                                        End If
                                    Next
                                Next
                                lstPointNew2.Add(lstPointClone22(0))
                                lstRanger22.Clear()
                                lstPointClone22.Clear()
                            Else
                                lstPointNew2.Add(lstPointFinal(0))
                                lstCheck2.RemoveAt(RemovePoint(lstPointFinal(0), lstCheck2))
                            End If

                            lstRanger2.Clear()
                            LstPointClone2.Clear()
                        End If
                    Next
                    If ChecklistPointTriangle(lstPointNew2) Then

                        kqTong.Add(lstPointNew2)
                    Else

                    End If
                End If

#End Region
            End If


            'lstPointold.Add(lstPoint(i))
            'lstPointoldTong.Add(lstPointold)
            ''lstCheckPoint.Add(Check4Point)
            'kq = lstPointNew
            'kqTong.Add(kq)

            'DeletePoint(lstPoint, DiemGoc)
        Next


        Return kqTong
    End Function
    Public Function ChecklistPointTriangle(ByVal lstPoint As List(Of XYZ)) As Boolean
        Dim kq As Boolean = True



        Dim Canh1 As Double = Math.Round(DetermindRange(lstPoint(0), lstPoint(1)), 1)
        Dim Canh2 As Double = Math.Round(DetermindRange(lstPoint(0), lstPoint(3)), 1)
        Dim Canh3 As Double = Math.Round(DetermindRange(lstPoint(1), lstPoint(2)), 1)
        Dim Canh4 As Double = Math.Round(DetermindRange(lstPoint(2), lstPoint(3)), 1)
        Dim CanhHuyen As Double = Math.Round(DetermindRange(lstPoint(1), lstPoint(3)), 1)
        Dim c = CanhHuyen
        Dim a As Double = Canh1 * Canh1
        Dim b As Double = Canh2 * Canh2
        Dim e As Double = Canh3 * Canh3
        Dim f As Double = Canh4 * Canh4
        Dim T1 As Double = Math.Sqrt(a + b)
        Dim T2 As Double = Math.Sqrt(e + f)
        If T1 > c Then
            If ((T1 - c) <= 1 And (T1 - c) >= 0) And ((T2 - c) <= 1 And (T2 - c) >= 0) Then

            Else
                kq = False
            End If
        Else
            If ((c - T1) <= 1 And (c - T1) >= 0) And ((c - T2) <= 1 And (c - T2) >= 0) Then

            Else
                kq = False
            End If
        End If


        Return kq
    End Function
    Public Function DetermindRange(ByVal Diem1 As XYZ, ByVal Diem2 As XYZ) As Double
        Dim kq As Double = Math.Sqrt(((Diem1.X - Diem2.X) * (Diem1.X - Diem2.X)) + ((Diem1.Y - Diem2.Y) * (Diem1.Y - Diem2.Y)))
        Return kq
    End Function
    Public Function DeletePoint(ByVal lstPoint As List(Of XYZ), ByVal Point As XYZ) As List(Of XYZ)
        Dim kq As New List(Of XYZ)
        Dim index As Integer = 0
        For i = 0 To lstPoint.Count - 1
            If Point.X = lstPoint(i).X And Point.Y = lstPoint(i).Y Then
                index = i
            End If
        Next
        'For i = index To lstPoint.Count - 2
        '    lstPoint(i) = lstPoint(i + 1)
        'Next
        lstPoint.RemoveAt(index)
        kq = lstPoint
        Return kq
    End Function
    Public Function Findintersectline(ByVal lstab As List(Of List(Of Double)), ByVal z As Double, ByVal ListDoanThang As List(Of List(Of XYZ))) As List(Of XYZ)
        Dim lstPointintersect As New List(Of XYZ)
        If lstab.Count >= 1 Then

            For i = 0 To lstab.Count - 2
                For j = i + 1 To lstab.Count - 1
                    Dim abDoanThang1 As List(Of Double) = lstab(i)
                    Dim abDoanThang2 As List(Of Double) = lstab(j)
                    Dim DoanThang1 As List(Of XYZ) = ListDoanThang(i)
                    Dim DoanThang2 As List(Of XYZ) = ListDoanThang(j)
                    Dim a1 As Double = abDoanThang1(0)
                    Dim b1 As Double = abDoanThang1(1)
                    Dim a2 As Double = abDoanThang2(0)
                    Dim b2 As Double = abDoanThang2(1)
                    Dim x, y As Double
                    If (a1 = 0 And b1 = 0) Then
                        x = Math.Round(DoanThang1(0).X, 3)
                        y = a2 * x + b2
                    ElseIf (a2 = 0 And b2 = 0) Then
                        x = Math.Round(DoanThang2(0).X, 3)
                        y = a1 * x + b1
                    Else
                        x = Math.Round(((b2 - b1) / (a1 - a2)), 3)
                        y = Math.Round((a1 * x + b1), 3)
                    End If
                    Dim x0 As Double = Math.Round(DoanThang1(0).X, 3)
                    Dim x1 As Double = Math.Round(DoanThang1(1).X, 3)
                    Dim X2 As Double = Math.Round(DoanThang2(0).X, 3)
                    Dim x3 As Double = Math.Round(DoanThang2(1).X, 3)
                    If x >= Math.Min(x0, x1) And x <= Math.Max(x0, x1) And x >= Math.Min(X2, x3) And x <= Math.Max(X2, x3) Then
                        Dim Pointintersect As XYZ = New XYZ(x, y, z)
                        lstPointintersect.Add(Pointintersect)
                    End If
                Next
            Next
        End If
        Return lstPointintersect
    End Function
    Private Sub CreateExtrusion(document As Autodesk.Revit.DB.Document, sketchPlane As SketchPlane, lstPoint As List(Of XYZ), z1 As Double, z2 As Double)
        Try
            'If Check4Point(lstPoint, lstCheckPoint) Then
            Dim rectExtrusion As Extrusion = Nothing
                lstCheckPoint.Add(lstPoint)
            'Dim S As Double = 0
            'For i = 0 To lstPoint.Count - 2
            '    For j = i + 1 To lstPoint.Count - 1
            '        S = S + Edge(lstPoint(i), lstPoint(j))
            '        If S < 0 Then

            '        Else
            '            Dim tg As XYZ = lstPoint(i)
            '            lstPoint(i) = lstPoint(j)
            '            lstPoint(j) = tg
            '            Exit For
            '        End If
            '    Next
            'Next
            If True = document.IsFamilyDocument Then

                Dim curveArrArray As New CurveArrArray()
                Dim curveArray1 As New CurveArray()
                For i = 0 To lstPoint.Count - 1
                    Dim line As Line
                    If i = lstPoint.Count - 1 Then
                        Dim p0 As XYZ = lstPoint(i)
                        Dim p1 As XYZ = lstPoint(0)
                        line = Line.CreateBound(p0, p1)
                    Else
                        Dim p0 As XYZ = lstPoint(i)
                        Dim p1 As XYZ = lstPoint(i + 1)
                        line = Line.CreateBound(p0, p1)
                    End If
                    curveArray1.Append(line)
                Next
                curveArrArray.Append(curveArray1)
                ' create solid rectangular extrusion
                Dim random As New Random
                Dim Ex = random.Next(z1, z2)
                rectExtrusion = document.FamilyCreate.NewExtrusion(True, curveArrArray, sketchPlane, Ex)

                '    If rectExtrusion IsNot Nothing Then
                '        ' move extrusion to proper place
                '        Dim transPoint1 As New XYZ(-16, 0, 0)
                '        ElementTransformUtils.MoveElement(document, rectExtrusion.Id, transPoint1)
                '    Else
                '        Throw New Exception("Create new Extrusion failed.")
                '    End If
            Else
                Throw New Exception("Please open a Family document before invoking this command.")
            End If

            ' End If
        Catch ex As Exception
            lstCheckPoint.Clear()
            Dim [error] As String = ex.Message
        End Try


    End Sub
    Public Function Check4Point(ByVal listPoint As List(Of XYZ), ByVal ListCheck As List(Of List(Of XYZ))) As Boolean
        Dim kq As Boolean = True
        Dim kq1 As Boolean = True
        Dim kq2 As Boolean = True
        Dim kq3 As Boolean = True
        Dim kq4 As Boolean = True
        Dim kt As Double = 0
        If ListCheck Is Nothing Then
        Else
            For i = 0 To ListCheck.Count - 1
                Dim Check As List(Of XYZ) = ListCheck(i)
                For j = 0 To Check.Count - 1
                    If listPoint(0).X = Check(j).X And listPoint(0).Y = Check(j).Y Then
                        kq1 = False
                        kt = kt + 1
                    End If
                Next
                For j = 0 To Check.Count - 1
                    If listPoint(1).X = Check(j).X And listPoint(1).Y = Check(j).Y Then
                        kq2 = False
                        kt = kt + 1
                    End If
                Next
                For j = 0 To Check.Count - 1
                    If listPoint(2).X = Check(j).X And listPoint(2).Y = Check(j).Y Then
                        kq3 = False
                        kt = kt + 1
                    End If
                Next
                For j = 0 To Check.Count - 1
                    If listPoint(3).X = Check(j).X And listPoint(3).Y = Check(j).Y Then
                        kq4 = False
                        kt = kt + 1
                    End If
                Next
            Next
            If kt >= 4 Then
                kq = False
            End If
        End If

        Return kq
    End Function
    Public Function SortListPoint(ByVal lstpoint As List(Of XYZ)) As List(Of XYZ)
        Dim kq As New List(Of XYZ)
        Dim S As Double = 0
        For i = 0 To lstpoint.Count - 2
            For j = i + 1 To lstpoint.Count - 1
                S = S + Edge(lstpoint(i), lstpoint(j))
                If S < 0 Then

                Else
                    Dim tg As XYZ = lstpoint(i)
                    lstpoint(i) = lstpoint(j)
                    lstpoint(j) = tg
                    Exit For
                End If
            Next
        Next
        For i = 0 To lstpoint.Count - 1
            kq.Add(lstpoint(i))
        Next
        Return kq
    End Function
    Public Function BoolLess(a As XYZ, b As XYZ, Center As XYZ) As Boolean
        Dim kq As Boolean
        Dim det As Double = (a.X - Center.X) * (b.Y - Center.Y) - (b.X - Center.X) * (a.Y - Center.Y)
        If det < 0 Then
            kq = True
        Else
            kq = False
        End If
        Return kq
    End Function
    Public Function Edge(diem1 As XYZ, diem2 As XYZ) As Double
        Dim kq As Double = (Math.Round((diem2.X), 3) - Math.Round((diem1.X), 3)) * (Math.Round((diem2.Y), 3) + Math.Round((diem1.Y), 3))
        Return kq
    End Function
    Public Sub GetInfo_BoundarySegment(room As Autodesk.Revit.DB.Architecture.Room)
        Dim segments As IList(Of IList(Of Autodesk.Revit.DB.BoundarySegment)) = room.GetBoundarySegments(New SpatialElementBoundaryOptions())

        If segments IsNot Nothing Then
            'the room may not be bound
            Dim message As String = "BoundarySegment"
            For Each segmentList As IList(Of Autodesk.Revit.DB.BoundarySegment) In segments
                For Each boundarySegment As Autodesk.Revit.DB.BoundarySegment In segmentList

                    ' Get curve start point
                    message += ((vbLf & "Curve start point: (" + Convert.ToString(boundarySegment.GetCurve().GetEndPoint(0).X) & ",") + Convert.ToString(boundarySegment.GetCurve().GetEndPoint(0).Y) & ",") + Convert.ToString(boundarySegment.GetCurve().GetEndPoint(0).Z) & ")"
                    ' Get curve end point
                    message += ((";" & vbLf & "Curve end point: (" + Convert.ToString(boundarySegment.GetCurve().GetEndPoint(1).X) & ",") + Convert.ToString(boundarySegment.GetCurve().GetEndPoint(1).Y) & ",") + Convert.ToString(boundarySegment.GetCurve().GetEndPoint(1).Z) & ")"
                    ' Get document path name
                    'message += ";" & vbLf & "Document path name: " + room.Document.PathName
                    ' Get boundary segment element name
                    If boundarySegment.ElementId IsNot ElementId.InvalidElementId Then
                        message += ";" & vbLf & "Element name: " + room.Document.GetElement(boundarySegment.ElementId).Name
                    End If

                Next
            Next
            TaskDialog.Show("Revit", message)
        End If
    End Sub
    Private Sub CreateRooms(ByVal document As Document)
        Dim phases As PhaseArray = document.Phases
        Dim createRoomsInPhase As Phase = phases.Item(phases.Size - 1)
        Dim collector As FilteredElementCollector = New FilteredElementCollector(document)
        collector.OfClass(GetType(Level))

        Using tran As Transaction = New Transaction(document)
            Dim x As Integer = 1
            tran.Start("tran1")

            For Each level As Level In collector
                Dim topology As PlanTopology = document.PlanTopology(level, createRoomsInPhase)
                Dim circuitSet As PlanCircuitSet = topology.Circuits

                For Each circuit As PlanCircuit In circuitSet

                    If Not circuit.IsRoomLocated Then
                        Dim room As Room = document.Create.NewRoom(Nothing, circuit)
                        room.Name = "Room name: " & x
                        x += 1
                        GetInfo_BoundarySegment(room)
                    End If
                Next
            Next

            tran.Commit()
        End Using
    End Sub
    Private Sub CreateRoom(ByVal doc As Document, ByVal TenLevel As String)
        Dim roomLocation As UV = New UV(0, 0)
        Dim lev As Level = FindLevel(doc, TenLevel)
        Try
            Using tr As New Transaction(doc, "CreateRoom")
                tr.Start()
                Dim room As Room = doc.Create.NewRoom(lev, roomLocation)
                If room Is Nothing Then
                    Throw New Exception("Create a new room failed.")
                End If
                tr.Commit()
            End Using
        Catch ex As Exception
            Dim [error] As String = ex.Message
        End Try
    End Sub
    Private Sub CreateRoom(ByVal uidoc As UIDocument)
        'Dim filledRegions As List(Of Element) = New List(Of Element)()
        'If Utility.GetSelectedElementsOrAll(filledRegions, uidoc, GetType(FilledRegion)) Then
        '    Dim n As Integer = filledRegions.Count
        '    Dim results As String() = New String(n - 1) {}
        '    Dim i As Integer = 0

        '    For Each region As FilledRegion In filledRegions.Cast(Of FilledRegion)()
        '        Dim desc As String = Util.ElementDescription(region)
        '        Dim corners As List(Of XYZ) = GetBoundaryCorners(region)
        '        Dim result As String = If((corners Is Nothing), "failed", String.Join(", ", corners.ConvertAll(Of String)(Function(p) Util.PointString(p)).ToArray()))
        '        results(Math.Min(System.Threading.Interlocked.Increment(i), i - 1)) = String.Format("{0}: {1}", desc, result)
        '    Next

        '    Dim s As String = String.Format("Retrieving corners for {0} filled region{1}{2}", n, Util.PluralSuffix(n), Util.DotOrColon(n))
        '    Dim t As String = String.Join(vbCrLf, results)
        '    Util.InfoMsg(s, t)
        'End If

        'Try
        '    Using tr As New Transaction(doc, "CreateRoom")
        '        tr.Start()

        '        tr.Commit()
        '    End Using
        'Catch ex As Exception
        '    Dim [error] As String = ex.Message
        'End Try
    End Sub
    Private Function GetBoundaryCorners(ByVal region As FilledRegion) As List(Of XYZ)
        Dim result As List(Of XYZ) = New List(Of XYZ)()
        Dim id As ElementId = New ElementId(region.Id.IntegerValue - 1)
        Dim sketch As Sketch = TryCast(region.Document.GetElement(id), Sketch)

        If sketch IsNot Nothing Then
            Dim curves As CurveArray = sketch.Profile.Item(0)

            If curves IsNot Nothing Then

                For Each cur As Curve In curves
                    Dim corner As XYZ = cur.GetEndPoint(0)
                    result.Add(corner)
                Next
            End If
        End If

        Return result
    End Function
    Public Function Lineequations(ByVal ListDoanThangPoint As List(Of List(Of XYZ))) As List(Of List(Of Double))
        Dim lst As New List(Of List(Of Double))


        For i = 0 To ListDoanThangPoint.Count - 1
            Dim listTowPoint As List(Of XYZ) = ListDoanThangPoint(i)
#Region "Tìm giao điểm"
            Dim stpoint As XYZ = listTowPoint(0)
            Dim EndPoint As XYZ = listTowPoint(1)
            Dim VTCP As XYZ = New XYZ(EndPoint.X - stpoint.X, EndPoint.Y - stpoint.Y, 0) 'Vectp chi phuong
            Dim VTPT As XYZ = New XYZ(-VTCP.X, VTCP.Y, VTCP.Z) 'Vecto phap tuyen 
            Dim n As XYZ = New XYZ(VTPT.X, VTPT.Y, 0)
            Dim a, b As New Double 'pt có dạng y=ax+b
            ' giải phương trình đi qua 2 điểm tìm a,b
            Dim y1 As Double = stpoint.Y
            Dim x1 As Double = stpoint.X
            Dim y2 As Double = EndPoint.Y
            Dim x2 As Double = EndPoint.X
            '2 pt có dạng y1=x1a+b,y2=x2a+b
            Dim lstd As List(Of Double) = SolCamer(x1, 1, y1, x2, 1, y2)
            a = lstd(0)
            b = lstd(1)
            Dim lstab As New List(Of Double)
            lstab.Add(a)
            lstab.Add(b)
            lst.Add(lstab)
#End Region
        Next


        'pt dường thẳng tạo đc có dạng y=ax+b
        Return lst
    End Function

    Public Function SolCamer(a1 As Double, b1 As Double, c1 As Double, a2 As Double, b2 As Double, c2 As Double) As List(Of Double)
        Dim kq As New List(Of Double)
        'pt có dạng a1x+b1y=c1,a2x+b2y=c2
        Dim D, DX, DY, x, y As Double
        D = Math.Round(a1, 1) * Math.Round(b2, 1) - Math.Round(a2, 1) * Math.Round(b1, 1)
        DX = c1 * b2 - c2 * b1
        DY = a1 * c2 - a2 * c1
        If (D = 0) Then
            If (DX + DY = 0) Then
                'TaskDialog.Show("Lỗi", "Hệ phương trình vô số nghiệm")
            Else
                'TaskDialog.Show("Lỗi", "Hệ phương trình vô nghiệm")
                If Math.Round(a1, 1) = Math.Round(a2, 1) Then
                    x = 0 'a 'Phương trình có dạng x = a1 ,
                    y = 0 'b
                ElseIf Math.Round(c1, 1) = Math.Round(c2, 1) Then
                    x = 0  'a
                    y = c1 'b 'phương trình có dạng y = c1 ,b = c1
                End If
                kq.Add(x)
                kq.Add(y)
            End If
        Else
            x = DX / D
            y = DY / D
            kq.Add(x)
            kq.Add(y)
        End If
        Return kq
    End Function
End Class

