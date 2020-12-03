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
Module Module1
    Public LstModelLines As New List(Of ModelLine)
    Public curveArr As New CurveArrArray()
    Public ModelCurveArr As New ModelCurveArrArray
    Public sketplane As SketchPlane
    Public lstCheckPoint As New List(Of List(Of XYZ))
End Module
