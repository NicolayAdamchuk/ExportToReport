Imports Autodesk.Revit.DB
Imports System.Collections

Public Class Sch

    Public schudelerId As Autodesk.Revit.DB.ElementId

    Public schudelerName As String

End Class


Class IncorrectKeyInElements


    Public Name As String
    Public Id_int As Integer
    Public Keynote As String
    Public Sub New(Keynote As String, Id_int As Integer, Name As String)

        Me.Name = Name
        Me.Id_int = Id_int
        Me.Keynote = Keynote

    End Sub
End Class

Public Class TableKeyNotes
    Public param1 As ParamColumn
    Public param2 As ParamColumn
    Public param3 As ParamColumn
    Public sort1 As ParamColumn
    Public sort2 As String
    Public category As BuiltInCategory
    Public Keynote As String
    Public Name As String
    Public Unit1 As String
    Public Unit2 As String


    Public Sub New(Keynote As String, Name As String, Unit1 As String, Unit2 As String, category As BuiltInCategory, param1 As ParamColumn, param2 As ParamColumn, param3 As ParamColumn, sort1 As ParamColumn)

        Me.Keynote = Keynote
        Me.param1 = param1
        Me.param2 = param2
        Me.param3 = param3
        Me.category = category
        Me.Name = Name
        Me.Unit1 = Unit1
        Me.Unit2 = Unit2
        Me.sort1 = sort1

    End Sub

End Class


Public Class ParamColumn

    Public param As BuiltInParameter
    Public type As TypeParam
    Public Sub New(param As BuiltInParameter, typeParam As TypeParam)
        Me.param = param
        Me.type = typeParam
    End Sub



End Class
Public Enum TypeParam
    Instance
    Symbol
End Enum

Public Class DataTableExcel

    Public kode_key As String
    Public kode_name As String
    Public Unit1 As String
    Public Unit2 As String
    Public sort1 As String
    Public sort2 As String
    Public mark As String
    Public kol As Integer
    Public param1 As Double
    Public param2 As Double
    Public param3 As Double

    Public Sub New(kode_key As String, kode_name As String, Unit1 As String, Unit2 As String, sort1 As String, sort2 As String, mark As String, kol As Integer, param1 As Double, param2 As Double, param3 As Double)
        Me.kode_key = kode_key
        Me.kode_name = kode_name
        Me.Unit1 = Unit1
        Me.Unit2 = Unit2
        If sort1 Is Nothing Then
            sort1 = ""
        End If
        If sort2 Is Nothing Then
            sort2 = ""
        End If
        Me.sort1 = sort1
        Me.sort2 = sort2
        If mark Is Nothing Then
            mark = ""
        End If
        Me.mark = mark
        Me.kol = kol
        Me.param1 = param1 * 0.3048
        Me.param2 = param2 * 0.3048
        Me.param3 = param3 * 0.3048
    End Sub

    Public ReadOnly Property param4() As Double
        Get
            param4 = kol * param1 * param2 * param3
        End Get
    End Property

    Public ReadOnly Property param5() As Double
        Get
            param5 = 2 * (param1 + param2)
        End Get
    End Property

    Public ReadOnly Property param6() As Double
        Get
            param6 = kol * param5 * param3
        End Get
    End Property


End Class
