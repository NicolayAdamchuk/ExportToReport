' 
' (C) Copyright 2003-2019 by Autodesk, Inc.
' 
' Permission to use, copy, modify, and distribute this software in
' object code form for any purpose and without fee is hereby granted,
' provided that the above copyright notice appears in all copies and
' that both that copyright notice and the limited warranty and
' restricted rights notice below appear in all supporting
' documentation.
'
' AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
' AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
' MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
' DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
' UNINTERRUPTED OR ERROR FREE.
' 
' Use, duplication, or disclosure by the U.S. Government is subject to
' restrictions set forth in FAR 52.227-19 (Commercial Computer
' Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
' (Rights in Technical Data and Computer Software), as applicable.
'
Imports MsExcel = Microsoft.Office.Interop.Excel
Imports Autodesk.Revit
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.DB
Imports System.IO



<Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)> _
<Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)>
<Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)>
Public Class Command
    Implements IExternalCommand

    Private excel As MsExcel.Application
    Private m_tempFileName As String
    Private m_application As Autodesk.Revit.UI.UIApplication
    Private m_AllElementsWithKey = New List(Of Element)
    Private m_AllElementsWithOutKey = New List(Of Element)
    Private m_newFileName As String
    Private newWorkbook As MsExcel.Workbook
    Private oldWorkbook As MsExcel.Workbook
    Private keyTable As KeynoteTable
    Private BaseKeyTable As List(Of TableKeyNotes)
    Private m_IncorrectElements As List(Of IncorrectKeyInElements)
    Private m_DataTableExcel As List(Of DataTableExcel)
    Private m_DataTableExcelShort As List(Of DataTableExcel)
    Private m_param_project As String
    Private m_def_param_project As Definition


    ''' <summary>
    ''' Implement this method as an external command for Revit.
    ''' </summary>
    ''' <param name="commandData">An object that is passed to the external application 
    ''' which contains data related to the command, 
    ''' such as the application object and active view.</param>
    ''' <param name="message">A message that can be set by the external application 
    ''' which will be displayed if a failure or cancellation is returned by 
    ''' the external command.</param>
    ''' <param name="elements">A set of elements to which the external application 
    ''' can add elements that are to be highlighted in case of failure or cancellation.</param>
    ''' <returns>Return the status of the external command. 
    ''' A result of Succeeded means that the API external method functioned as expected. 
    ''' Cancelled can be used to signify that the user cancelled the external operation 
    ''' at some point. Failure should be returned if the application is unable to proceed with 
    ''' the operation.</returns>
    Public Function Execute(ByVal commandData As ExternalCommandData, ByRef message As String, ByVal elements As Autodesk.Revit.DB.ElementSet) _
    As Autodesk.Revit.UI.Result Implements IExternalCommand.Execute


        m_application = commandData.Application
        BaseKeyTable = New List(Of TableKeyNotes)
        m_IncorrectElements = New List(Of IncorrectKeyInElements)
        m_DataTableExcel = New List(Of DataTableExcel)
        m_DataTableExcelShort = New List(Of DataTableExcel)

        ' Setup a default message in case any exceptions are thrown that we have not
        ' explicitly handled. On failure the message will be displayed by Revit
        message = "The sample failed"
        Execute = Autodesk.Revit.UI.Result.Failed
        m_tempFileName = "Template1.xlsx"     ' имя файла шаблона
        m_param_project = "Bouwdeel"          ' имя параметра проекта

        Try

            If Not CheckKeyNote() Then      ' проверяем наличие загруженной таблицы кодов
                Return Execute
            End If

            GetTableKeyNote()         ' получить таблицу с кодам и параметрами
            GetAllElementsWithKey()   ' получить все элементы с кодами

            If m_DataTableExcel.Count = 0 Then
                message = "Not any elements with keynote"
                Return Execute
            End If

            CalculateTableData()            ' рассчитать данные для таблицы

            CheckExcel()                    ' проверяем наличие excel

            If Not GetTemplateFile() Then   ' получить шаблон файла и создать копию для заполнения
                excel.Quit()
                Return Execute
            End If

            newWorkbook.Application.ActiveWindow.WindowState = XlWindowState.xlMinimized
            newWorkbook.Application.ActiveWindow.WindowState = XlWindowState.xlMaximized

            WriteToExcel()            ' записываем данные в файл  excel

            ' change our result to successful
            Execute = Autodesk.Revit.UI.Result.Succeeded
            Return Execute
        Catch ex As Exception
            message = ex.Message
            Return Execute
        End Try

        Return Execute

        'Dim browserForm As New FormExportSchedule

        'Using (browserForm)
        '    Dim initResult As Boolean = browserForm.Initialize(commandData.Application)

        '    If (initResult = False) Then
        '        Return Execute
        '    End If

        '    browserForm.ShowDialog()
        'End Using

        'Try
        '    ' Use Microsoft Office 2003 or later.
        '    Dim excel As MsExcel.Application
        '    excel = Nothing

        '    ' search the entire Revit project, grouping elements by category
        '    Dim sortedElements As System.Collections.Generic.Dictionary(Of String, ElementSet)
        '    sortedElements = GetSortedNonSymbolElements(commandData.Application, commandData.Application.ActiveUIDocument.Document)
        '    If (sortedElements Is Nothing) Then
        '        Return Execute
        '    End If

        '    ' loop through all the categories found and send them to Microsoft Excel
        '    Dim iter As System.Collections.Generic.Dictionary(Of String, ElementSet).Enumerator = sortedElements.GetEnumerator()

        '    Do While (iter.MoveNext())
        '        ' the Revit map iterator provides access to the key as well as the values
        '        Dim categoryName As String = iter.Current.Key
        '        Dim elementSet As Autodesk.Revit.DB.ElementSet = iter.Current.Value

        '        Dim sendSuccess As Boolean = SendToExcel(excel, commandData.Application.Application, categoryName, elementSet)
        '        If (sendSuccess = False) Then
        '            Return Execute
        '        End If

        '    Loop

        '    ' change our result to successful
        '    Execute = Autodesk.Revit.UI.Result.Succeeded
        '    Return Execute
        'Catch ex As Runtime.InteropServices.COMException
        '    message = "Something wrong with Microsoft Office Object library."
        '    Return Execute
        'Catch ex As Exception
        '    message = ex.Message
        '    Return Execute
        'End Try

    End Function


    Public ReadOnly Property Document() As Autodesk.Revit.DB.Document
        Get
            Document = m_application.ActiveUIDocument.Document
        End Get
    End Property

    Public ReadOnly Property Application() As Autodesk.Revit.UI.UIApplication
        Get
            Application = m_application
        End Get
    End Property


    '''' <summary>
    '''' Получить файл шаблона
    '''' </summary>
    '''' <returns>Возвращает файл шаблона Excel для заполнения данными </returns>
    '''' <remarks></remarks>
    'Private Function ExportToFile() As MsExcel.Workbook

    '    ' Dim excel As MsExcel.Application = New MsExcel.ApplicationClass()

    '    'excel = New MsExcel.ApplicationClass()

    '    'If (excel Is Nothing) Then
    '    '    MsgBox("Excel not installed")
    '    '    Return Nothing
    '    'End If

    '    '' make excel visible so that operations made are visible to the user
    '    'excel.Visible = False

    '    If (GetTemplateFile()) Then
    '        ' Если файл шаблоан найден и доступен - откроем его скрытно
    '        Dim workbook As MsExcel.Workbook = excel.Workbooks.Open(m_tempFileName)
    '        Return workbook
    '    Else
    '        Return Nothing
    '    End If

    'End Function


    ''' <summary>
    ''' Получить файл шаблона и создать новую книгу
    ''' </summary>
    ''' <returns>Да, если файл найден и доступен</returns>
    ''' <remarks></remarks>
    Private Function GetTemplateFile() As Boolean

        Dim directory As String
        Dim directory2 As String
        Dim nameProject As String

        directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        directory2 = Path.GetDirectoryName(Document.PathName)
        nameProject = Path.GetFileNameWithoutExtension(Document.PathName)

        m_tempFileName = directory + "\" + m_tempFileName   ' полное имя с путем
        m_newFileName = directory2 + "\" + nameProject       ' полное имя с путем

        Dim strFileExists As String
        strFileExists = Dir(m_tempFileName)
        If strFileExists = "" Then
            MsgBox("Template file doesn't exist")
            Return False
        Else
            ' Если файл шаблон найден и доступен - откроем его скрытно
            oldWorkbook = excel.Workbooks.Open(m_tempFileName)
            Try
                oldWorkbook.SaveCopyAs(m_newFileName)
                oldWorkbook.Close(SaveChanges:=False)
                newWorkbook = excel.Workbooks.Open(m_newFileName)
            Catch ex As Exception
                MsgBox("Impossible to save new Excel file.")
                Return False

            End Try
            Return True
        End If
    End Function

    ''' <summary>
    ''' Запустить Excel и открыть заполненный файл щаблона
    ''' </summary>  
    ''' <remarks></remarks>
    Sub CheckExcel()

        excel = New MsExcel.ApplicationClass()

        If (excel Is Nothing) Then
            MsgBox("Excel not installed")
            Return
        End If

        'List<Element> all_rebar = new List<Element>();
        '// элементы, включенные в группу не обрабатываются
        '            FilteredElementCollector collector = New FilteredElementCollector(doc);
        '            all_rebar = collector.WherePasses(New ElementClassFilter(TypeOf (Rebar))).OfType < Element > ().ToList();


        'IEnumerable<MarkR> mark = rebar.Select(x => new MarkR(
        '        doc.GetElement(x.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString(),
        '        doc.GetElement(x.get_Parameter(BuiltInParameter.REBAR_SHAPE).AsElementId()).Name,
        '        Math.Round(x.get_Parameter(BuiltInParameter.REBAR_ELEM_LENGTH).AsDouble() / 0.00328, 0))).Distinct();


        ' make excel visible so that operations made are visible to the user
        excel.Visible = True


    End Sub

    Private Sub CalculateTableData()
        ' получить уникальный список: сорт1, сорт2, ключ, марка, парам1, парам2, парам3

        Dim IEuni_table = From x In m_DataTableExcel
                          Select x.sort1, x.sort2, x.kode_key, x.mark, x.param1, x.param2, x.param3
                          Distinct


        Dim iter As IEnumerator
        iter = IEuni_table.GetEnumerator

        Do While (iter.MoveNext())

            Dim my_count As List(Of DataTableExcel) = m_DataTableExcel.FindAll(Function(value As DataTableExcel)
                                                                                   Return value.kode_key = iter.Current.kode_key And
                                                                                   value.sort1 = iter.Current.sort1 And
                                                                                   value.sort2 = iter.Current.sort2 And
                                                                                   value.mark = iter.Current.mark And
                                                                                   value.param1 = iter.Current.param1 And
                                                                                   value.param2 = iter.Current.param2 And
                                                                                   value.param3 = iter.Current.param3
                                                                               End Function)
            my_count(0).kol = my_count.Count
            m_DataTableExcelShort.Add(my_count(0))
        Loop
    End Sub
    Private Sub WriteToExcel()

        Dim workSheet As Worksheet = newWorkbook.Worksheets.Item(2)
        workSheet.Activate()

        Dim line As Integer = 10
        Dim sum1 As Double = 0
        Dim sum2 As Double = 0

        ' "шапка" для элемента
        workSheet.Cells(line, 1).Value = m_DataTableExcelShort.Item(0).kode_key + " " + m_DataTableExcelShort.Item(0).kode_name

        'Dim r As Range = workSheet.Cells(line, 1)
        'r.HorizontalAlignment = HorizontalAlign.Left
        'r.Font.Bold = True

        workSheet.Cells(line, 1).HorizontalAlignment = XlHAlign.xlHAlignLeft
        workSheet.Cells(line, 1).Font.Bold = True
        workSheet.Cells(line, 2).Value = m_DataTableExcelShort.Item(0).Unit1
        workSheet.Cells(line, 3).Value = m_DataTableExcelShort.Item(0).Unit2
        line = line + 1

        For Each element In m_DataTableExcelShort

            workSheet.Cells(line, 1).Value = element.mark
            workSheet.Cells(line, 1).HorizontalAlignment = XlHAlign.xlHAlignRight
            workSheet.Cells(line, 4).Value = element.kol
            workSheet.Cells(line, 5).Value = element.param1
            workSheet.Cells(line, 6).Value = element.param2
            workSheet.Cells(line, 7).Value = element.param3
            workSheet.Cells(line, 8).Value = element.param4
            sum1 = sum1 + element.param4
            workSheet.Cells(line, 9).Value = element.kol
            workSheet.Cells(line, 10).Value = element.param5
            workSheet.Cells(line, 11).Value = element.param3
            workSheet.Cells(line, 12).Value = element.param6
            sum2 = sum2 + element.param6
            line = line + 1
        Next
        line = line + 1
        workSheet.Cells(line, 1).Value = "TOTAAL :"
        workSheet.Cells(line, 1).HorizontalAlignment = XlHAlign.xlHAlignRight
        workSheet.Cells(line, 1).Font.Bold = True
        workSheet.Cells(line, 8).Value = sum1
        workSheet.Cells(line, 8).Font.Bold = True
        workSheet.Cells(line, 8).Borders(XlBordersIndex.xlEdgeTop).LineStyle = 7
        workSheet.Cells(line, 12).Value = sum2
        workSheet.Cells(line, 12).Font.Bold = True
        workSheet.Cells(line, 12).Borders(XlBordersIndex.xlEdgeTop).LineStyle = 7
    End Sub

    Function CheckKeyNote() As Boolean

        ' получить данные по таблице с ключами
        keyTable = KeynoteTable.GetKeynoteTable(Document)

        If keyTable Is Nothing Then
            MsgBox("Table KeyNotes not loaded")
            Return False
        End If

        Return True

    End Function



    Private Sub GetAllElementsWithKey()


        Dim filter As Autodesk.Revit.DB.ElementIsElementTypeFilter
        filter = New Autodesk.Revit.DB.ElementIsElementTypeFilter(True)

        Dim key_impty As Boolean = False

        Dim collector As Autodesk.Revit.DB.FilteredElementCollector
        collector = New Autodesk.Revit.DB.FilteredElementCollector(Application.ActiveUIDocument.Document)
        collector.WherePasses(filter)
        Dim iter As IEnumerator
        iter = collector.GetElementIterator

        Do While (iter.MoveNext())

            Dim element As Autodesk.Revit.DB.Element
            element = iter.Current

            Dim list_param As List(Of Parameter) = element.GetParameters(m_param_project)

            If list_param.Count = 0 Then
                Continue Do
            End If

            ' должен иметь параметр проекта - по нему определяем включать в таблицу или нет
            m_def_param_project = list_param.Item(0).Definition

            Dim type As Autodesk.Revit.DB.Element
            type = Document.GetElement(element.GetTypeId())

            If Not type Is Nothing Then

                If Not type.Parameter(BuiltInParameter.KEYNOTE_PARAM) Is Nothing Then

                    Dim keynote As String = type.Parameter(BuiltInParameter.KEYNOTE_PARAM).AsString()

                    If Not keynote Is Nothing Then

                        ' проверим корректность кода и категории элемента
                        If CheckKeyInElements(element, keynote) Then
                            ' заполняем данные по текущему элементу
                            SetDataTableExcel(element, keynote)
                            ' m_AllElementsWithKey.Add(element)
                        Else
                            Dim problem As IncorrectKeyInElements
                            problem = New IncorrectKeyInElements(keynote, element.Id.IntegerValue, element.Name)
                            m_IncorrectElements.Add(problem)
                        End If


                    Else
                        m_AllElementsWithOutKey.Add(element)
                    End If

                End If


            End If

        Loop

        'If key_impty Then

        '    MsgBox("Find elements with empty keynotes")

        'End If

    End Sub


    '''' <summary>
    '''' the SendToExcel method takes a category and a set of elements that belong to that category
    '''' and finds the common properties between those elements. Excel is then sent the names and
    '''' values of these properties, adding the category as a new sheet
    '''' </summary>
    '''' <param name="excel">Excel application</param>
    '''' <param name="application">Revit application</param>
    '''' <param name="categoryName">name of category</param>
    '''' <param name="elementSet">the elements in the category</param>
    '''' <returns>If function succeed, return True. Otherwise False</returns>
    '''' <remarks></remarks>
    'Function SendToExcel(ByRef excel As MsExcel.Application, ByVal application As Autodesk.Revit.ApplicationServices.Application,
    '    ByVal categoryName As String, ByVal elementSet As Autodesk.Revit.DB.ElementSet) As Boolean

    '    'Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet
    '    Dim worksheet As MsExcel.Worksheet
    '    SendToExcel = False

    '    ' If excel is not running, then launch it which will result in one sheet remaining. If excel
    '    ' is already running then we need to create a new sheet.
    '    If (excel Is Nothing) Then
    '        excel = LaunchExcel()

    '        If (excel Is Nothing) Then
    '            Return SendToExcel
    '        End If

    '        worksheet = excel.ActiveSheet

    '    Else
    '        worksheet = excel.Worksheets.Add()

    '    End If

    '    ' Set the name of the sheet to be that of the category
    '    If (categoryName.Length > 31) Then
    '        worksheet.Name = categoryName.Substring(0, 31)
    '    Else
    '        worksheet.Name = categoryName
    '    End If


    '    ' using all the elements, find common properties amongst them
    '    Dim propertyNames As System.Collections.Generic.List(Of String) = GetCommonPropertyNames(application, elementSet)

    '    ' the first column in the sheet will be the Element ID of the element
    '    worksheet.Cells(1, 1).Value = "ID"

    '    ' now add the common property names as the column headers
    '    Dim propertyName As String
    '    Dim column As Integer = 2
    '    For Each propertyName In propertyNames
    '        worksheet.Cells(1, column).Value = propertyName
    '        column = column + 1
    '    Next

    '    ' devote a row to each element that belongs to the category
    '    Dim element As Autodesk.Revit.DB.Element
    '    Dim row As Integer = 2
    '    For Each element In elementSet

    '        ' first column is the element id - display it as an integer
    '        worksheet.Cells(row, 1).Value = element.Id.ToString()

    '        ' retrieve all the values, in string form for the common properties of the element
    '        Dim values As System.Collections.Generic.Dictionary(Of String, String) = GetValuesOfNamedProperties(application, element, propertyNames)

    '        column = 2
    '        If (values.Count > 0) Then

    '            For Each propertyName In propertyNames
    '                ' check to see if the element actually supports that property and then set the
    '                ' excel cell to be the value
    '                If (values.ContainsKey(propertyName)) Then
    '                    worksheet.Cells(row, column).Value = values.Item(propertyName)
    '                End If
    '                column = column + 1
    '            Next

    '            row = row + 1

    '        End If

    '    Next

    '    SendToExcel = True

    'End Function

    ''' <summary>
    ''' GetValuesOfNamedProperties takes a set of property names and looks up the values for those properties
    ''' returning them in a string form
    ''' </summary>
    ''' <param name="application">Revit application</param>
    ''' <param name="element">element</param>
    ''' <param name="propertyNames">names of properties</param>
    ''' <returns>The map stores the values according to the names of properties</returns>
    ''' <remarks></remarks>
    Function GetValuesOfNamedProperties(ByVal application As Autodesk.Revit.ApplicationServices.Application,
        ByVal element As Autodesk.Revit.DB.Element, ByVal propertyNames As System.Collections.Generic.List(Of String)) _
        As System.Collections.Generic.Dictionary(Of String, String)

        Dim values As System.Collections.Generic.Dictionary(Of String, String) = New System.Collections.Generic.Dictionary(Of String, String)

        Dim parameter As Autodesk.Revit.DB.Parameter

        ' loop through all the parameters that the element has
        For Each parameter In element.Parameters

            ' the name for the parameter is held in the parameter definition
            If propertyNames.Contains(parameter.Definition.Name) Then

                Dim stringValue As String = ""

                ' take the internal type of the parameter and convert it into a string
                Select Case parameter.StorageType

                    Case Autodesk.Revit.DB.StorageType.Double
                        stringValue = parameter.AsDouble

                        ' in the case of ElementId, retrieve the element, if possible, and use its name
                    Case Autodesk.Revit.DB.StorageType.ElementId
                        Dim paramElement As Autodesk.Revit.DB.Element = element.Document.GetElement(parameter.AsElementId)
                        If Not (paramElement Is Nothing) Then
                            stringValue = paramElement.Name
                        End If

                    Case Autodesk.Revit.DB.StorageType.Integer
                        stringValue = parameter.AsInteger

                    Case Autodesk.Revit.DB.StorageType.String
                        stringValue = parameter.AsString

                    Case Else

                End Select

                Try
                    values.Add(parameter.Definition.Name, stringValue)
                Catch
                End Try

            End If

        Next

        GetValuesOfNamedProperties = values

    End Function

    ''' <summary>
    ''' GetCommonPropertyNames takes a set of elements and seeks the property names that are common between them
    ''' If an element does not support any properties at all it is ignored.
    ''' The process of finding the common elements is done by collecting all the names of parameters for
    ''' the first element and then removing those that are not used by the other elements
    ''' </summary>
    ''' <param name="application">Revit application</param>
    ''' <param name="elementSet">elements</param>
    ''' <returns>A set of common properties</returns>
    ''' <remarks></remarks>
    Function GetCommonPropertyNames(ByVal application As Autodesk.Revit.ApplicationServices.Application,
        ByVal elementSet As Autodesk.Revit.DB.ElementSet) As System.Collections.Generic.List(Of String)

        Dim commonProperties As System.Collections.Generic.List(Of String) = New System.Collections.Generic.List(Of String)

        ' the first element that we handle (that has properties), a
        Dim addAllProperties As Boolean = True

        ' loop through all the elements passed to the method
        Dim iter As IEnumerator
        iter = elementSet.ForwardIterator

        Do While (iter.MoveNext())
            Dim element As Autodesk.Revit.DB.Element = iter.Current

            ' get the parameters that the element supports
            Dim parameters As Autodesk.Revit.DB.ParameterSet = element.Parameters

            If Not (parameters.IsEmpty) Then

                ' definitionNames will contain all the parameter names that this element supports
                Dim definitionNames As System.Collections.Generic.List(Of String) = New System.Collections.Generic.List(Of String)

                Dim paramIter As Autodesk.Revit.DB.ParameterSetIterator = parameters.ForwardIterator

                ' loop through all of the parameters and retrieve their names
                Do While paramIter.MoveNext()

                    Dim parameter As Autodesk.Revit.DB.Parameter = paramIter.Current
                    Dim definitionName As String = parameter.Definition.Name

                    If (addAllProperties) Then
                        commonProperties.Add(definitionName)
                    End If

                    definitionNames.Add(definitionName)

                Loop

                If addAllProperties Then
                    addAllProperties = False
                Else

                    'now loop through all of the common parameter we have found so far and see if they are
                    'supported by the element. If not, then remove it from the common set
                    Dim commonDefinitionNamesToRemove As System.Collections.Generic.List(Of String) = New System.Collections.Generic.List(Of String)

                    Dim commonDefinitionName As String
                    For Each commonDefinitionName In commonProperties

                        If (definitionNames.Contains(commonDefinitionName) = False) Then
                            commonDefinitionNamesToRemove.Add(commonDefinitionName)
                        End If

                    Next

                    ' remove the uncommon parameters
                    Dim commonDefinitionNameToRemove As String
                    For Each commonDefinitionNameToRemove In commonDefinitionNamesToRemove

                        commonProperties.Remove(commonDefinitionNameToRemove)

                    Next

                End If

            End If

        Loop

        GetCommonPropertyNames = commonProperties

    End Function

    ''' <summary>
    ''' GetSortedNonSymbolElements searches the entire Revit project and sorts the elements based
    ''' upon category. Revit Symbols (Types) are ignored as we are only interested in instances of elements.
    ''' </summary>
    ''' <param name="application">Revit application</param>
    ''' <param name="document">Revit document</param>
    ''' <returns>The map stores all the elements according to their category</returns>
    ''' <remarks></remarks>
    Function GetSortedNonSymbolElements(ByVal application As Autodesk.Revit.UI.UIApplication,
        ByVal document As Autodesk.Revit.DB.Document) As System.Collections.Generic.Dictionary(Of String, ElementSet)

        GetSortedNonSymbolElements = New System.Collections.Generic.Dictionary(Of String, ElementSet)

        'get a set of element which is not elementType
        Dim filter As Autodesk.Revit.DB.ElementIsElementTypeFilter
        filter = New Autodesk.Revit.DB.ElementIsElementTypeFilter(True)
        Dim collector As Autodesk.Revit.DB.FilteredElementCollector
        collector = New Autodesk.Revit.DB.FilteredElementCollector(application.ActiveUIDocument.Document)
        collector.WherePasses(filter)
        Dim iter As IEnumerator
        iter = collector.GetElementIterator

        Do While (iter.MoveNext())

            Dim element As Autodesk.Revit.DB.Element
            element = iter.Current
            If Not (TypeOf element Is Autodesk.Revit.DB.ElementType) Then
                ' retrieve the category of the element
                Dim category As Autodesk.Revit.DB.Category
                category = element.Category

                If Not (category Is Nothing) Then
                    Dim elementSet As Autodesk.Revit.DB.ElementSet

                    Try
                        ' if this is a category that we have seen before, then add this element to that set,
                        ' otherwise create a new set and add the element to it.
                        If GetSortedNonSymbolElements.ContainsKey(category.Name) Then
                            elementSet = GetSortedNonSymbolElements.Item(category.Name)
                        Else
                            elementSet = application.Application.Create.NewElementSet()
                            GetSortedNonSymbolElements.Add(category.Name, elementSet)
                        End If

                        elementSet.Insert(element)
                    Catch ex As Exception
                    End Try

                End If
            End If
        Loop

    End Function

    Sub GetTableKeyNote()

        ' правила для ключевых меток
        Dim element As TableKeyNotes = New TableKeyNotes("26.", "STRUCTUURELEMENTEN GEWAPEND BETON", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.00.", "structuurelementen gewapend beton - algemeen", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.01", "algemeen - betonstudie", "PM", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)


        element = New TableKeyNotes("26.10.", "materialen algemeen", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.11.", "materialen wapening", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.11.10.", "materialen - wapening/staven en netten", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.11.11.", "materialen - wapening/staven en netten - staven", "VH", "kg", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.11.12.", "materialen - wapening/staven en netten - netten", "VH", "kg", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.12.", "materialen - beton", "VH", "kg", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.12.10.", "materialen - beton/stortklaar beton", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        'element = New TableKeyNotes("26.12.11.", "materialen - beton/stortklaar beton - met staaf- en netwapening", "PM", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        'BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.12.13.", "materialen - beton/stortklaar beton - zichtbeton", "PM", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.12.20.", "materialen - beton/geprefabriceerd beton", "PM", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.12.30.", "materialen - beton/architectonisch beton", "PM", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.13.", "materialen - bekistingen", "PM", "m2", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.14.", "materialen - nabehandeling", "PM", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.15.", "materialen - chemische verankering", "PM", "st.", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.16.", "materialen - thermische onderbreking", "PM", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.20.", "ter plaatse gestorte elementen - algemeen", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.21.", "ter plaatse gestorte elementen - wanden", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.21.10.", "ter plaatse gestorte elementen - wanden/traditionele bekisting", "FH", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.22.", "ter plaatse gestorte elementen - kolommen", "FH", "m3", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.23.", "ter plaatse gestorte elementen - balken", "FH", "m3", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.25.", "ter plaatse gestorte elementen - trappen en bordessen", "FH", "m3", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.26.", "ter plaatse gestorte elementen - draagvloeren", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        'Dim param1 As ParamColumn = New ParamColumn(BuiltInParameter.CURVE_ELEM_LENGTH, TypeParam.Instance)
        'Dim param2 As ParamColumn = New ParamColumn(BuiltInParameter.WALL_USER_HEIGHT_PARAM, TypeParam.Instance)
        'Dim param3 As ParamColumn = New ParamColumn(BuiltInParameter.WALL_ATTR_WIDTH_PARAM, TypeParam.Symbol)
        'Dim sort As ParamColumn = New ParamColumn(BuiltInParameter.WALL_BASE_CONSTRAINT, TypeParam.Instance)

        'element = New TableKeyNotes("26.26.10.", "", "", "", BuiltInCategory.OST_Walls, param1, param2, param3, sort)
        'BaseKeyTable.Add(element)


        Dim param1 As ParamColumn = New ParamColumn(BuiltInParameter.CURVE_ELEM_LENGTH, TypeParam.Instance)
        Dim param2 As ParamColumn = New ParamColumn(BuiltInParameter.WALL_USER_HEIGHT_PARAM, TypeParam.Instance)
        Dim param3 As ParamColumn = New ParamColumn(BuiltInParameter.WALL_ATTR_WIDTH_PARAM, TypeParam.Symbol)
        Dim sort As ParamColumn = New ParamColumn(BuiltInParameter.WALL_BASE_CONSTRAINT, TypeParam.Instance)

        element = New TableKeyNotes("26.12.11.", "materialen - beton/stortklaar beton - met staaf- en netwapening", "PM", "", BuiltInCategory.OST_Walls, param1, param2, param3, sort)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.26.20.", "ter plaatse gestorte elementen - draagvloeren/verloren bekisting", "FH", "m3", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.26.30.", "ter plaatse gestorte elementen - draagvloeren/verloren bekisting", "FH", "m3", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.26.31.", "ter plaatse gestorte elementen - draagvloeren/breedplaatvloeren - prefab breedplaten", "FH", "m2", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.26.32.", "ter plaatse gestorte elementen - draagvloeren/breedplaatvloeren - opstort", "FH", "m3", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.30.", "prefab elementen - algemeen", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.34.", "prefab elementen - deur- en raamlateien", "PM", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.35.", "prefab elementen - trappen en bordessen", "FH", "st.", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.36.", "prefab elementen - draagvloeren", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.36.10", "prefab elementen - draagvloeren/welfsels", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.36.12", "prefab elementen - draagvloeren/welfsels - met druklaag /prefab", "FH", "m2", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.36.13", "prefab elementen - draagvloeren/welfsels - met druklaag /druklaag", "PM", "m3", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.36.20.", "prefab elementen - draagvloeren/voorgespannen welfsels", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.36.22.", "prefab elementen - draagvloeren/voorgespannen welfsels - met druklaag / prefab", "FH", "m2", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.36.23.", "prefab elementen - draagvloeren/voorgespannen welfsels - met druklaag / druklaag", "PM", "m3", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.37.", "prefab elementen - uitkragende elementen", "", "", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.37.10.", "prefab elementen - uitkragende elementen/balkons", "FH", "m2", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.37.20.", "prefab elementen - uitkragende elementen/galerijen", "FH", "m2", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)

        element = New TableKeyNotes("26.37.30.", "prefab elementen - uitkragende elementen/galerijen", "FH", "m2", Nothing, Nothing, Nothing, Nothing, Nothing)
        BaseKeyTable.Add(element)







    End Sub

    Function CheckKeyInElements(element As Element, key_element As String) As Boolean

        ' проверим совпадение кода и категории элемента
        Dim category As BuiltInCategory = element.Category.BuiltInCategory
        Dim a1 As Boolean
        a1 = BaseKeyTable.Exists(Function(v1 As TableKeyNotes)
                                     Return v1.Keynote = key_element
                                 End Function)

        Dim a2 As Boolean
        a2 = BaseKeyTable.Exists(Function(v1 As TableKeyNotes)
                                     Return v1.category = category
                                 End Function)

        Dim a3 As Boolean
        a3 = BaseKeyTable.Exists(Function(v1 As TableKeyNotes)
                                     Return v1.Keynote = "26.12.11."
                                 End Function)


        Return BaseKeyTable.Exists(Function(v1 As TableKeyNotes)
                                       Return v1.Keynote = key_element And v1.category = category
                                   End Function)


    End Function

    Sub SetDataTableExcel(element As Element, key_element As String)

        ' получить объект для извлечения параметров типа
        Dim symbol As Element = Document.GetElement(element.GetTypeId())
        ' Dim key_element As String = element.Parameter(BuiltInParameter.KEYNOTE_NUMBER).AsString()
        Dim category As BuiltInCategory = element.Category.BuiltInCategory

        Dim key_table As TableKeyNotes = BaseKeyTable.Find(Function(v1 As TableKeyNotes)
                                                               Return v1.Keynote = key_element And v1.category = category
                                                           End Function)

        ' заполняем данные по элементу таблицы
        Dim kode_name As String = key_table.Name
        Dim kode_key As String = key_table.Keynote

        Dim sort1 As String
        If key_table.sort1.type = TypeParam.Instance Then
            sort1 = element.Parameter(key_table.sort1.param).AsValueString()
        Else
            sort1 = symbol.Parameter(key_table.sort1.param).AsValueString()
        End If

        Dim sort2 As String = element.Parameter(m_def_param_project).AsString()
        Dim unit1 As String = key_table.Unit1
        Dim unit2 As String = key_table.Unit2

        Dim mark As String = element.Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString()

        Dim param1 As Double
        If key_table.param1.type = TypeParam.Instance Then
            param1 = element.Parameter(key_table.param1.param).AsDouble()
        Else
            param1 = symbol.Parameter(key_table.param1.param).AsDouble()
        End If

        Dim param2 As Double
        If key_table.param2.type = TypeParam.Instance Then
            param2 = element.Parameter(key_table.param2.param).AsDouble()
        Else
            param2 = symbol.Parameter(key_table.param2.param).AsDouble()
        End If

        Dim param3 As Double
        If key_table.param3.type = TypeParam.Instance Then
            param3 = element.Parameter(key_table.param3.param).AsDouble()
        Else
            param3 = symbol.Parameter(key_table.param3.param).AsDouble()
        End If

        Dim elem_table As DataTableExcel = New DataTableExcel(kode_key, kode_name, unit1, unit2, sort1, sort2, mark, 1, param1, param2, param3)
        m_DataTableExcel.Add(elem_table)

    End Sub

End Class
