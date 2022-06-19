Imports MsExcel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports Autodesk.Revit.DB

Public Class FormExportSchedule
    Private m_application As Autodesk.Revit.UI.UIApplication
    Private m_initialized As Boolean
    Private m_scheduler As System.Collections.Generic.SortedList(Of ElementId, String)
    Private m_newFileName As String
    Private m_tempFileName As String
    Private newWorkbook As MsExcel.Workbook
    Private excel As MsExcel.Application


    Public Function Initialize(ByVal Application As Autodesk.Revit.UI.UIApplication) As Boolean


        ' Dim m_scheduler = New System.Collections.Generic.SortedList(Of ElementId, String)

        Try
            m_application = Application
            If Not (m_application Is Nothing) Then
                m_initialized = True
            End If




            'm_scheduler = LoadBindings()
            'Dim values = New System.Collections.Generic.List(Of String)
            'Dim v As String

            'For Each v In m_scheduler.Values
            '    values.Add(v)
            'Next

            'Me.ComboBoxScheduler.DataSource = values
            'Me.ComboBoxScheduler.DisplayMember = Name

            'If (m_scheduler.Count = 0) Then
            '    MsgBox("No scheduler to export")
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public ReadOnly Property Initialized() As Boolean
        Get
            Initialized = m_initialized
        End Get
    End Property

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




    Private Function LoadBindings() As System.Collections.Generic.SortedList(Of ElementId, String) ' Autodesk.Revit.DB.ViewSchedule)

        LoadBindings = New System.Collections.Generic.SortedList(Of ElementId, String)
        Dim filter As Autodesk.Revit.DB.ElementIsElementTypeFilter
        filter = New Autodesk.Revit.DB.ElementIsElementTypeFilter(True)
        Dim collector As Autodesk.Revit.DB.FilteredElementCollector
        collector = New Autodesk.Revit.DB.FilteredElementCollector(Application.ActiveUIDocument.Document)
        collector.WherePasses(filter)
        Dim iter As IEnumerator
        iter = collector.GetElementIterator

        Do While (iter.MoveNext())

            Dim element As Autodesk.Revit.DB.Element
            element = iter.Current

            If (TypeOf element Is Autodesk.Revit.DB.ViewSchedule) Then

                'Dim view As Autodesk.Revit.DB.ViewSchedule
                'view = Document.GetElement(element.Id)
                ' m_schedulerId.Add(element.Id.IntegerValue)

                'Dim new_sch As Sch
                'new_sch = New Sch
                'new_sch.schudelerId = element.Id
                'new_sch.schudelerName = element.Name

                LoadBindings.Add(element.Id, element.Name)

            End If


            'If Not (TypeOf element Is Autodesk.Revit.DB.ElementType) Then
            '    ' retrieve the category of the element
            '    Dim category As Autodesk.Revit.DB.Category
            '    category = element.Category

            '    If Not (category Is Nothing) Then
            '        Dim elementSet As Autodesk.Revit.DB.ElementSet

            '        Try
            '            ' if this is a category that we have seen before, then add this element to that set,
            '            ' otherwise create a new set and add the element to it.
            '            If GetSortedNonSymbolElements.ContainsKey(category.Name) Then
            '                elementSet = GetSortedNonSymbolElements.Item(category.Name)
            '            Else
            '                elementSet = Application.Application.Create.NewElementSet()
            '                GetSortedNonSymbolElements.Add(category.Name, elementSet)
            '            End If

            '            elementSet.Insert(element)
            '        Catch ex As Exception
            '        End Try

            '    End If
            'End If
        Loop













        'Try
        '    Dim bindingsRoot As System.Windows.Forms.TreeNode
        '    bindingsRoot = Me.BindingsTree.Nodes.Add("Bindings")

        '    Dim bindingsMap As Autodesk.Revit.DB.BindingMap
        '    bindingsMap = Document.ParameterBindings

        '    Dim iterator As Autodesk.Revit.DB.DefinitionBindingMapIterator
        '    iterator = bindingsMap.ForwardIterator

        '    Do While (iterator.MoveNext)
        '        Dim elementBinding As Autodesk.Revit.DB.ElementBinding
        '        elementBinding = iterator.Current

        '        ' get the name of the parameter 
        '        Dim definition As Autodesk.Revit.DB.Definition
        '        definition = iterator.Key

        '        Dim definitionNode As System.Windows.Forms.TreeNode = Nothing

        '        '  Note: the description of parameter binding is as follows:
        '        '  "a parameter definition is bound to elements within one or more categories."
        '        '  But this seems to return a one-to-one map.  
        '        '  The following for loop is a workaround. 

        '        '  do we have it in the node? 
        '        '  if yes, use the exisiting one. 
        '        Dim node As System.Windows.Forms.TreeNode
        '        For Each node In bindingsRoot.Nodes
        '            If (node.Text = definition.Name) Then
        '                definitionNode = node
        '            End If
        '        Next

        '        ' if the new parameter, add a new node. 
        '        If (definitionNode Is Nothing) Then
        '            definitionNode = bindingsRoot.Nodes.Add(definition.Name)
        '        End If

        '        ' add the category name.  
        '        If (Not elementBinding Is Nothing) Then
        '            Dim categories As Autodesk.Revit.DB.CategorySet
        '            categories = elementBinding.Categories

        '            Dim category As Autodesk.Revit.DB.Category
        '            For Each category In categories
        '                If (Not category Is Nothing) Then
        '                    definitionNode.Nodes.Add(category.Name)
        '                End If
        '            Next
        '        End If
        '    Loop

        '    Return True
        'Catch ex As Exception
        '    Return False
        'End Try

    End Function

    Private Sub ButtonExport_Click(sender As Object, e As EventArgs) Handles ButtonExport.Click


        excel = New MsExcel.ApplicationClass()

        If (excel Is Nothing) Then
            MsgBox("Excel not installed")
            'Return Nothing
        End If

        ' make excel visible so that operations made are visible to the user
        excel.Visible = False

        m_tempFileName = "Template1.xlsx"     ' имя файла шаблона

        ' newWorkbook As MsExcel.Workbook

        newWorkbook = ExportToFile() ' получить файл шаблона

        If Not (newWorkbook Is Nothing) Then
            newWorkbook.Application.ActiveWindow.WindowState = XlWindowState.xlMinimized
            newWorkbook.Application.ActiveWindow.WindowState = XlWindowState.xlNormal
            newWorkbook.Application.ActiveWindow.Width = 1
            newWorkbook.Application.ActiveWindow.Height = 1

            ' заполнить шаблон данными
            If Not (UpdateDataInTemplate()) Then
                MsgBox("Failed to update template file")
                Return
            End If

            Dim result As Object
            ' Выбрать папку и назначить файл для сохранения
            result = newWorkbook.Application.GetSaveAsFilename(InitialFilename:=m_newFileName, FileFilter:="Excel Files (*.xlsx), *.xlsx")

            If result.GetType().Name = "Boolean" Then
                ' была ОТМЕНА ничего не делаем
                newWorkbook.Close(SaveChanges:=False)
                excel.Quit()
                Return
            End If

            'Dim strFileExists As String
            'strFileExists = Dir(result)
            'If strFileExists = "" Then
            '    MsgBox("Failed to save file")
            'Else
            '    newWorkbook.SaveCopyAs(result)
            'End If

            ' newWorkbook.Application.ActiveWindow.Visible = False
            ' newWorkbook.Application.ActiveWindow.WindowState = XlWindowState.xlMinimized
            ' newWorkbook.Application.ActiveWindow.Width = 100
            ' newWorkbook.Application.ActiveWindow.Height = 100
            newWorkbook.SaveCopyAs(result)
            newWorkbook.Close(SaveChanges:=False)
            excel.Quit()
        End If

    End Sub



    Private Function UpdateDataInTemplate() As Boolean

        ' получить данные по таблице с ключами
        Dim keyTable As KeynoteTable
        keyTable = KeynoteTable.GetKeynoteTable(Document)

        If keyTable Is Nothing Then
            MsgBox("Table KeyNotes not loaded")
            Return False
        End If

        excel = New MsExcel.ApplicationClass()

        If (excel Is Nothing) Then
            MsgBox("Excel not installed")
            Return False
        End If

        Dim keyEntries As KeyBasedTreeEntries
        keyEntries = keyTable.GetKeyBasedTreeEntries()

        Dim keyEntr As KeynoteEntry

        For Each keyEntr In keyEntries

            Dim keystr As String
            keystr = keyEntr.Key
            ' выбрать все элементы с подобным ключом


        Next


        'Dim keystr As String

        'keystr = keyEntr.Key
        'keystr = keyEntr.KeynoteText

        Dim currentId As ElementId
        currentId = m_scheduler.Keys(Me.ComboBoxScheduler.SelectedIndex)

        If currentId = Nothing Then
            Return False
        End If

        Dim viewScheduler As Autodesk.Revit.DB.ViewSchedule

        Try
            viewScheduler = Document.GetElement(currentId)
        Catch ex As Exception
            Return False
        End Try

        'Dim body As String
        'body = viewScheduler.GetCellText(SectionType.Body, 0, 0)
        'body = viewScheduler.GetCellText(SectionType.Body, 1, 0)
        'body = viewScheduler.GetCellText(SectionType.Body, 1, 1)
        'body = viewScheduler.GetCellText(SectionType.Body, 2, 2)

        Dim table As TableData
        table = viewScheduler.GetTableData()


        Dim section As TableSectionData
        section = table.GetSectionData(SectionType.Body)



        Dim def_table As ScheduleDefinition
        def_table = viewScheduler.Definition

        Dim fields As System.Collections.Generic.List(Of SchedulableField)
        fields = def_table.GetSchedulableFields()


        'Dim field As SchedulableField

        Dim worksheet As MsExcel.Worksheet

        For Each worksheet In newWorkbook.Sheets

            If worksheet.Name = "Sheet1" Then

                '' заполним шапку таблицы
                ''Dim i As Integer
                ''i = 1

                'For i = 0 To def_table.GetFieldCount() - 1
                '    worksheet.Cells(1, i + 1).Value = def_table.GetField(i).GetName()  ' field.GetName(Document)
                '    'i = i + 1
                'Next


                'For s = 0 To table.NumberOfSections - 1
                Dim dataT As TableSectionData
                dataT = table.GetSectionData(SectionType.Body)

                ' число рядов и строк в секции
                Dim r, c As Integer
                r = dataT.NumberOfRows
                c = dataT.NumberOfColumns

                ' открыть цикл по секциям
                Dim line As Integer
                line = 1

                ' открыть цикл по данным секции - строки
                For row = 0 To r - 1
                    ' открыть цикл по данным секции - колонки
                    For col = 0 To c - 1

                        Dim cell As String
                        cell = viewScheduler.GetCellText(SectionType.Body, row, col)
                        worksheet.Cells(line, col + 1).Value = cell

                    Next
                    line = line + 1
                Next


                '' открыть цикл по данным секции - строки
                'For row = 1 To 10
                '    ' открыть цикл по данным секции - колонки
                '    For col = 1 To 10

                '        worksheet.Cells(row, col).Value = row + col

                '    Next

                'Next






                'Next

            End If

        Next

        Return True

    End Function


    Private Function GetElementsByKey(ByVal key As String) As System.Collections.Generic.List(Of Integer)

        GetElementsByKey = New List(Of Integer)

        Dim filter As Autodesk.Revit.DB.ElementIsElementTypeFilter
        filter = New Autodesk.Revit.DB.ElementIsElementTypeFilter(True)
        Dim collector As Autodesk.Revit.DB.FilteredElementCollector
        collector = New Autodesk.Revit.DB.FilteredElementCollector(Application.ActiveUIDocument.Document)
        collector.WherePasses(filter)
        Dim iter As IEnumerator
        iter = collector.GetElementIterator

        Do While (iter.MoveNext())

            Dim element As Autodesk.Revit.DB.Element
            element = iter.Current
            Dim type As Autodesk.Revit.DB.Element
            type = Document.GetElement(element.GetTypeId())
            If type.Parameter(BuiltInParameter.KEYNOTE_PARAM).AsString() = key Then
                GetElementsByKey.Add(element.Id.IntegerValue)
            End If
        Loop
    End Function



    Private Function GetAllElementsWithKey() As System.Collections.Generic.List(Of Element)

        GetAllElementsWithKey = New List(Of Element)

        Dim filter As Autodesk.Revit.DB.ElementIsElementTypeFilter
        filter = New Autodesk.Revit.DB.ElementIsElementTypeFilter(True)
        Dim collector As Autodesk.Revit.DB.FilteredElementCollector
        collector = New Autodesk.Revit.DB.FilteredElementCollector(Application.ActiveUIDocument.Document)
        collector.WherePasses(filter)
        Dim iter As IEnumerator
        iter = collector.GetElementIterator
        Dim noKey As Boolean = False
        Do While (iter.MoveNext())

            Dim element As Autodesk.Revit.DB.Element
            element = iter.Current
            Dim type As Autodesk.Revit.DB.Element
            type = Document.GetElement(element.GetTypeId())
            If type.Parameter(BuiltInParameter.KEYNOTE_PARAM).AsString().Length > 0 Then
                GetAllElementsWithKey.Add(element)
            Else
                noKey = True
            End If
        Loop

        If noKey Then
            MsgBox("Were find elements without notekey")
        End If

    End Function





    ''' <summary>
    ''' Получить файл шаблона
    ''' </summary>
    ''' <returns>Возвращает файл шаблона Excel для заполнения данными </returns>
    ''' <remarks></remarks>
    Private Function ExportToFile() As MsExcel.Workbook

        ' Dim excel As MsExcel.Application = New MsExcel.ApplicationClass()

        'excel = New MsExcel.ApplicationClass()

        'If (excel Is Nothing) Then
        '    MsgBox("Excel not installed")
        '    Return Nothing
        'End If

        '' make excel visible so that operations made are visible to the user
        'excel.Visible = False

        If (GetTemplateFile()) Then
            ' Если файл шаблоан найден и доступен - откроем его скрытно
            Dim workbook As MsExcel.Workbook = excel.Workbooks.Open(m_tempFileName)
            Return workbook
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' Получить файл шаблона
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
            Return True
        End If

    End Function



    Private Sub ComboBoxScheduler_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxScheduler.SelectedIndexChanged

    End Sub
End Class