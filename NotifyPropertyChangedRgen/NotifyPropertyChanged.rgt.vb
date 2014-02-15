Imports System
Imports System.Linq
Imports System.CodeDom
Imports EnvDTE
Imports EnvDTE80
Imports System.Text.RegularExpressions
Imports Kodeo.Reegenerator.Generators
Imports Microsoft.VisualBasic.MyServices

Imports ThisClass = NotifyPropertyChangedRgen.NotifyPropertyChanged
Imports System.IO

Imports TaggedSegmentWriter = NotifyPropertyChangedRgen.TagManager(Of NotifyPropertyChangedRgen.NotifyPropertyChanged_GenAttribute).TaggedSegmentWriter
Imports FoundTaggedSegment = NotifyPropertyChangedRgen.TagManager(Of NotifyPropertyChangedRgen.NotifyPropertyChanged_GenAttribute).FoundTaggedSegment

''' <summary>
''' To use this renderer, attach to the target file. And add AutoGenerateAttribute to the class
''' </summary>
''' <remarks></remarks>
Partial Public Class NotifyPropertyChanged


    Shared ReadOnly INotifyPropertyChangedName As String = GetType(System.ComponentModel.INotifyPropertyChanged).FullName
    Shared ReadOnly SharedTagManager As TagManager(Of NotifyPropertyChanged_GenAttribute)
    ReadOnly AttrType As Type = GetType(NotifyPropertyChanged_GenAttribute)

    Shared Sub New()
        Dim tagName = (New NotifyPropertyChanged_GenAttribute).TagName

        SharedTagManager = New TagManager(Of NotifyPropertyChanged_GenAttribute)()
    End Sub
    ReadOnly Property TagManager As TagManager(Of NotifyPropertyChanged_GenAttribute)
        Get
            Return SharedTagManager
        End Get
    End Property

    Public Sub RenderLibrary()
        'Check for existing class
        Dim prj = Me.ProjectItem.Project

        Dim className = LibraryRenderer.DefaultClassName
        Dim classes = prj.GetCodeElements(Of EnvDTE.CodeClass)()
        Dim classFullname = prj.DefaultNamespace.DotJoin(className)
        Dim matchingClasses As List(Of EnvDTE.CodeClass) = Nothing
        classes.TryGetValue(classFullname, matchingClasses)

        Dim classItem As ProjectItem = Nothing
        If matchingClasses Is Nothing Then
            'Class not found, generate
            Dim filePath = Path.Combine(prj.FullPath, className + ".vb")
            'if file exists, warn then quit
            If File.Exists(filePath) Then
                MsgBox(String.Format("Trying to add {0} to project, but file {1} already exists", className, filePath))
                Return
            End If
            'Create new empty file
            File.WriteAllText(filePath, "")
            prj.CheckOut()
            'Add it to the project
            classItem = prj.DteObject.ProjectItems.AddFromFile(filePath)
        Else
            'Class found, get corresponding project item
            classItem = matchingClasses.First.ProjectItem
        End If

        'Open file for editing
        Dim wasOpen = classItem.IsOpen(EnvDTE.Constants.vsViewKindCode)
        If Not wasOpen Then classItem.Open(EnvDTE.Constants.vsViewKindCode)
        Dim textDoc = classItem.Document.ToTextDocument

        Dim genInfo As TaggedSegmentWriter = New TaggedSegmentWriter() With
                        {.SearchStart = textDoc.StartPoint,
                         .InsertStart = textDoc.StartPoint,
                        .SearchEnd = textDoc.EndPoint,
                        .SegmentType = SegmentTypes.Region}
        If TagManager.IsAnyOutdated(genInfo) Then
            'generate text if outdated
            Dim extRgen = New LibraryRenderer()

            Dim code = extRgen.RenderToString()
            genInfo.Content = code
            TagManager.InsertOrReplace(genInfo)
            classItem.Save()
        End If

        'restore to previous state
        If Not wasOpen Then classItem.Document.Close()
    End Sub

    ''' <summary>
    ''' Render within target file, instead of into a separate file
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RenderWithinTarget()

        Dim undoCtx = DTE.UndoContext
        undoCtx.Open(AttrType.Name)
        Try

            Dim validClasses As CodeClass2() = GetValidClasses()

            Dim sw As New Stopwatch
            Dim hasError = False
            '!for each class 
            For Each cc In validClasses
                sw.Start()

                Dim classWriter = New TaggedSegmentWriter() With {.Class = cc}

                '!generate
                GenerateInClass(classWriter)

                '!if also doing derivedClasses
                If classWriter.GenAttribute.ApplyToDerivedClasses Then

                    '!for each subclass
                    For Each derivedC In cc.GetSubclasses
                        Dim childInfo = New TaggedSegmentWriter() With {.TriggeringBaseClass = cc, .Class = derivedC}
                        'generate
                        GenerateInClass(childInfo)
                        'combine status
                        If childInfo.HasError Then
                            classWriter.HasError = True
                            classWriter.Status.AppendLine(childInfo.Status.ToString)
                        End If
                    Next
                End If

                'if there's error
                If classWriter.HasError Then
                    hasError = True
                    MsgBox(classWriter.Status.ToString)
                End If
                'finish up
                sw.Stop()
                DebugWriteLine(String.Format("Finished {0} in {1}", cc.Name, sw.Elapsed))
                sw.Reset()
            Next
            'render shared library
            RenderLibrary()

            'if there's error
            If hasError Then
                'undo everything
                undoCtx.SetAborted()
            Else
                undoCtx.Close()
                'automatically save, since we are changing the target file
                Dim doc = Me.ProjectItem.DteObject.Document
                'if anything is changed, save
                If doc IsNot Nothing AndAlso Not doc.Saved Then doc.Save()
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
            If undoCtx.IsOpen Then undoCtx.SetAborted()
        End Try
    End Sub

    Private Function GetValidClasses() As CodeClass2()
        Dim validClasses As CodeClass2() = Nothing
        'get only classes marked with the attribute
        If validClasses Is Nothing Then validClasses = (From cc In ProjectItem.GetClassesWithAttribute(AttrType)).ToArray()
        Return validClasses
    End Function

    ''' <summary>
    ''' Expand Auto properties into a normal properties, so we can insert Notify statement in the setter
    ''' </summary>
    ''' <param name="tsWriter"></param>
    ''' <remarks></remarks>
    Sub ExpandAutoProperties(tsWriter As TaggedSegmentWriter)
        Dim autoProps = tsWriter.Class.GetAutoProperties.Where(Function(x) Not (New NotifyPropertyChanged_GenAttribute(x).IsIgnored))
        For Each p In autoProps
            ExpandAutoProperty(p, tsWriter)
        Next
    End Sub
    ''' <summary>
    ''' Expand auto property into a normal property
    ''' </summary>
    ''' <param name="prop"></param>
    ''' <param name="parentWriter"></param>
    ''' <remarks></remarks>
    Sub ExpandAutoProperty(ByVal prop As CodeProperty2, ByVal parentWriter As TaggedSegmentWriter)
        'Save existing doc comment
        Dim commentStart = prop.GetCommentStartPoint
        Dim comment = commentStart.GetText(prop.GetStartPoint(vsCMPart.vsCMPartWholeWithAttributes))
        'Get its attribute
        Dim propAttrs = prop.GetText(vsCMPart.vsCMPartAttributesWithDelimiter)
        'Interface implementation 
        Dim interfaceImpl = prop.GetInterfaceImplementation


        Dim tsWriter = New TaggedSegmentWriter(parentWriter) With {
                            .SegmentType = SegmentTypes.Region,
                            .TagComment = String.Format("{0} auto expanded by", prop.Name)}
        'only do this once, since once it is expanded it will no longer be detected as auto property
        tsWriter.GenAttribute.RegenMode = GeneratorAttribute.RegenModes.Once

        Dim completeProp = GetIsolatedOutput(Sub() OutProperty(tsWriter.CreateTaggedRegionName, prop.Name, prop.Type.AsFullName, comment, propAttrs, interfaceImpl))

        'Replace all code starting from comment to endPoint of the property
        Dim ep = commentStart.CreateEditPoint
        ep.ReplaceText(prop.EndPoint, completeProp, EnvDTE.vsEPReplaceTextOptions.vsEPReplaceTextAutoformat Or vsEPReplaceTextOptions.vsEPReplaceTextNormalizeNewlines)

    End Sub

    ''' <summary>
    ''' Generate code that will notify other propertyName different from the member with the attribute.
    ''' </summary>
    ''' <param name="genAttr"></param>
    ''' <param name="parentWriter"></param>
    ''' <remarks>
    ''' Example Add NotifyPropertyChanged_GenAttribute with ExtraNotifications="OtherProperty1,OtherProperty2" to SomeProperty.
    ''' This method will generate code for Notify("OtherProperty1") and Notify("OtherProperty2") within that member
    ''' This is useful for Property that affects other Property, or a method that affects another property.
    ''' This has the advantage of generation/compile time verification of the properties
    ''' </remarks>
    Private Function GenInMember_ExtraNotifications(ByVal genAttr As NotifyPropertyChanged_GenAttribute, ByVal parentWriter As TaggedSegmentWriter) As String


        'Render extra notifications (notifications for other related properties)
        If genAttr.ExtraNotifications = "" Then Return Nothing
        'also split by space to trim it
        Dim extras = genAttr.ExtraNotifications.Split({",", " "}, StringSplitOptions.RemoveEmptyEntries)

        'Verify that all properties listed in ExtraNotification actually exists
        Dim invalids = genAttr.ValidateExtraNotifications(parentWriter.Class, extras)
        If invalids.Count > 0 Then
            parentWriter.HasError = True
            parentWriter.Status.AppendFormat("Properties:{0} to be notified are not found in the class", String.Join(", ", invalids))
            Return ""
        End If

        Return String.Format("Me.NotifyChanged({0})", String.Join(",", extras.Select(Function(x) x.Quote)))


    End Function

    ''' <summary>
    ''' Generates code in 
    ''' </summary>
    ''' <param name="genAttr"></param>
    ''' <param name="parentWriter"></param>
    ''' <remarks></remarks>
    Private Sub GenInMember(genAttr As NotifyPropertyChanged_GenAttribute, parentWriter As TaggedSegmentWriter)
        Dim prop = genAttr.ParentProperty
        '!Parent can be either CodeFunction(only for ExtraNotifications) or CodeProperty
        Dim code As String

        Select Case genAttr.GenerationType
            Case NotifyPropertyChanged_GenAttribute.GenerationTypes.NotifyOnly
                'Only notification
                code = String.Format("Me.NotifyChanged({0})", prop.Name)
            Case Else
                'Set property backing field and notify
                code = If(genAttr.ParentProperty IsNot Nothing, String.Format("Me.SetPropertyAndNotify(_{0}, value, ""{0}"")", prop.Name), "")
        End Select

        'Extra notifications
        Dim extraNotifyCode = GenInMember_ExtraNotifications(genAttr, parentWriter)
        code = code.Conjoin(vbNewLine, extraNotifyCode)

        'Code Element, could be property setter or a method
        Dim codeElement = CType(If(prop IsNot Nothing, prop.Setter, genAttr.ParentFunction), CodeElement2)
        Dim memberWriter = New TaggedSegmentWriter(parentWriter) With
                 {.GenAttribute = genAttr,
                 .SearchStart = codeElement.StartPoint,
                 .SearchEnd = codeElement.EndPoint,
                 .Content = code,
                 .SegmentType = SegmentTypes.Statements}

        'Find insertion point
        Dim insertPoint As EditPoint
        Dim insertTag = TagManager.FindInsertionPoint(memberWriter)
        If insertTag Is Nothing Then
            '!No insertion point tag specified, by default insert as last line of setter
            insertPoint = codeElement.EndPoint.CreateEditPoint()
            insertPoint.StartOfLine()
        Else
            '!InsertPoint Tag found, insert right after it
            insertPoint = insertTag.EndPoint.CreateEditPoint
            insertPoint.LineDown()
            insertPoint.StartOfLine()
        End If

        memberWriter.InsertStart = insertPoint
        TagManager.InsertOrReplace(memberWriter)

    End Sub
    Private Sub GenInMembers(ByVal tsWriter As TaggedSegmentWriter)
        '!Generate in properties
        Dim props = tsWriter.Class.GetProperties.ToArray
        Dim propAttrs = (From p In props Select New NotifyPropertyChanged_GenAttribute(p)).ToArray

        Dim functions = From f In tsWriter.Class.GetFunctions Where f.AsElement.HasAttribute(AttrType)
        Dim funcAttrs = (From f In functions Select New NotifyPropertyChanged_GenAttribute(f)).ToArray



        Dim dpFields = tsWriter.Class.GetDependencyProperties


        Dim notDpField = Function(x As NotifyPropertyChanged_GenAttribute) Not dpFields.Any(Function(dp) dp.Name = x.ParentProperty.Name & "Property")
        Dim notIgnored = Function(x As NotifyPropertyChanged_GenAttribute) Not (x.IsIgnored)


        Dim propsWithSetters = propAttrs.Where(Function(x) x.ParentProperty.Setter IsNot Nothing)

        '?filter out property for DependencyProperties 
        Dim validMembers = funcAttrs.Concat(propsWithSetters.
                                                Where(notDpField).
                                                Where(notIgnored))

        For Each pa As NotifyPropertyChanged_GenAttribute In validMembers

            GenInMember(pa, tsWriter)
        Next

    End Sub

    Private Shared Function GetFirstAncestorImplementing(ByVal ancestorClasses As System.Collections.Generic.IEnumerable(Of CodeClass2), ByVal interfaceName As String) As CodeClass2
        Return ancestorClasses.
                        FirstOrDefault(Function(x) x.ImplementedInterfaces.Cast(Of CodeInterface).
                                           Any(Function(i) i.FullName = interfaceName))

    End Function
    Private Sub GenerateNotifyFunctions(tsWriter As TaggedSegmentWriter)
        If tsWriter.IsTriggeredByBaseClass Then Return
        Dim firstMember = tsWriter.Class.Members.Cast(Of EnvDTE.CodeElement).FirstOrDefault
        If firstMember Is Nothing Then Return '?if there's no member, there won't be any properties. Skip

        '!If INotify is already implemented by base class, do not generate (only generate tag)
        Dim ancestorClasses = tsWriter.Class.GetAncestorClasses
        Dim ancestorImplementingINPC = GetFirstAncestorImplementing(ancestorClasses, INotifyPropertyChangedName)
        Dim inotifierFullname As String = String.Format("{0}.{1}", Me.ProjectItem.Project.DefaultNamespace, LibraryRenderer.INotifierName)
        Dim ancestorImplementingINotifier = GetFirstAncestorImplementing(ancestorClasses, inotifierFullname)
        Dim code As String = ""
        If ancestorImplementingINotifier IsNot Nothing Then
            code = String.Format("'{0} already implemented by {1}", LibraryRenderer.INotifierName, ancestorImplementingINotifier.FullName)
        ElseIf ancestorImplementingINPC IsNot Nothing Then
            code = String.Format("'{0} already implemented by {1}{2}{3}", ancestorImplementingINotifier.FullName, INotifyPropertyChangedName, vbNewLine,
                                 GetIsolatedOutput(Sub() OutFunctions(tsWriter.Class.FullName, False)))
        Else
            code = GetIsolatedOutput(Sub() OutFunctions(tsWriter.Class.FullName, True))
        End If

        Dim insertPoint = tsWriter.Class.GetStartPoint(vsCMPart.vsCMPartBody).CreateEditPoint
        insertPoint.StartOfLine()
        Dim body = tsWriter.Class.GetText(vsCMPart.vsCMPartBody)
        Dim header = tsWriter.Class.GetText(vsCMPart.vsCMPartHeader)
        Dim x = tsWriter.Class.GetText(vsCMPart.vsCMPartBodyWithDelimiter)
        Dim name = tsWriter.Class.GetText(vsCMPart.vsCMPartName)
        'copy info, instead of using the passed parameter, prevent unintentionally using irrelevant property set 
        ' by other code
        Dim newInfo = New TaggedSegmentWriter(tsWriter) With
                             {.SearchStart = tsWriter.Class.StartPoint,
                              .SearchEnd = tsWriter.Class.EndPoint,
                              .InsertStart = insertPoint,
                              .Content = code,
                              .SegmentType = SegmentTypes.Region,
                              .TagComment = "INotifier Functions"}

        newInfo.GenAttribute.SegmentClass = "INotifierFunctions"

        Dim isUpdated = TagManager.InsertOrReplace(newInfo)
        If isUpdated Then tsWriter.Class.AddInterfaceIfNotExists(LibraryRenderer.INotifierName)


    End Sub
    Public Sub GenerateInClass(writer As TaggedSegmentWriter)

        GenerateNotifyFunctions(writer)
        ExpandAutoProperties(writer)
        GenInMembers(writer)
        AppendWarning(writer)
    End Sub

    Private Sub AppendWarning(ByVal writer As TaggedSegmentWriter)
        Dim autoProperties As CodeProperty2() = writer.Class.GetAutoProperties.Where(Function(x) Not New NotifyPropertyChanged_GenAttribute(x).IsIgnored).ToArray
        '?Warn unprocesssed autoproperties
        If autoProperties.Count > 0 Then
            writer.HasError = True
            With writer.Status
                .AppendFormat("{0} Autoproperties skipped:", writer.Class.Name).AppendLine()

                For Each ap In autoProperties
                    .AppendIndent(1, ap.Name).AppendLine()
                Next
            End With


        End If
    End Sub

    Public Overrides Function Render() As RenderResults

        RenderWithinTarget()
        Return New RenderResults("'Because of the way custom tool works a file has to be generated. This file can be safely ignored.")

    End Function



End Class