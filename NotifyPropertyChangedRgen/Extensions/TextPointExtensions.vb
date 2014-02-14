Imports System.Runtime.CompilerServices
Imports EnvDTE

Public Module TextPointExtensions

#Region "Text manipulation using EditPoint and TextPoint"

    <Extension>
    Function ToTextDocument(doc As EnvDTE.Document) As TextDocument
        Return CType(doc.Object("TextDocument"), TextDocument)
    End Function

    <Extension>
    Public Function GetLineText(point As TextPoint) As String
        Dim start = point.CreateEditPoint
        start.StartOfLine()
        Dim endP = point.CreateEditPoint
        endP.EndOfLine()
        Return start.GetText(endP)

    End Function
    <Extension>
    Public Function GetText(doc As EnvDTE.Document) As String

        Dim textdoc = CType(doc.Object("TextDocument"), TextDocument)

        Return textdoc.StartPoint.CreateEditPoint.GetText(textdoc.EndPoint)
    End Function
    <Extension>
    Public Function GetLineTextAndNeighbor(point As TextPoint) As String
        Dim start = point.CreateEditPoint
        start.LineUp()
        start.StartOfLine()
        Dim endP = point.CreateEditPoint
        endP.LineDown()
        endP.EndOfLine()

        Return start.GetText(endP)

    End Function
    ''' <summary>
    ''' Inserts text and format the text (=Format Selection command)
    ''' </summary>
    ''' <param name="tp"></param>
    ''' <param name="text"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function InsertAndFormat(tp As TextPoint, text As String) As TextPoint
        Dim start = tp.CreateEditPoint
        'preserve editPoint
        Dim ep = tp.CreateEditPoint()
        ep.Insert(text)
        start.SmartFormat(ep)
        Return ep
    End Function




    ''' <summary>
    ''' Unlike  CharRight, CharRightExact counts newline \r\n as two instead of one char.
    ''' </summary>
    ''' <param name="point"></param>
    ''' <param name="count"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' DTE functions that moves editpoint counts newline as single character, 
    ''' since we get the character count from regular regex not the DTE find, the char count is slightly off
    ''' </remarks>
    <Extension>
    Public Function CharRightExact(point As EditPoint, count As Integer) As EditPoint
        Return CharMoveExact(point, count, 1)
    End Function

    ''' <summary>
    ''' See CharMoveExact
    ''' </summary>
    ''' <param name="point"></param>
    ''' <param name="count"></param>
    ''' <returns></returns>
    ''' <remarks>See CharMoveExact</remarks>
    <Extension>
    Public Function CharLeftExact(point As EditPoint, count As Integer) As EditPoint

        Return CharMoveExact(point, count, -1)
    End Function

    ''' <summary>
    ''' Moves cursor/editpoint exactly.
    ''' </summary>
    ''' <param name="point"></param>
    ''' <param name="count"></param>
    ''' <param name="direction"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' DTE functions that moves editpoint counts newline as single character, 
    ''' since we get the character count from regular regex not the DTE find, the char count is slightly off
    ''' </remarks>
    Public Function CharMoveExact(point As EditPoint, count As Integer, direction As Integer) As EditPoint
        Do While count > 0
            'Normalize
            If direction > 1 Then
                direction = 1
            ElseIf direction < 0 Then
                direction = -1
            End If


            'If we are asking 1 and getting 2, its a newline. This is a quirk/feature of EnvDTE where all its functions treats newline as single character
            If point.GetText(direction).Length = 2 Then
                count -= 1
            End If
            If direction < 0 Then point.CharLeft() Else point.CharRight()
            count -= 1
        Loop
        Return point
    End Function

#End Region


End Module
