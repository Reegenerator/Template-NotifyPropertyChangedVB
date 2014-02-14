Imports System.Runtime.CompilerServices
Imports EnvDTE

Module ProjectSolutionExtensions

#Region "Solution and project navigation helpers"

    ''' <summary>
    ''' Get solution name
    ''' </summary>
    ''' <param name="solution"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function GetName(solution As EnvDTE.Solution) As String
        Return solution.Properties.Item("Name").Value.ToString()
    End Function

    ''' <summary>
    ''' Get solution
    ''' </summary>
    ''' <param name="project"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function Solution(project As EnvDTE.Project) As Solution

        Return project.DTE.Solution
    End Function


    ''' <summary>
    ''' Get path to project node. To be used to select the node in Solution Explorer
    ''' </summary>
    ''' <param name="project"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function GetNodePath(project As EnvDTE.Project) As String
        Return String.Format("{0}\{1}", project.Solution.GetName, project.Name)

    End Function

    ''' <summary>
    '''  Get path to project item node. To be used to select the node in Solution Explorer
    ''' </summary>
    ''' <param name="projectItem"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function GetNodePath(projectItem As EnvDTE.ProjectItem) As String
        'Dim sln = TryCast(projectItem, EnvDTE.Solution)
        'If sln IsNot Nothing Then Return sln.GetName

        'Dim prj = TryCast(projectItem, EnvDTE.Project)
        'If prj IsNot Nothing Then Return prj.GetNodePath

        Return String.Format("{0}\{1}", projectItem.ContainingProject.GetNodePath, projectItem.Name)

    End Function

    ''' <summary>
    ''' Selects project item in Solution Explorer
    ''' </summary>
    ''' <param name="projectItem"></param>
    ''' <remarks></remarks>
    <Extension>
    Public Sub SelectSolutionExplorerNode(ByVal projectItem As EnvDTE.ProjectItem)
        CType(projectItem.DTE, EnvDTE80.DTE2).SelectSolutionExplorerNode(projectItem.GetNodePath)
    End Sub

    ''' <summary>
    ''' Selects project item in Solution Explorer
    ''' </summary>
    ''' <remarks></remarks>
    <Extension>
    Public Sub SelectSolutionExplorerNode(dte2 As EnvDTE80.DTE2, ByVal nodePath As String)
        Dim item As EnvDTE.UIHierarchyItem
        Try
            item = dte2.ToolWindows.SolutionExplorer.GetItem(nodePath)
            item.Select(vsUISelectionType.vsUISelectionTypeSelect)
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString())
        End Try
    End Sub
#End Region
End Module
