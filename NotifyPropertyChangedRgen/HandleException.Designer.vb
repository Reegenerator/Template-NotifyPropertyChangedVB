Imports System
Imports System.Linq
Imports Kodeo.Reegenerator
Imports Wrappers = Kodeo.Reegenerator.Wrappers
Imports EnvDTE

'-------------------------------------------------------
' Automatically generated with Kodeo's Reegenerator
' Generator: RgenTemplate (internal)
' Generation date: 2014-02-14 01:59
' Generated by: GATSU-DEV\Tedy.Pranolo
' -------------------------------------------------------

<System.CodeDom.Compiler.GeneratedCodeAttribute("Reegenerator", "2.0.5.0")>  _
Partial Public Class HandleException
    Inherits CodeRendererEx
    
    '''<summary>
    '''Renders the code as defined in the source script file.
    '''</summary>
    '''<returns></returns>
    Public Overrides Function Render() As kodeo.reegenerator.generators.renderresults
        Me.Output.Write(" '")
        Me.Output.WriteLine
        Return New Kodeo.Reegenerator.Generators.RenderResults(Me.Output.ToString)
    End Function
End Class
