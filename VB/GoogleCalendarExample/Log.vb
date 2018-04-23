Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace GoogleCalendarExample
	Public NotInheritable Class Log
		Private Shared logAction As Action(Of String)
		Private Sub New()
		End Sub
		Public Shared Sub Register(ByVal logDelegate As Action(Of String))
			logAction = logDelegate
		End Sub

		Public Shared Sub WriteLine(ByVal message As String)
			If logAction Is Nothing Then
				Return
			End If
			logAction(message)
		End Sub
	End Class
End Namespace
