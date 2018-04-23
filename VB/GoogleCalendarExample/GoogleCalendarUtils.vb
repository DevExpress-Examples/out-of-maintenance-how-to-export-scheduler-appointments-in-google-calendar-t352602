Imports Microsoft.VisualBasic
Imports Google.Apis.Calendar.v3.Data
Imports System

Namespace GoogleCalendarExample
	Public NotInheritable Class GoogleCalendarUtils
		Private Sub New()
		End Sub
		Public Shared Function ConvertDateTime(ByVal start As EventDateTime) As DateTime
			If start.DateTime.HasValue Then
				Return start.DateTime.Value
			End If
			Return DateTime.Parse(start.Date)
		End Function

		Public Shared Function ConvertEventDateTime(ByVal dateTime As DateTime) As EventDateTime
			Dim result As New EventDateTime()
			result.DateTime = dateTime

			result.TimeZone = ToIana(TimeZoneInfo.Local.Id)
			Return result
		End Function

		Private Shared Function ToIana(ByVal tzid As String) As String
			If tzid.Equals("UTC", StringComparison.Ordinal) Then
				Return "Etc/UTC"
			End If

			Dim tzdbSource = NodaTime.TimeZones.TzdbDateTimeZoneSource.Default
			Dim tzi = TimeZoneInfo.FindSystemTimeZoneById(tzid)
			If tzi Is Nothing Then
				Return Nothing
			End If
			Dim ianaTzid = tzdbSource.MapTimeZoneId(tzi)
			If ianaTzid Is Nothing Then
				Return Nothing
			End If
			Return tzdbSource.CanonicalIdMap(ianaTzid)
		End Function
	End Class
End Namespace
