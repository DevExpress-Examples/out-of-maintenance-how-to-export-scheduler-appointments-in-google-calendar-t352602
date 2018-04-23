Imports Microsoft.VisualBasic
Imports System
Imports DevExpress.XtraScheduler
Imports Google.Apis.Calendar.v3
Imports Google.Apis.Calendar.v3.Data
Imports DevExpress.XtraScheduler.iCalendar.Components
Imports DevExpress.XtraScheduler.iCalendar.Native
Imports System.IO
Imports System.Text
Imports System.Diagnostics
Imports System.Globalization
Imports System.Collections.Generic

Namespace GoogleCalendarExample
	' Userful links:
	'      http://stackoverflow.com/questions/32726096/creating-recurring-events-via-google-calendar-api-v3        
	'      https://developers.google.com/google-apps/calendar/recurringevents

	Public Class CalendarExporter
		Public Sub New(ByVal calendarService As CalendarService, ByVal calendarEntry As CalendarListEntry)
			CalendarService = calendarService
			CalendarEntry = calendarEntry
		End Sub

		Private privateCalendarService As CalendarService
		Private Property CalendarService() As CalendarService
			Get
				Return privateCalendarService
			End Get
			Set(ByVal value As CalendarService)
				privateCalendarService = value
			End Set
		End Property
		Private privateCalendarEntry As CalendarListEntry
		Private Property CalendarEntry() As CalendarListEntry
			Get
				Return privateCalendarEntry
			End Get
			Set(ByVal value As CalendarListEntry)
				privateCalendarEntry = value
			End Set
		End Property

		Public Sub Export(ByVal appointments As AppointmentBaseCollection)
			For Each apt In appointments
				ExportAppointment(apt)
			Next apt
		End Sub

		Public Sub Export(ByVal appointments As IList(Of Appointment))
			For Each apt In appointments
				ExportAppointment(apt)
			Next apt
		End Sub

		Private Sub ExportAppointment(ByVal apt As Appointment)
			Dim aptType As AppointmentType = apt.Type
            If aptType = AppointmentType.Pattern Then
                EnsurePatternId(apt)
            ElseIf aptType <> AppointmentType.Normal Then
                Dim eventPatternId As String = EnsurePatternId(apt.RecurrencePattern)
                Debug.Assert((Not String.IsNullOrEmpty(eventPatternId)))
                If aptType = AppointmentType.Occurrence Then
                    Return
                End If
                Dim instancesRequest As EventsResource.InstancesRequest = CalendarService.Events.Instances(CalendarEntry.Id, eventPatternId)
                Dim calculator As OccurrenceCalculator = OccurrenceCalculator.CreateInstance(apt.RecurrencePattern.RecurrenceInfo)
                instancesRequest.OriginalStart = GoogleCalendarUtils.ConvertEventDateTime(calculator.CalcOccurrenceStartTime(apt.RecurrenceIndex)).DateTimeRaw
                Dim occurrenceEvents As Events = instancesRequest.Execute()
                Debug.Assert(occurrenceEvents.Items.Count = 1)
                Dim occurrence As [Event] = occurrenceEvents.Items(0)
                If aptType = AppointmentType.ChangedOccurrence Then
                    Me.AssignProperties(apt, occurrence)
                ElseIf aptType = AppointmentType.DeletedOccurrence Then
                    occurrence.Status = "cancelled"
                End If
                Dim changedOccurrence As [Event] = CalendarService.Events.Update(occurrence, CalendarEntry.Id, occurrence.Id).Execute()
                apt.CustomFields("eventId") = changedOccurrence.Id
                Log.WriteLine(String.Format("Exported {0} occurrance: {1}, id={2}", If((aptType = AppointmentType.ChangedOccurrence), "changed", "deleted"), apt.Subject, changedOccurrence.Id))
                Return
            End If

			Dim instance As [Event] = Me.CreateEvent(aptType)
			AssignProperties(apt, instance)

			Dim result As [Event] = CalendarService.Events.Insert(instance, CalendarEntry.Id).Execute()
			Log.WriteLine(String.Format("Exported appointment: {0}, id={1}", apt.Subject, result.Id))
		End Sub

		Private Function EnsurePatternId(ByVal pattern As Appointment) As String
			Dim eventId As String = TryCast(pattern.CustomFields("eventId"), String)
			If (Not String.IsNullOrEmpty(eventId)) Then
				Return eventId
			End If
			Dim converter As VRecurrenceConverter = VRecurrenceConverter.CreateInstance(pattern.RecurrenceInfo.Type)
			Dim rule As DevExpress.XtraScheduler.iCalendar.Internal.IWritable = TryCast(converter.FromRecurrenceInfo(pattern.RecurrenceInfo, pattern), DevExpress.XtraScheduler.iCalendar.Internal.IWritable)
			Dim someRule As String = String.Empty
			Using stream As New MemoryStream()
				Using tw As New StreamWriter(stream)
					Dim cw As New iCalendarWriter(tw)
					rule.WriteToStream(cw)
				End Using
				stream.Flush()
				someRule = Encoding.ASCII.GetString(stream.ToArray())
			End Using

			Dim instance2 As [Event] = CreateEvent(pattern.Type)
			AssignProperties(pattern, instance2)
			instance2.Recurrence = New String() { "RRULE:" & someRule }

			Dim patternEvent As [Event] = CalendarService.Events.Insert(instance2, CalendarEntry.Id).Execute()
			pattern.CustomFields("eventId") = patternEvent.Id
			Return patternEvent.Id
		End Function

		Private Sub AssignProperties(ByVal apt As Appointment, ByVal instance As [Event])
			instance.Summary = apt.Subject
			instance.Description = apt.Description
			instance.Location = apt.Location

			instance.Start = GoogleCalendarUtils.ConvertEventDateTime(apt.Start)
			instance.End = GoogleCalendarUtils.ConvertEventDateTime(apt.End)
		End Sub

		Private Function CreateEvent(ByVal type As AppointmentType) As [Event]
			Return New [Event]()
		End Function
	End Class
End Namespace
