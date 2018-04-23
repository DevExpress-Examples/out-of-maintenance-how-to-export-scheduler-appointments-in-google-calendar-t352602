Imports Microsoft.VisualBasic
Imports DevExpress.XtraScheduler
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Calendar.v3
Imports Google.Apis.Calendar.v3.Data
Imports Google.Apis.Services
Imports Google.Apis.Util.Store
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Threading
Imports System.Windows.Forms

Namespace GoogleCalendarExample
	Partial Public Class Form1
		Inherits Form
		Private Const DefaultCaption As String = "Google Calendar Importer"
		Private Shared Scopes() As String = { CalendarService.Scope.Calendar }
		Private Shared ApplicationName As String = "GoogleCalendarExample"
		Private isConnected As Boolean = False
		Public Sub New()
			InitializeComponent()
			Me.schedulerControl1.Storage.Appointments.CustomFieldMappings.Add(New AppointmentCustomFieldMapping("eventId", "eventId"))
			GoogleCalendarExample.Log.Register(AddressOf Log)
			UpdateFormState()

			AddHandler Me.schedulerControl1.Storage.AppointmentsDeleted, AddressOf OnAppointmentsChanged
			AddHandler Me.schedulerControl1.Storage.AppointmentsChanged, AddressOf OnAppointmentsChanged
			AddHandler Me.schedulerControl1.Storage.AppointmentsInserted, AddressOf OnAppointmentsChanged
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
		Private privateLockStorageEvents As Boolean
		Private Property LockStorageEvents() As Boolean
			Get
				Return privateLockStorageEvents
			End Get
			Set(ByVal value As Boolean)
				privateLockStorageEvents = value
			End Set
		End Property

		Private Sub OnAppointmentsChanged(ByVal sender As Object, ByVal e As PersistentObjectsEventArgs)
			If LockStorageEvents OrElse (Not isConnected) Then
				Return
			End If
			Dim calendarEntry As CalendarListEntry = TryCast(Me.cbCalendars.SelectedValue, CalendarListEntry)
			Dim exporter As New CalendarExporter(CalendarService, calendarEntry)
			exporter.Export(TryCast(e.Objects, IList(Of Appointment)))
		End Sub

		Private Sub OnBtnConnectClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnConnect.Click
			Dim credential As UserCredential
			Using stream = New FileStream("secret\client_secret.json", FileMode.Open, FileAccess.Read)
				Dim credPath As String = Environment.CurrentDirectory
				credPath = Path.Combine(credPath, ".credentials")

				credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None, New FileDataStore(credPath, True)).Result
				Log("Credential file saved to: " & credPath)
			End Using
			isConnected = True
			' Create Google Calendar API service.
			CalendarService = New CalendarService(New BaseClientService.Initializer() With {.HttpClientInitializer = credential, .ApplicationName = ApplicationName})
			Dim calendarListRequst = CalendarService.CalendarList.List()
			Dim calendarList As CalendarList = calendarListRequst.Execute()
			For Each item In calendarList.Items
				Log(item.Summary)
			Next item
			cbCalendars.DisplayMember = "Summary"
			cbCalendars.DataSource = calendarList.Items

			AddHandler cbCalendars.SelectedValueChanged, AddressOf OnCbCalendarsSelectedValueChanged
			UpdateFromGoogleCalendar()
		End Sub
		Private Sub OnCbCalendarsSelectedValueChanged(ByVal sender As Object, ByVal e As EventArgs)
			UpdateFromGoogleCalendar()
		End Sub

		Private Sub UpdateFromGoogleCalendar()
			LockStorageEvents = True

			Dim calendarEntry As CalendarListEntry = TryCast(Me.cbCalendars.SelectedValue, CalendarListEntry)
			Dim calendar As Calendar = CalendarService.Calendars.Get(calendarEntry.Id).Execute()
			Dim listRequest As EventsResource.ListRequest = CalendarService.Events.List(calendarEntry.Id)
			listRequest.MaxResults = 10000
			Dim events As Events = listRequest.Execute()
			Log("Loaded {0} events", events.Items.Count)
			Me.schedulerStorage1.Appointments.Items.Clear()
			Me.schedulerStorage1.BeginUpdate()
			Try
				Dim importer As New CalendarImporter(Me.schedulerStorage1)
				importer.Import(events.Items)
			Finally
				Me.schedulerStorage1.EndUpdate()
			End Try
			SetStatus(String.Format("Loaded {0} events", events.Items.Count))
			UpdateFormState()

			LockStorageEvents = False
		End Sub

		Private Sub UpdateFormState()
			If CalendarService Is Nothing Then
				Me.cbCalendars.Enabled = False
				Me.btnConnect.Enabled = True
			Else
				Me.cbCalendars.Enabled = True
				Me.btnConnect.Enabled = False
			End If
		End Sub

		Private Sub OnBtnRefrehsClick(ByVal sender As Object, ByVal e As EventArgs) Handles btnRefresh.Click
			UpdateFromGoogleCalendar()
		End Sub

		#Region "Logging"
		Private Sub Log(ByVal message As String)
			Me.tbLog.AppendText(message & Constants.vbCrLf)
		End Sub
		Private Sub Log(ByVal format As String, ParamArray ByVal args() As Object)
			Log(String.Format(format, args))
		End Sub
		Private Sub SetStatus(ByVal message As String)
			Me.tsStatus.Text = message
		End Sub
		#End Region

	End Class
End Namespace
