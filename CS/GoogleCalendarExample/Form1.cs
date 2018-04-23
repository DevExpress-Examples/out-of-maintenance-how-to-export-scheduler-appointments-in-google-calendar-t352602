using DevExpress.XtraScheduler;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace GoogleCalendarExample {
    public partial class Form1 : Form {
        const string DefaultCaption = "Google Calendar Importer";
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "GoogleCalendarExample";
        bool isConnected = false;
        public Form1() {
            InitializeComponent();
            this.schedulerControl1.Storage.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("eventId", "eventId"));
            GoogleCalendarExample.Log.Register(Log);
            UpdateFormState();

            this.schedulerControl1.Storage.AppointmentsDeleted += OnAppointmentsChanged;
            this.schedulerControl1.Storage.AppointmentsChanged += OnAppointmentsChanged;
            this.schedulerControl1.Storage.AppointmentsInserted += OnAppointmentsChanged;
        }
               

        CalendarService CalendarService { get; set; }
        bool LockStorageEvents { get; set; }

        void OnAppointmentsChanged(object sender, PersistentObjectsEventArgs e) {
            if (LockStorageEvents || !isConnected)
                return;
            CalendarListEntry calendarEntry = this.cbCalendars.SelectedValue as CalendarListEntry;
            CalendarExporter exporter = new CalendarExporter(CalendarService, calendarEntry);
            exporter.Export(e.Objects as IList<Appointment>);
        }

        void OnBtnConnectClick(object sender, EventArgs e) {
            UserCredential credential;
            using (var stream =
                new FileStream("secret\\client_secret.json", FileMode.Open, FileAccess.Read)) {
                string credPath = Environment.CurrentDirectory;
                credPath = Path.Combine(credPath, ".credentials");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Log("Credential file saved to: " + credPath);
            }
            isConnected = true;
            // Create Google Calendar API service.
            CalendarService = new CalendarService(new BaseClientService.Initializer() {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            var calendarListRequst = CalendarService.CalendarList.List();
            CalendarList calendarList = calendarListRequst.Execute();
            foreach (var item in calendarList.Items) {
                Log(item.Summary);
            }
            cbCalendars.DisplayMember = "Summary";
            cbCalendars.DataSource = calendarList.Items;

            cbCalendars.SelectedValueChanged += OnCbCalendarsSelectedValueChanged;
            UpdateFromGoogleCalendar();
        }
        void OnCbCalendarsSelectedValueChanged(object sender, EventArgs e) {
            UpdateFromGoogleCalendar();
        }

        void UpdateFromGoogleCalendar() {
            LockStorageEvents = true;

            CalendarListEntry calendarEntry = this.cbCalendars.SelectedValue as CalendarListEntry;
            Calendar calendar = CalendarService.Calendars.Get(calendarEntry.Id).Execute();
            EventsResource.ListRequest listRequest = CalendarService.Events.List(calendarEntry.Id);
            listRequest.MaxResults = 10000;
            Events events = listRequest.Execute();
            Log("Loaded {0} events", events.Items.Count);
            this.schedulerStorage1.Appointments.Items.Clear();
            this.schedulerStorage1.BeginUpdate();
            try {
                CalendarImporter importer = new CalendarImporter(this.schedulerStorage1);
                importer.Import(events.Items);
            } finally {
                this.schedulerStorage1.EndUpdate();
            }
            SetStatus(String.Format("Loaded {0} events", events.Items.Count));
            UpdateFormState();

            LockStorageEvents = false;
        }

        void UpdateFormState() {
            if (CalendarService == null) {
                this.cbCalendars.Enabled = false;
                this.btnConnect.Enabled = true;
            } else {
                this.cbCalendars.Enabled = true;
                this.btnConnect.Enabled = false;
            }
        }
                
        void OnBtnRefrehsClick(object sender, EventArgs e) {
            UpdateFromGoogleCalendar();
        }

        #region Logging
        void Log(string message) {
            this.tbLog.AppendText(message + "\r\n");
        }
        void Log(string format, params object[] args) {
            Log(string.Format(format, args));
        }
        void SetStatus(string message) {
            this.tsStatus.Text = message;
        }
        #endregion

    }
}
