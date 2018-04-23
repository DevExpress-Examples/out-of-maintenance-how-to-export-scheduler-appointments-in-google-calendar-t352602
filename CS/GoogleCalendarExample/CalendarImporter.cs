using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraScheduler;

namespace GoogleCalendarExample {
    public class CalendarImporter {
        SchedulerStorage storage;
        public CalendarImporter(SchedulerStorage storage) {
            this.storage = storage;
        }

        public SchedulerStorage Storage { get { return storage; } }

        public void Import(IList<Event> events) {
            Dictionary<string, Appointment> patternHash = new Dictionary<string, Appointment>();
            Dictionary<string, Event> occurrenceHash = new Dictionary<string, Event>();

            Storage.BeginUpdate();
            try {
                RecurrencePatternParser parser = new RecurrencePatternParser(storage);
                foreach (Event item in events) {
                    Appointment appointment = null;
                    if (item.RecurringEventId != null) { //occurrence?
                        occurrenceHash.Add(item.Id, item);
                    } else if (item.Recurrence != null) { //recurrence                    
                        appointment = parser.Parse(item.Recurrence, GoogleCalendarUtils.ConvertDateTime(item.Start), GoogleCalendarUtils.ConvertDateTime(item.End)); //parse and create pattern
                        patternHash.Add(item.Id, appointment);
                    } else { //normal appointment
                        appointment = storage.CreateAppointment(AppointmentType.Normal);
                        AssingTimeIntervalPropertiesTo(appointment, item);
                    }
                    if (appointment == null)
                        continue;
                    AssignCommonPropertiesTo(appointment, item);
                    appointment.CustomFields["eventId"] = item.Id;
                    Storage.Appointments.Add(appointment);
                }
                LinkOccurrencesWithPatterns(occurrenceHash, patternHash);
            } finally {
                Storage.EndUpdate();
            }
        }
        void LinkOccurrencesWithPatterns(Dictionary<string, Event> occurrenceHash, Dictionary<string, Appointment> patternHash) {
            foreach(KeyValuePair<string, Event> occurrenceEntry in occurrenceHash) {
                Event occurrenceEvent = occurrenceEntry.Value;
                string patternId = occurrenceEntry.Value.RecurringEventId;
                if(occurrenceEvent != null && patternHash.ContainsKey(patternId)) {
                    Appointment pattern = patternHash[patternId];
                    AppointmentType exceptionType = (occurrenceEvent.Sequence == null) ? AppointmentType.DeletedOccurrence : AppointmentType.ChangedOccurrence;
                    if(exceptionType == AppointmentType.DeletedOccurrence)
                        CreateDeletedOccurrence(pattern, occurrenceEvent);
                    else
                        CreateChangedOccurrence(pattern, occurrenceEvent);
                }
            }
        }
        void CreateChangedOccurrence(Appointment pattern, Event occurrenceEvent) {
            int indx = CalculateOccurrenceIndex(pattern, occurrenceEvent.OriginalStartTime);
            Appointment newOccurrence = pattern.CreateException(AppointmentType.ChangedOccurrence, indx);
            newOccurrence.Start = GoogleCalendarUtils.ConvertDateTime(occurrenceEvent.Start);
            newOccurrence.End = GoogleCalendarUtils.ConvertDateTime(occurrenceEvent.End);
            AssignCommonPropertiesTo(newOccurrence, occurrenceEvent);
        }
        void CreateDeletedOccurrence(Appointment pattern, Event occurrenceEvent) {
            int indx = CalculateOccurrenceIndex(pattern, occurrenceEvent.OriginalStartTime);
            Appointment exception = pattern.CreateException(AppointmentType.DeletedOccurrence, indx);
        }

        int CalculateOccurrenceIndex(Appointment pattern, EventDateTime originalStartTime) {
            OccurrenceCalculator calculator = OccurrenceCalculator.CreateInstance(pattern.RecurrenceInfo);
            return calculator.FindOccurrenceIndex(GoogleCalendarUtils.ConvertDateTime(originalStartTime), pattern);
        }
        void AssignCommonPropertiesTo(Appointment target, Event source) {
            target.Subject = source.Summary;
            target.Description = source.Description;
        }
        void AssingTimeIntervalPropertiesTo(Appointment target, Event source) {
            if(!source.Start.DateTime.HasValue)
                target.AllDay = true;
            target.Start = GoogleCalendarUtils.ConvertDateTime(source.Start);
            target.End = GoogleCalendarUtils.ConvertDateTime(source.End);
        }            
    }
}
