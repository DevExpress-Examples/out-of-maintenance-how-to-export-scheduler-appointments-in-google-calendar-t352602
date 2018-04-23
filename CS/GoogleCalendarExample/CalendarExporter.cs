using System;
using DevExpress.XtraScheduler;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using DevExpress.XtraScheduler.iCalendar.Components;
using DevExpress.XtraScheduler.iCalendar.Native;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Globalization;
using System.Collections.Generic;

namespace GoogleCalendarExample {
    // Userful links:
    //      http://stackoverflow.com/questions/32726096/creating-recurring-events-via-google-calendar-api-v3        
    //      https://developers.google.com/google-apps/calendar/recurringevents

    public class CalendarExporter {
        public CalendarExporter(CalendarService calendarService, CalendarListEntry calendarEntry) {
            CalendarService = calendarService;
            CalendarEntry = calendarEntry;
        }

        CalendarService CalendarService { get; set; }
        CalendarListEntry CalendarEntry { get; set; }

        public void Export(AppointmentBaseCollection appointments) {
            foreach (var apt in appointments)
                ExportAppointment(apt);
        }

        public void Export(IList<Appointment> appointments) {
            foreach (var apt in appointments)
                ExportAppointment(apt);
        }

        void ExportAppointment(Appointment apt) {
            AppointmentType aptType = apt.Type;
            if (aptType == AppointmentType.Pattern)
            {
                EnsurePatternId(apt);
            }
            else if (aptType != AppointmentType.Normal) { 
                string eventPatternId = EnsurePatternId(apt.RecurrencePattern);
                Debug.Assert(!String.IsNullOrEmpty(eventPatternId));
                if (aptType == AppointmentType.Occurrence)
                    return;
                EventsResource.InstancesRequest instancesRequest = CalendarService.Events.Instances(CalendarEntry.Id, eventPatternId);
                OccurrenceCalculator calculator = OccurrenceCalculator.CreateInstance(apt.RecurrencePattern.RecurrenceInfo);
                instancesRequest.OriginalStart = GoogleCalendarUtils.ConvertEventDateTime(calculator.CalcOccurrenceStartTime(apt.RecurrenceIndex)).DateTimeRaw;
                Events occurrenceEvents = instancesRequest.Execute();
                Debug.Assert(occurrenceEvents.Items.Count == 1);
                Event occurrence = occurrenceEvents.Items[0];
                if (aptType == AppointmentType.ChangedOccurrence) {
                    this.AssignProperties(apt, occurrence);
                } else if (aptType == AppointmentType.DeletedOccurrence) {
                    occurrence.Status = "cancelled";
                }
                Event changedOccurrence = CalendarService.Events.Update(occurrence, CalendarEntry.Id, occurrence.Id).Execute();
                apt.CustomFields["eventId"] = changedOccurrence.Id;
                Log.WriteLine(String.Format("Exported {0} occurrance: {1}, id={2}", (aptType == AppointmentType.ChangedOccurrence) ? "changed" : "deleted", apt.Subject, changedOccurrence.Id));
                return;
            }
            
            Event instance = this.CreateEvent(aptType);
            AssignProperties(apt, instance);
            
            Event result = CalendarService.Events.Insert(instance, CalendarEntry.Id).Execute();
            Log.WriteLine(String.Format("Exported appointment: {0}, id={1}", apt.Subject, result.Id));
        }
        
        private string EnsurePatternId(Appointment pattern) {
            string eventId = pattern.CustomFields["eventId"] as string;
            if (!String.IsNullOrEmpty(eventId))
                return eventId;
            VRecurrenceConverter converter = VRecurrenceConverter.CreateInstance(pattern.RecurrenceInfo.Type);
            DevExpress.XtraScheduler.iCalendar.Internal.IWritable rule = converter.FromRecurrenceInfo(pattern.RecurrenceInfo, pattern) as DevExpress.XtraScheduler.iCalendar.Internal.IWritable;
            string someRule = string.Empty;
            using (MemoryStream stream = new MemoryStream()) {
                using (StreamWriter tw = new StreamWriter(stream)) {
                    iCalendarWriter cw = new iCalendarWriter(tw);
                    rule.WriteToStream(cw);
                }
                stream.Flush();
                someRule = Encoding.ASCII.GetString(stream.ToArray());
            }

            Event instance2 = CreateEvent(pattern.Type);
            AssignProperties(pattern, instance2);
            instance2.Recurrence = new String[] { "RRULE:" + someRule };

            Event patternEvent = CalendarService.Events.Insert(instance2, CalendarEntry.Id).Execute();
            pattern.CustomFields["eventId"] = patternEvent.Id;
            return patternEvent.Id;
        }

        private void AssignProperties(Appointment apt, Event instance) {
            instance.Summary = apt.Subject;
            instance.Description = apt.Description;
            instance.Location = apt.Location;
            
            instance.Start = GoogleCalendarUtils.ConvertEventDateTime(apt.Start);
            instance.End = GoogleCalendarUtils.ConvertEventDateTime(apt.End);
        }

        Event CreateEvent(AppointmentType type) {
            return new Event();
        }
    }
}
