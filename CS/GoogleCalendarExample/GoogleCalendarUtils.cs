using Google.Apis.Calendar.v3.Data;
using System;

namespace GoogleCalendarExample {
    public static class GoogleCalendarUtils {
        public static DateTime ConvertDateTime(EventDateTime start) {
            if (start.DateTime.HasValue)
                return start.DateTime.Value;
            return DateTime.Parse(start.Date);
        }

        public static EventDateTime ConvertEventDateTime(DateTime dateTime) {
            EventDateTime result = new EventDateTime();
            result.DateTime = dateTime;

            result.TimeZone = ToIana(TimeZoneInfo.Local.Id);
            return result;
        }

        static string ToIana(string tzid) {
            if (tzid.Equals("UTC", StringComparison.Ordinal))
                return "Etc/UTC";

            var tzdbSource = NodaTime.TimeZones.TzdbDateTimeZoneSource.Default;
            var tzi = TimeZoneInfo.FindSystemTimeZoneById(tzid);
            if (tzi == null)
                return null;
            var ianaTzid = tzdbSource.MapTimeZoneId(tzi);
            if (ianaTzid == null)
                return null;
            return tzdbSource.CanonicalIdMap[ianaTzid];
        }
    }
}
