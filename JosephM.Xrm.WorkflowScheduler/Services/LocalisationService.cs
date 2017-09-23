using System;

namespace JosephM.Xrm.WorkflowScheduler.Services
{
    public class LocalisationService
    {
        public LocalisationService(LocalisationSettings localisationSettings)
        {
            LocalisationSettings = localisationSettings;
        }

        private LocalisationSettings LocalisationSettings { get; set; }

        public DateTime TargetTodayUniversal
        {
            get
            {
                var localNow = DateTime.Now;
                var targetNow = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(localNow,
                    LocalisationSettings.TargetTimeZoneId);
                var difference = localNow - targetNow;
                return targetNow.Date.Add(difference).ToUniversalTime();
            }
        }

        public DateTime TargetToday
        {
            get { return ConvertToTargetTime(TargetTodayUniversal); }
        }

        public DateTime TodayUnspecifiedType
        {
            get
            {
                var today = TargetToday;
                return new DateTime(today.Year, today.Month, today.Day, 0, 0, 0, DateTimeKind.Unspecified);
            }
        }

        public DateTime ConvertToTargetTime(DateTime dateTime)
        {
            return TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTime, LocalisationSettings.TargetTimeZoneId);
        }

        public DateTime ConvertToTargetDayUtc(DateTime day)
        {
            var sourceDate = day.Date;
            var targetDayTime = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(day,
                LocalisationSettings.TargetTimeZoneId);
            var difference = sourceDate - targetDayTime;
            return sourceDate.Add(difference).ToUniversalTime();
        }

        public DateTime ChangeUtcToLocal(DateTime utcDateTime)
        {
            var convert = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(utcDateTime,
                LocalisationSettings.TargetTimeZoneId);
            var difference = convert - utcDateTime;
            convert = convert.Subtract(difference);
            return convert;
        }

        public string TimeZonename
        {
            get { return LocalisationSettings.TargetTimeZoneId; }
        }
    }
}