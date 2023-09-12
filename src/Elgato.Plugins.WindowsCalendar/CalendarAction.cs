using BarRaider.SdTools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Svg;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Windows.ApplicationModel.Appointments;

namespace Elgato.Plugins.WindowsCalendar
{
    public static class TaskExtensions
    {
        public static Task DelayUntil(DateTimeOffset date, System.Threading.CancellationToken token = default)
        {
            return Task.Run(async () =>
            {
                token.ThrowIfCancellationRequested();
                while (date.UtcDateTime > DateTime.UtcNow)
                    await Task.Delay(1000, token).ConfigureAwait(false);
            });
        }
    }


    [PluginActionId("elgato.plugins.windowscalendar.calendaraction")]
    public class CalendarAction : PluginBase
    {
        private class PluginSettings
        {
            public static PluginSettings CreateDefaultSettings()
            {
                PluginSettings instance = new PluginSettings();
                instance.OutOfOffice = false;
                instance.AllDay = false;
                instance.Free = false;
                return instance;
            }

            [JsonProperty(PropertyName = "out_of_office")]
            public bool OutOfOffice { get; set; }

            [JsonProperty(PropertyName = "free")]
            public bool Free { get; set; }

            [JsonProperty(PropertyName = "all_day")]
            public bool AllDay { get; set; }
        }

        #region Private Members

        private PluginSettings settings;

        #endregion

        private object eventslock = new object();
        AppointmentStore store;
        private List<Appointment> events = null;
        private Appointment nextAppointment = null;


        public CalendarAction(SDConnection connection, InitialPayload payload) : base(connection, payload)
        {
            // System.Diagnostics.Debugger.Launch();
            if (payload.Settings == null || payload.Settings.Count == 0)
            {
                this.settings = PluginSettings.CreateDefaultSettings();
                SaveSettings();
            }
            else
            {
                this.settings = payload.Settings.ToObject<PluginSettings>();
            }
            _ = LoadCalendar();
        }

        private async Task LoadCalendar()
        {
            store = await Windows.ApplicationModel.Appointments.AppointmentManager.RequestStoreAsync(Windows.ApplicationModel.Appointments.AppointmentStoreAccessType.AllCalendarsReadOnly);
            await LoadNextAppointments();
            store.StoreChanged += (s, e) =>
            {
                _ = LoadNextAppointments();
            };
        }

        private async Task LoadNextAppointments()
        {
            DateTime today = DateTime.Now.Date;
            DateTime end = DateTime.Now.Date.AddMonths(3);

            var ap = await store.FindAppointmentsAsync(today, end - today);
            lock (eventslock)
                events = new List<Appointment>(ap);

            var next = GetNextAppointment();
            if (next.StartTime < DateTimeOffset.Now && // Has started
                next.StartTime + next.Duration < DateTimeOffset.UtcNow + TimeSpan.FromMinutes(Math.Min(15, next.Duration.TotalMinutes / 2)) // Is almost over
              )
            {
                // If this appointment is almost over, find the next one
                next = GetNextAppointment(true);
            }
            next = await store.GetAppointmentInstanceAsync(next.LocalId, next.StartTime); // Ensures all the meeting properties are fully loaded 
            if (next != null)
            {
                ScheduleAlert(next);
            }
        }

        private System.Threading.CancellationTokenSource alertCancellation;

        private void ScheduleAlert(Appointment appointment, bool alert = false)
        {
            nextAppointment = appointment;
            alertCancellation?.Cancel();
            alertCancellation = new System.Threading.CancellationTokenSource();
            var timeToStart = appointment.StartTime - DateTimeOffset.Now;
            if (timeToStart.TotalMinutes > 0)
            {

                if (timeToStart.TotalMinutes < 1)
                {
                    StartCountdown();
                    return;
                }
                else
                {
                    _ = TaskExtensions.DelayUntil(appointment.StartTime.AddMinutes(-1), alertCancellation.Token).
                            ContinueWith(t => { if (!t.IsCanceled) { StartCountdown(); } });
                }
            }
            RefreshIcon();
        }

        private async void RefreshIcon()
        {
            var a = nextAppointment;
            if (a is null)
            {
                await Connection.SetDefaultImageAsync();
                return;
            }
            var now = DateTimeOffset.Now;
            var time = a.StartTime.ToLocalTime();
            string timeString = a.AllDay ? time.ToString("d") : time.ToString();
            if (a.StartTime < now.Date.AddDays(1))
            {
                if (a.AllDay)
                    timeString = "Today";
                else
                    timeString = time.ToString("HH:mm");
            }
            else if (a.StartTime < now.AddDays(7))
            {
                if (a.AllDay)
                    timeString = time.ToString("ddd");
                else
                    timeString = time.ToString("ddd HH:mm");
            }
            if (!a.AllDay && a.Duration.TotalMinutes >= 2)
            {
                var end = time.Add(a.Duration);
                timeString += " - " + end.ToString("HH:mm");
            }
            Color background = Color.Black;
            if (a.StartTime - now <= TimeSpan.Zero)
                background = Color.Green;
            else if (a.StartTime - now < TimeSpan.FromMinutes(1))
                background = now.Second % 2 == 0 ? Color.Red : Color.Black; //Flashes
            else if (a.StartTime - now < TimeSpan.FromMinutes(5))
                background = Color.Red;
            else if (a.StartTime - now < TimeSpan.FromMinutes(10))
                background = Color.Orange;
            else if (a.StartTime - now < TimeSpan.FromMinutes(15))
                background = Color.Yellow;
            else if (a.StartTime - now < TimeSpan.FromHours(1))
                background = Color.CornflowerBlue;
            Color textColor = background == Color.Black ? Color.White : Color.Black;

            var doc = new SvgDocument
            {
                Width = 72,
                Height = 72,
                ViewBox = new SvgViewBox(0, 0, 72, 72),
            };
            doc.Children.Add(new SvgRectangle()
            {
                Fill = new SvgColourServer(background),
                X = 0,
                Y = 0,
                Height = 72,
                Width = 72,
            });
            float y = 30;
            foreach (var word in a.Subject.Split(' '))
            {
                doc.Children.Add(new SvgText(word)
                {
                    FontSize = 10,
                    TextAnchor = SvgTextAnchor.Middle,
                    FontWeight = SvgFontWeight.Bold,
                    Color = new SvgColourServer(textColor),
                    Fill = new SvgColourServer(textColor),
                    X = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 36) },
                    Y = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, y) },
                });
                y += 12;
                if (y > 76) break;
            }

            doc.Children.Add(new SvgText(timeString)
            {
                FontSize = 10,
                TextAnchor = SvgTextAnchor.Start,
                FontWeight = SvgFontWeight.Normal,
                Color = new SvgColourServer(textColor),
                Fill = new SvgColourServer(textColor),
                X = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 2) },
                Y = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 15) },
            });
            string content = null;
            using (MemoryStream ms = new MemoryStream())
            {
                doc.Write(ms);
                ms.Position = 0;
                using (var reader = new StreamReader(ms, System.Text.Encoding.UTF8))
                {
                    content = reader.ReadToEnd();
                }
            }
            await Connection.SetImageAsync($"data:image/svg+xml;charset=utf8,{content}");
        }

        private async void StartCountdown()
        {
            if (nextAppointment is null) return;
            while (nextAppointment.StartTime >= DateTimeOffset.Now)
            {
                RefreshIcon();
                await Task.Delay(250);
            }
            RefreshIcon();
        }

        private Appointment GetNextAppointment(bool notStarted = false)
        {
            Appointment[] nextAppointments = GetNextAppointments(1, notStarted);
            if (nextAppointments.Length > 0)
            {
                return nextAppointments[0];
            }
            return null;
        }

        private Appointment[] GetNextAppointments(int count, bool notStarted = false)
        {
            lock (eventslock)
            {
                if (events != null)
                {
                    IEnumerable<Appointment> candidates = events;
                    candidates = candidates.Where(e => !e.IsCanceledMeeting);
                    if (!settings.OutOfOffice)
                        candidates = candidates.Where(e => e.BusyStatus != AppointmentBusyStatus.OutOfOffice);
                    if (!settings.Free)
                        candidates = candidates.Where(e => e.BusyStatus != AppointmentBusyStatus.Free);
                    if (!settings.AllDay)
                        candidates = candidates.Where(e => !e.AllDay);
                    
                    if (notStarted)
                        candidates = candidates.Where(e => e.StartTime > DateTimeOffset.UtcNow); // Hasn't begun
                    else
                        candidates = candidates.Where(e => e.StartTime + e.Duration >= DateTimeOffset.UtcNow); // Hasn't ended

                    return candidates.OrderBy(e => e.StartTime).Take(count).ToArray();
                }
            }
            return new Appointment[0];
        }

        public override void Dispose()
        {
            Logger.Instance.LogMessage(TracingLevel.INFO, $"Destructor called");
        }

        public override void KeyPressed(KeyPayload payload)
        {
            if (nextAppointment != null)
            {
                // If there's an online link and meeting starts in less than 5 mins, just open that link instead of the appointment details
                var onlineMeetingLink = nextAppointment.OnlineMeetingLink;
                if (onlineMeetingLink != null && (nextAppointment.StartTime - DateTimeOffset.Now < TimeSpan.FromMinutes(5)))
                {
                    _ = Process.Start(onlineMeetingLink);
                }
                else
                {
                    _ = store.ShowAppointmentDetailsAsync(nextAppointment.LocalId);
                }
            }
            Logger.Instance.LogMessage(TracingLevel.INFO, "Key Pressed");
        }

        public override void KeyReleased(KeyPayload payload)
        {
            // Required to implement - no action needed
        }

        public override void ReceivedSettings(ReceivedSettingsPayload payload)
        {
            Tools.AutoPopulateSettings(settings, payload.Settings);
            SaveSettings();
            _ = LoadNextAppointments();
        }

        public override void ReceivedGlobalSettings(ReceivedGlobalSettingsPayload payload) { }

        #region Private Methods

        private Task SaveSettings()
        {
            return Connection.SetSettingsAsync(JObject.FromObject(settings));
        }

        public override void OnTick()
        {
            // Required to implement - no action needed
        }

        #endregion
    }
}