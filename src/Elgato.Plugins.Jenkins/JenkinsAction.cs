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
using System.Net.Http;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Web.Compilation;

namespace Elgato.Plugins.Jenkins
{
    [PluginActionId("elgato.plugins.jenkins.jenkinsaction")]
    public class JenkinsAction : PluginBase
    {
        private class PluginSettings
        {
            public static PluginSettings CreateDefaultSettings()
            {
                PluginSettings instance = new PluginSettings();
                instance.JenkinsUrl = "";
                instance.JobName = "";
                return instance;
            }
            [JsonProperty(PropertyName = "jenkinsUrl")]
            public string JenkinsUrl { get; set; }

            [JsonProperty(PropertyName = "jobName")]
            public string JobName { get; set; }
            [JsonProperty(PropertyName = "username")]
            public string Username { get; set; }
            [JsonProperty(PropertyName = "access_token")]
            public string AccessToken { get; set; }
        }

        #region Private Members

        private PluginSettings settings;

        #endregion

        private object eventslock = new object();


        public JenkinsAction(SDConnection connection, InitialPayload payload) : base(connection, payload)
        {
            //System.Diagnostics.Debugger.Launch();
            if (payload.Settings == null || payload.Settings.Count == 0)
            {
                this.settings = PluginSettings.CreateDefaultSettings();
                SaveSettings();
            }
            else
            {
                this.settings = payload.Settings.ToObject<PluginSettings>();
            }
        }
        [DataContract]
        public class JenkinsJob
        {
            [DataMember]
            public string displayName { get; set; }
            [DataMember]
            public string fullDisplayName { get; set; }
            [DataMember]
            public string fullName { get; set; }
            [DataMember]
            public string name { get; set; }
            [DataMember]
            public string url { get; set; }
            [DataMember]
            public string color { get; set; }
            [DataMember]
            public Build lastBuild { get; set; }
            [DataMember]
            public Build currentBuild { get; set; }
            [DataMember]
            public int nextBuildNumber { get; set; }
            [DataMember]
            public bool inQueue { get; set; }
            
        }
        [DataContract]
        public class JenkinsBuild
        {
            [DataMember]
            public string id { get; set; }
            [DataMember]
            public string displayName { get; set; }
            [DataMember]
            public string fullDisplayName { get; set; }
            [DataMember]
            public int duration { get; set; }
            [DataMember]
            public int estimatedDuration { get; set; }
            [DataMember]
            public long timestamp { get; set; }            
            [DataMember]
            public bool inProgress { get; set; }
            [DataMember]
            public int number { get; set; }
            [DataMember]
            public string result { get; set; }
            [DataMember]
            public string url { get; set; }
            [DataMember]
            public Dictionary<string,object>[] actions { get; set; }
        }
        [DataContract]
        public class Build
        {
            [DataMember]
            public int number { get; set; }
            [DataMember]
            public string url { get; set; }
        }

        int lastBuild = -1;
        private async void Refresh()
        {
            HttpClient client = new HttpClient();
            Dictionary<string, object> badgeAction = null;
            Dictionary<string, object> testResultAction = null;
            JenkinsBuild buildInfo = null;
            JenkinsJob jobInfo = null;
            try
            {
                string jsonUrl = $"{settings.JenkinsUrl?.TrimEnd('/')}/job/{settings.JobName.Replace("/", "/job/")}/api/json";
                if (!string.IsNullOrWhiteSpace(settings.Username) && !string.IsNullOrWhiteSpace(settings.AccessToken))
                    client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("basic", 
                        Convert.ToBase64String(Encoding.UTF8.GetBytes($"{settings.Username}:{settings.AccessToken}")));
                if (Uri.TryCreate(jsonUrl, UriKind.Absolute, out var jsonuri))
                {
                    var jobJson = await client.GetStringAsync(jsonuri);
                    jobInfo = JsonConvert.DeserializeObject<JenkinsJob>(jobJson);
                    if (jobInfo.lastBuild.number == lastBuild) return;

                    if (Uri.TryCreate(jobInfo.lastBuild.url?.TrimEnd('/') + "/api/json", UriKind.Absolute, out var jobUri))
                    {
                        var buildJson = await client.GetStringAsync(jobUri);
                        buildInfo = JsonConvert.DeserializeObject<JenkinsBuild>(buildJson, new JsonSerializerSettings
                        {
                            TypeNameHandling = TypeNameHandling.All,
                            TypeNameAssemblyFormatHandling = TypeNameAssemblyFormatHandling.Simple
                        });

                        badgeAction = buildInfo?.actions.Where(a => a.ContainsKey("_class") && a["_class"] as string == "com.jenkinsci.plugins.badge.action.BadgeAction").FirstOrDefault();
                        testResultAction = buildInfo?.actions.Where(a => a.ContainsKey("_class") && a["_class"] as string == "hudson.tasks.junit.TestResultAction").FirstOrDefault();

                    }
                }
            }
            catch(System.Exception ex)
            {
                return; 
            }
            if (buildInfo is null)
                return;

            string statusText = "";

            Color background = Color.Black;
            Color textColor = Color.White;

            if (buildInfo.inProgress)
            {
                background = Color.Green;
                statusText = "Running...";
            }
            else if (buildInfo.result == "FAILURE")
            {
                background = Color.Red;
                statusText = "FAILED";
            }
            else if (buildInfo.result == "SUCCESS")
            {
                textColor = Color.Green;
                statusText = "Success";
            }
            else if (buildInfo.result == "UNSTABLE")
            {
                textColor = Color.Yellow;
                statusText = "Unstable";
            }
            if(!buildInfo.inProgress && jobInfo.inQueue)
            {
                statusText += "\nQueued...";
                background = Color.Purple;
            }
            if(badgeAction != null && badgeAction.ContainsKey("text"))
            {
                statusText += "\n" + Uri.UnescapeDataString(badgeAction["text"]?.ToString()).Replace("&#43;", "±").Replace(" (","\n(");
            }
            double rectWidth = 72;
            if(buildInfo.inProgress && buildInfo.estimatedDuration > 0 && buildInfo.timestamp > 0)
            {
                var startTime = DateTimeOffset.FromUnixTimeMilliseconds(buildInfo.timestamp);
                var endTime = startTime.AddMilliseconds(buildInfo.estimatedDuration);
                var progress = (DateTimeOffset.UtcNow - startTime).TotalMilliseconds / buildInfo.estimatedDuration * 100;
                rectWidth = 72 * progress / 100;
                statusText += $"\n{Math.Min(99, Math.Round(progress))}";
            }
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
                Width = (float)rectWidth,
            });
            var lines = statusText.Split('\n');
            float y = 72f / 2 - (lines.Length-1) * 12f / 2;
            foreach (var line in lines)
            {
                doc.Children.Add(new SvgText(line)
                {
                    FontSize = 10,
                    TextAnchor = SvgTextAnchor.Middle,
                    FontWeight = SvgFontWeight.Bold,
                    Color = new SvgColourServer(textColor),
                    Fill = new SvgColourServer(textColor),
                    X = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 36) },
                    Y = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, y) },
                }); ;
                y += 12;
            }
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

        public override void Dispose()
        {
            Logger.Instance.LogMessage(TracingLevel.INFO, $"Destructor called");
        }
        public override void KeyPressed(KeyPayload payload)
        {
            string url = $"{settings.JenkinsUrl.TrimEnd('/')}/job/{settings.JobName.Replace("/","/job/")}";
            if(Uri.TryCreate(url, UriKind.Absolute, out var uri))
            {
                var process = new System.Diagnostics.Process();
                process.StartInfo.FileName = "rundll32.exe";
                process.StartInfo.Arguments = "url.dll,FileProtocolHandler " + uri.OriginalString;
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }
        }

        public override void KeyReleased(KeyPayload payload)
        {
            // Required to implement - no action needed
        }

        public override void ReceivedSettings(ReceivedSettingsPayload payload)
        {
            Tools.AutoPopulateSettings(settings, payload.Settings);
            SaveSettings();
            Refresh();
        }

        public override void ReceivedGlobalSettings(ReceivedGlobalSettingsPayload payload) { }

        #region Private Methods

        private Task SaveSettings()
        {
            return Connection.SetSettingsAsync(JObject.FromObject(settings));
        }
        int lastMinuteUpdate;
        public override void OnTick()
        {
            if(DateTime.Now.Minute != lastMinuteUpdate)
            {
                lastMinuteUpdate = DateTime.Now.Minute;
                Refresh();
            }
        }

        #endregion
    }
}