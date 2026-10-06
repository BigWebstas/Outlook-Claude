using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;

namespace ClaudeMailSorter
{
    // Stored in %AppData%\ClaudeMailSorter. The API key is encrypted with DPAPI for the current Windows user.
    public class Settings
    {
        public const string DefaultModel = "claude-haiku-4-5";

        static readonly string Dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ClaudeMailSorter");
        static readonly string FilePath = Path.Combine(Dir, "settings.json");

        public string ProtectedApiKey { get; set; } = "";
        public string Model { get; set; } = "";
        public bool AutoProcess { get; set; }
        public bool CreateCategories { get; set; }
        public bool SortOnStartup { get; set; }

        [JsonIgnore]
        public string ApiKey
        {
            get
            {
                if (string.IsNullOrEmpty(ProtectedApiKey)) return "";
                try
                {
                    var bytes = ProtectedData.Unprotect(
                        Convert.FromBase64String(ProtectedApiKey), null, DataProtectionScope.CurrentUser);
                    return Encoding.UTF8.GetString(bytes);
                }
                catch (CryptographicException) { return ""; }
            }
            set
            {
                ProtectedApiKey = string.IsNullOrEmpty(value)
                    ? ""
                    : Convert.ToBase64String(ProtectedData.Protect(
                        Encoding.UTF8.GetBytes(value), null, DataProtectionScope.CurrentUser));
            }
        }

        [JsonIgnore]
        public string EffectiveModel => string.IsNullOrWhiteSpace(Model) ? DefaultModel : Model.Trim();

        public static Settings Load()
        {
            try { return JsonConvert.DeserializeObject<Settings>(File.ReadAllText(FilePath)) ?? new Settings(); }
            catch (IOException) { return new Settings(); }
            catch (JsonException) { return new Settings(); }
        }

        public void Save()
        {
            Directory.CreateDirectory(Dir);
            File.WriteAllText(FilePath, JsonConvert.SerializeObject(this, Formatting.Indented));
        }
    }

    // Remembers which emails were already sorted so the Inbox sweep and new-mail handler skip them.
    public class ProcessedStore
    {
        const int MaxIds = 5000;
        static readonly string FilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ClaudeMailSorter", "processed.json");

        readonly List<string> order;
        readonly HashSet<string> set;

        public ProcessedStore()
        {
            try { order = JsonConvert.DeserializeObject<List<string>>(File.ReadAllText(FilePath)) ?? new List<string>(); }
            catch (IOException) { order = new List<string>(); }
            catch (JsonException) { order = new List<string>(); }
            set = new HashSet<string>(order);
        }

        public bool Contains(string id) => !string.IsNullOrEmpty(id) && set.Contains(id);

        public void Add(string id)
        {
            if (string.IsNullOrEmpty(id)) return;
            if (set.Add(id)) order.Add(id);
            if (order.Count > MaxIds)
            {
                foreach (var old in order.GetRange(0, order.Count - MaxIds)) set.Remove(old);
                order.RemoveRange(0, order.Count - MaxIds);
            }
            File.WriteAllText(FilePath, JsonConvert.SerializeObject(order));
        }
    }
}
