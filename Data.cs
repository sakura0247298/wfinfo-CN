﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.WebSockets;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using WFInfo.Services.WarframeProcess;
using WFInfo.Services.WindowInfo;
using WFInfo.Settings;

namespace WFInfo
{

    class Data
    {
        public JObject marketItems; // Warframe.market item listing           {<id>: "<name>|<url_name>", ...}
        public JObject marketData; // Contains warframe.market ducatonator listing     {<partName>: {"ducats": <ducat_val>,"plat": <plat_val>}, ...}
        public JObject relicData; // Contains relicData from Warframe PC Drops        {<Era>: {"A1":{"vaulted": true,<rare1/uncommon[12]/common[123]>: <part>}, ...}, "Meso": ..., "Neo": ..., "Axi": ..., "Vanguard": ...}
        public JObject equipmentData; // Contains equipmentData from Warframe PC Drops          {<EQMT>: {"vaulted": true, "PARTS": {<NAME>:{"relic_name":<name>|"","count":<num>}, ...}},  ...}
        public JObject nameData; // Contains relic to market name translation          {<relic_name>: <market_name>}

        private static readonly List<Dictionary<int, List<int>>> korean = new List<Dictionary<int, List<int>>>() {
            new Dictionary<int, List<int>>() {
                { 0, new List<int>{ 6, 7, 8, 16 } }, // ㅁ, ㅂ, ㅃ, ㅍ
                { 1, new List<int>{ 2, 3, 4, 16, 5, 9, 10 } }, // ㄴ, ㄷ, ㄸ, ㅌ, ㄹ, ㅅ, ㅆ
                { 2, new List<int>{ 12, 13, 14 } }, // ㅈ, ㅉ, ㅊ
                { 3, new List<int>{ 0, 1, 15, 11, 18 } } // ㄱ, ㄲ, ㅋ, ㅇ, ㅎ
            },
            new Dictionary<int, List<int>>() {
                { 0, new List<int>{ 20, 5, 1, 7, 3, 19 } }, // ㅣ, ㅔ, ㅐ, ㅖ, ㅒ, ㅢ
                { 1, new List<int>{ 16, 11, 15, 10 } }, // ㅟ, ㅚ, ㅞ, ㅙ
                { 2, new List<int>{ 4, 0, 6, 2, 14, 9 } }, // ㅓ, ㅏ, ㅕ, ㅑ, ㅝ, ㅘ
                { 3, new List<int>{ 18, 13, 8, 17, 12 } } // ㅡ, ㅜ, ㅗ, ㅠ, ㅛ
            },
            new Dictionary<int, List<int>>() {
                { 0, new List<int>{ 16, 17, 18, 26 } }, // ㅁ, ㅂ, ㅄ, ㅍ
                { 1, new List<int>{ 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 19, 20, 25 } }, // ㄴ, ㄵ, ㄶ, ㄷ, ㄹ, ㄺ, ㄻ, ㄼ, ㄽ, ㄾ, ㄿ, ㅀ, ㅅ, ㅆ, ㅌ
                { 2, new List<int>{ 22, 23 } }, // ㅈ, ㅊ
                { 3, new List<int>{ 1, 2, 3, 24, 21, 27 } }, // ㄱ, ㄲ, ㄳ, ㅋ, ㅑ, ㅎ
                { 4, new List<int>{ 0 } }, // 
            }
        };

        private readonly string applicationDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\WFInfo";
        private readonly string marketItemsPath;
        private readonly string marketDataPath;
        private readonly string equipmentDataPath;
        private readonly string relicDataPath;
        private readonly string nameDataPath;
        private readonly string filterAllJsonFallbackPath;
        private readonly string sheetJsonFallbackPath;
        private readonly Dictionary<string, string> wfmItemsFallbackPaths;
        public string JWT; // JWT is the security key, store this as email+pw combo'
        private ClientWebSocket marketSocket = new ClientWebSocket();
        private CancellationTokenSource marketSocketCancellation = new CancellationTokenSource();
        private readonly ManualResetEvent marketSocketOpenEvent = new ManualResetEvent(false);
        private TaskCompletionSource<bool> _authenticationCompletionSource;
        private bool _isWebSocketAuthenticated = false;
        private const string filterAllJSON = "https://api.warframestat.us/wfinfo/filtered_items";
        private const string sheetJsonUrl = "https://api.warframestat.us/wfinfo/prices";
        private const string wfmItemsUrl = "https://api.warframe.market/v2/items";
        public string inGameName = string.Empty;
        readonly HttpClient client;
        private string githubVersion;
        public bool rememberMe;
        private LogCapture EElogWatcher;
        private Task autoThread;

        // marketItems lock to ensure avoiding race conditions
        private static readonly object marketItemsLock = new object();

        // Reconnection mechanics for websocket
        // Exponential backoff
        private Timer _reconnectionTimer;
        private volatile bool _intentionalDisconnect = false;
        private volatile bool _reconnectionInProgress = false;
        private int _reconnectionAttempts = 0;
        private readonly int[] _reconnectionDelays = { 1000, 2000, 4000, 8000, 15000, 30000 }; // milliseconds
        private DateTime _lastConnectionTime = DateTime.UtcNow;
        private readonly object _reconnectionLock = new object();
        //

        private readonly IReadOnlyApplicationSettings _settings;
        private readonly IProcessFinder _process;
        private readonly IWindowInfoService _window;

        public static WebClient CreateWfmClient()
        {
            WebClient webClient = CustomEntrypoint.CreateNewWebClient();
            webClient.Headers.Add("platform", "pc");
            webClient.Headers.Add("language", "en");
            return webClient;
        }

        public Data(IReadOnlyApplicationSettings settings, IProcessFinder process, IWindowInfoService window)
        {
            _settings = settings;
            _process = process;
            _window = window;

            Main.AddLog("Initializing Databases");
            marketItemsPath = applicationDirectory + @"\market_items.json";
            marketDataPath = applicationDirectory + @"\market_data.json";
            equipmentDataPath = applicationDirectory + @"\eqmt_data.json";
            relicDataPath = applicationDirectory + @"\relic_data.json";
            nameDataPath = applicationDirectory + @"\name_data.json";
            filterAllJsonFallbackPath = applicationDirectory + @"\fallback_equipment_list.json";
            sheetJsonFallbackPath = applicationDirectory + @"\fallback_price_sheet.json";
            wfmItemsFallbackPaths = new Dictionary<string, string>();
            string[] locales = new string[] { "en", "ko", "zh-hans" };
            foreach (string locale in locales)
            {
                wfmItemsFallbackPaths[locale] = applicationDirectory + @"\fallback_names_" + locale + ".json";
            }

            Directory.CreateDirectory(applicationDirectory);

            // Create websocket for WFM
            WebProxy proxy = null;
            String proxy_string = Environment.GetEnvironmentVariable("http_proxy");
            if (proxy_string != null)
            {
                proxy = new WebProxy(new Uri(proxy_string));
            }
            HttpClientHandler handler = new HttpClientHandler
            {
                Proxy = proxy,
                UseCookies = false,
                CheckCertificateRevocationList = true
            };
            client = new HttpClient(handler);
            client.DefaultRequestHeaders.UserAgent.ParseAdd("WFInfo/" + Main.BuildVersion);
        }

        public void EnableLogCapture()
        {
            if (EElogWatcher == null)
            {
                try
                {
                    EElogWatcher = new LogCapture(_process);
                    EElogWatcher.TextChanged += LogChanged;
                }
                catch (Exception ex)
                {
                    Main.AddLog("Failed to start logcapture, exception: " + ex);
                    Main.StatusUpdate("Failed to start capturing log", 1);
                }
            }
        }

        public void DisableLogCapture()
        {
            if (EElogWatcher != null)
            {
                EElogWatcher.TextChanged -= LogChanged;
                EElogWatcher.Dispose();
                EElogWatcher = null;
            }
        }

        private static void SaveDatabase(string path, object db)
        {
            File.WriteAllText(path, JsonConvert.SerializeObject(db, Formatting.Indented));
        }

        public bool IsJwtLoggedIn()
        {
            return JWT != null && JWT.Length > 300; //check if the token is of the right length
        }

        public int GetGithubVersion()
        {
            WebClient githubWebClient = CustomEntrypoint.CreateNewWebClient();
            JObject github =
                JsonConvert.DeserializeObject<JObject>(
                    githubWebClient.DownloadString("https://api.github.com/repos/sakura0247298/wfinfo-CN/releases/latest"));
            if (github.ContainsKey("tag_name"))
            {
                githubVersion = github["tag_name"]?.ToObject<string>();
                return Main.VersionToInteger(githubVersion);
            }
            return Main.VersionToInteger(Main.BuildVersion);
        }

        // Load item list from Sheets
        public async Task<bool> ReloadItems()
        {
            var enItems = await GetWfmItemList("en");
            var localizedItems = _settings.Locale == "en" ? enItems : await GetWfmItemList(_settings.Locale);

            JObject tempMarketItems = new JObject();
            JArray items = JArray.FromObject(enItems.Data["data"]);

            int primeCount = 0;
            int totalCount = 0;

            foreach (var item in items)
            {
                totalCount++;
                string name = item["i18n"]["en"]["name"].ToString();

                // STRICT prime filtering - must contain " Prime" (with space)
                if (name.Contains(" Prime") && !name.Contains(" Set"))
                {
                    if ((name.Contains("Neuroptics") || name.Contains("Chassis") ||
                            name.Contains("Systems") || name.Contains("Harness") ||
                            name.Contains("Wings")))
                    {
                        name = name.Replace(" Blueprint", "");
                    }

                    tempMarketItems[item["id"].ToString()] = name + "|" + item["slug"];
                    primeCount++;
                    //Main.AddLog($"Added Prime item: {name}");
                }
            }

            items = JArray.FromObject(localizedItems.Data["data"]);
            foreach (var item in items)
            {
                string name = item["slug"].ToString();
                if (name.Contains("prime") && tempMarketItems.ContainsKey(item["id"].ToString()))
                {
                    try
                    {
                        // 尝试使用当前语言代码
                        string localeKey = _settings.Locale;
                        // 处理中文语言代码映射
                        if ((_settings.Locale == "zh-hans" || _settings.Locale == "zh-CN") && item["i18n"].ToObject<JObject>().ContainsKey("zh"))
                        {
                            localeKey = "zh";
                        }
                        
                        if (item["i18n"].ToObject<JObject>().ContainsKey(localeKey))
                        {
                            tempMarketItems[item["id"].ToString()] = tempMarketItems[item["id"].ToString()] + "|" + item["i18n"][localeKey]["name"];
                        }
                    }
                    catch (Exception ex)
                    {
                        Main.AddLog("Error processing localized item name: " + ex.Message);
                    }
                }
            }

            // Atomically replace marketItems under lock
            lock (marketItemsLock)
            {
                marketItems = tempMarketItems;
            }

            Main.AddLog("Item database has been downloaded");
            return enItems.IsFallback || localizedItems.IsFallback;
        }

        // Load market data from Sheets
        private JObject LoadMarket(JObject allFiltered, JArray sheetData)
        {
            // Initialize market data
            var newMarketData = new JObject();

            foreach (var item in sheetData)
            {
                var key = item["name"].ToString();
                var transformedItem = new JObject
                {
                    ["name"] = item["name"],
                    ["plat"] = item["custom_avg"], // Map custom_avg → plat
                    ["volume"] = item["today_vol"],
                    ["ducats"] = 0 // Will be filled by LoadEqmtData
                };

                newMarketData[key] = transformedItem;

                // Add a "Blueprint"-stripped alias
                var alias = key.Replace(" Blueprint", "");
                if (!string.Equals(alias, key, StringComparison.Ordinal)
                    && !newMarketData.TryGetValue(alias, out _))
                {
                    newMarketData[alias] = transformedItem;
                }
            }

            // Load ignored items
            foreach (KeyValuePair<string, JToken> ignored in allFiltered["ignored_items"].ToObject<JObject>())
            {
                newMarketData[ignored.Key] = ignored.Value;
            }

            Main.AddLog("Plat database has been downloaded");

            return newMarketData;
        }

        private async Task<JObject> LoadMarketItem(string url)
        {
            

            JObject stats = new JObject
                {
                    { "avg_price", 999 },
                    { "volume", 0 }
                };

            try
            {
                await Task.Delay(333);
                string statsResponse = await client.GetStringAsync("https://api.warframe.market/v1/items/" + url + "/statistics");
                JObject allStats = JsonConvert.DeserializeObject<JObject>(statsResponse);
                JToken latestStats = allStats["payload"]["statistics_closed"]["90days"].LastOrDefault();
                if (latestStats != null)
                {
                    stats = latestStats.ToObject<JObject>();
                } 
                else
                {
                    Main.AddLog("Using placeholder stats");
                }
            }
            catch (Exception ex)
            {
                Main.AddLog("Failed to fetch stats " + Environment.NewLine + ex.ToString());
            }

            string ducat = "0";
            try
            {
                await Task.Delay(333);
                string itemResponse = await client.GetStringAsync("https://api.warframe.market/v2/item/" + url);
                JObject responseJObject = JsonConvert.DeserializeObject<JObject>(itemResponse);
                if (responseJObject["data"].ToObject<JObject>().TryGetValue("ducats", out JToken temp))
                {
                    ducat = temp.ToObject<string>();
                }
                else
                {
                    Main.AddLog("Using placeholder ducats ");
                }
            }
            catch (Exception ex)
            {
                Main.AddLog("Failed to fetch ducats " + Environment.NewLine + ex.ToString());
            }


            return new JObject
            {
                { "ducats", ducat },
                { "plat", stats["avg_price"] },
                { "volume", stats["volume"] }
            };
        }

        private (JObject RelicData, JObject NameData) LoadEqmtData(JObject allFiltered, JObject mrktData, JObject eqmtData)
        {
            // fill in equipmentData (NO OVERWRITE)
            // fill in nameData
            // fill in relicData

            var newRelicData = new JObject();
            var newNameData = new JObject();

            foreach (KeyValuePair<string, JToken> era in allFiltered["relics"].ToObject<JObject>())
            {
                newRelicData[era.Key] = new JObject();
                foreach (KeyValuePair<string, JToken> relic in era.Value.ToObject<JObject>())
                    newRelicData[era.Key][relic.Key] = relic.Value;
            }

            foreach (KeyValuePair<string, JToken> prime in allFiltered["eqmt"].ToObject<JObject>())
            {
                string primeName = prime.Key.Substring(0, prime.Key.IndexOf("Prime") + 5);
                if (!eqmtData.TryGetValue(primeName, out _))
                    eqmtData[primeName] = new JObject();
                eqmtData[primeName]["vaulted"] = prime.Value["vaulted"];
                eqmtData[primeName]["type"] = prime.Value["type"];
                if (!eqmtData[primeName].ToObject<JObject>().TryGetValue("mastered", out _))
                    eqmtData[primeName]["mastered"] = false;

                if (!eqmtData[primeName].ToObject<JObject>().TryGetValue("parts", out _))
                    eqmtData[primeName]["parts"] = new JObject();


                foreach (KeyValuePair<string, JToken> part in prime.Value["parts"].ToObject<JObject>())
                {
                    string partName = part.Key;
                    if (!eqmtData[primeName]["parts"].ToObject<JObject>().TryGetValue(partName, out _))
                        eqmtData[primeName]["parts"][partName] = new JObject();
                    if (!eqmtData[primeName]["parts"][partName].ToObject<JObject>().TryGetValue("owned", out _))
                        eqmtData[primeName]["parts"][partName]["owned"] = 0;
                    eqmtData[primeName]["parts"][partName]["vaulted"] = part.Value["vaulted"];
                    eqmtData[primeName]["parts"][partName]["count"] = part.Value["count"];
                    eqmtData[primeName]["parts"][partName]["ducats"] = part.Value["ducats"];


                    if (part.Key != null && prime.Value?["type"] != null && part.Value?["ducats"] != null)
                    {
                        string gameName = part.Key;
                        string partType = prime.Value["type"].ToString();

                        if (partType == "Archwing" && (part.Key.Contains("Systems") || part.Key.Contains("Harness") || part.Key.Contains("Wings")))
                        {
                            gameName += " Blueprint";
                        }
                        else if (partType == "Warframes" && (part.Key.Contains("Systems") || part.Key.Contains("Neuroptics") || part.Key.Contains("Chassis")))
                        {
                            gameName += " Blueprint";
                        }

                        string targetKey = null;
                        if (mrktData.TryGetValue(partName, out _))
                            targetKey = partName;
                        else if (mrktData.TryGetValue(partName + " Blueprint", out _))
                            targetKey = partName + " Blueprint";

                        if (targetKey != null)
                        {
                            newNameData[gameName] = partName;
                            mrktData[targetKey]["ducats"] = Convert.ToInt32(part.Value["ducats"].ToString(), Main.culture);
                        }
                    }
                }
            }

            // Add default values for ignored items
            foreach (KeyValuePair<string, JToken> ignored in allFiltered["ignored_items"].ToObject<JObject>())
            {
                newNameData[ignored.Key] = ignored.Key;
            }

            Main.AddLog("Prime Database has been downloaded");
            return (newRelicData, newNameData);
        }

        private async Task<(JObject Data, bool IsFallback)> GetWfmItemList(string locale)
        {
            try
            {
                using (var request = new HttpRequestMessage()
                {
                    RequestUri = new Uri(wfmItemsUrl),
                    Method = HttpMethod.Get
                })
                {
                    request.Headers.Add("language", locale);
                    request.Headers.Add("accept", "application/json");
                    request.Headers.Add("platform", "pc");
                    await Task.Delay(333);
                    var response = await client.SendAsync(request).ConfigureAwait(false);
                    var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    var data = JsonConvert.DeserializeObject<JObject>(body);
                    if (wfmItemsFallbackPaths.TryGetValue(locale, out var fallbackPath)) 
                    {
                        File.WriteAllText(fallbackPath, body);
                    }
                    return (data, false);
                }
            }
            catch (Exception ex)
            {
                if (wfmItemsFallbackPaths.TryGetValue(locale, out var fallbackPath))
                {
                    Main.AddLog("Failed to fetch/parse " + wfmItemsUrl + ", using file " + fallbackPath + Environment.NewLine + ex.ToString());
                    if (File.Exists(fallbackPath))
                    {
                        string response = File.ReadAllText(fallbackPath);
                        JObject data = JsonConvert.DeserializeObject<JObject>(response);
                        return (data, true);
                    }
                }
                else
                {
                    Main.AddLog("Failed to fetch/parse " + wfmItemsUrl + ", and no fallback path found for locale: " + locale + Environment.NewLine + ex.ToString());
                }
                throw new AggregateException("No local fallback found", ex);
            }
        }

        private async Task<(JObject Data, bool IsFallback)> GetAllFiltered()
        {
            try
            {
                string response = await client.GetStringAsync(filterAllJSON);
                JObject data = JsonConvert.DeserializeObject<JObject>(response);
                File.WriteAllText(filterAllJsonFallbackPath, response);
                return (data, false);
            }
            catch (Exception ex)
            {
                Main.AddLog("Failed to fetch/parse " + filterAllJSON + ", using file " + filterAllJsonFallbackPath + Environment.NewLine + ex.ToString());
                if (File.Exists(filterAllJsonFallbackPath))
                {
                    string response = File.ReadAllText(filterAllJsonFallbackPath);
                    JObject data = JsonConvert.DeserializeObject<JObject>(response);
                    return (data, true);
                }
                else
                {
                    throw new AggregateException("No local fallback found", ex);
                }
            }
            
        }

        private async Task<(JArray Data, bool IsFallback)> GetSheetData()
        {
            try
            {
                string response = await client.GetStringAsync(sheetJsonUrl);
                JArray data = JsonConvert.DeserializeObject<JArray>(response);
                File.WriteAllText(sheetJsonFallbackPath, response);
                return (data, false);
            }
            catch (Exception ex)
            {
                Main.AddLog("Failed to fetch/parse " + sheetJsonUrl + ", using file " + sheetJsonFallbackPath + Environment.NewLine + ex.ToString());
                if (File.Exists(sheetJsonFallbackPath))
                {
                    string response = File.ReadAllText(sheetJsonFallbackPath);
                    JArray data = JsonConvert.DeserializeObject<JArray>(response);
                    return (data, true);
                }
                else
                {
                    throw new AggregateException("No local fallback found", ex);
                }
            }

        }

        private SemaphoreSlim _DataUpdateSema = new SemaphoreSlim(1);

        public async Task Update()
        {
            await _DataUpdateSema.WaitAsync();
            try
            {
                await UpdateInner(false);
            }
            finally
            {
                _DataUpdateSema.Release();
            }
        }

        public async Task ForceDataUpdate()
        {
            var acquired = await _DataUpdateSema.WaitAsync(TimeSpan.Zero);
            if (!acquired)
            {
                Main.AddLog("Data Update already in progress");
                Main.StatusUpdate("Data Update already in progress", 3);
                Main.RunOnUIThread(() =>
                {
                    MainWindow.INSTANCE.ReloadMarket.IsEnabled = true;
                });
                return;
            }

            try
            {
                await UpdateInner(true);
                Main.RunOnUIThread(() =>
                {
                    MainWindow.INSTANCE.ReloadMarket.IsEnabled = true;
                });
            }
            catch (Exception ex)
            {
                Main.AddLog( nameof(ForceDataUpdate)+ " FAILED " + ex);
                Main.StatusUpdate("Data Update Failed", 0);
                Main.RunOnUIThread(() =>
                {
                    Main.SpawnErrorPopup(DateTime.Now, 0);
                    MainWindow.INSTANCE.ReloadMarket.IsEnabled = true;
                });
            }
            finally
            {
                _DataUpdateSema.Release();
            }
        }

        private JObject ParseFileOrMakeNew(string path, ref bool parseHasFailed)
        {
            if (File.Exists(path))
            {
                return JsonConvert.DeserializeObject<JObject>(File.ReadAllText(path));
            }
            Main.AddLog(path + " missing, loading blank");
            parseHasFailed = true;
            return new JObject();
        }


        public async Task UpdateInner(bool force)
        {
            Main.AddLog("Starting UpdateInner, force: " + force);
            DateTime now = DateTime.Now;

            bool parseHasFailed = false;

            // Init core data objects, if necessary
            if (marketData == null)
            {
                marketData = ParseFileOrMakeNew(marketDataPath, ref parseHasFailed);
            }
            lock (marketItemsLock)
            {
                if (marketItems == null)
                {
                    marketItems = ParseFileOrMakeNew(marketItemsPath, ref parseHasFailed);
                }
            }
            if (equipmentData == null)
            {
                equipmentData = ParseFileOrMakeNew(equipmentDataPath, ref parseHasFailed);
            }
            if (relicData == null)
            {
                relicData = ParseFileOrMakeNew(relicDataPath, ref parseHasFailed);
            }
            if (nameData == null)
            {
                nameData = ParseFileOrMakeNew(nameDataPath, ref parseHasFailed);
            }

            string oldMarketTimeText;
            bool marketIsRecent = false;
            if (marketData.TryGetValue("version", out _) && (marketData["version"].ToObject<string>() == Main.BuildVersion)
                && marketData.TryGetValue("timestamp", out var timestamp) && timestamp.ToObject<DateTime>() > now.AddHours(-12))
            {
                // market data confirmed to be updated less than 12 hours ago. Actual data age can vary, due to pipeline delays
                marketIsRecent = true;
                oldMarketTimeText = timestamp.ToObject<DateTime>().ToString("MMM dd - HH:mm", Main.culture);
            }
            else
            {
                oldMarketTimeText = "UNKNOWN";
            }

                string oldEquipmentTimeText;
            bool equipmentIsRecent = false;
            if (equipmentData.TryGetValue("timestamp", out var equipmentTimestamp) && equipmentTimestamp.ToObject<DateTime>() > now.AddHours(-12))
            {
                // equipment data confirmed to be updated less than 12 hours ago. Actual data age can vary, due to pipeline delays
                equipmentIsRecent = true;
                oldEquipmentTimeText = equipmentTimestamp.ToObject<DateTime>().ToString("MMM dd - HH:mm", Main.culture);
            }
            else
            {
                oldEquipmentTimeText = "UNKNOWN";
            }

            if (!parseHasFailed && !force && marketIsRecent && equipmentIsRecent)
            {
                Main.RunOnUIThread(() =>
                {
                    MainWindow.INSTANCE.MarketData.Content = oldMarketTimeText;
                    MainWindow.INSTANCE.DropData.Content = oldEquipmentTimeText;
                });
                return;
            }

            var allFiltered = await GetAllFiltered();
            var sheetData = await GetSheetData();

            var marketItemsIsFallback = await ReloadItems();

            var newMarketData = LoadMarket(allFiltered.Data, sheetData.Data);

            // check for any items reported by WFM name table, but missing from LoadMarket results
            var missing = new List<(string Name, string Url)>();
            lock (marketItemsLock)
            {
                foreach (KeyValuePair<string, JToken> elem in marketItems)
                {
                    if (elem.Key == "version") continue;
                    string[] split = elem.Value.ToString().Split('|');
                    if (split.Length < 2) continue;
                    string itemName = split[0];
                    string itemUrl = split[1];
                    if (!itemName.Contains(" Set"))
                    {
                        // Try direct lookup first, then try with " Blueprint" appended
                        if (!newMarketData.ContainsKey(itemName) &&
                            !newMarketData.ContainsKey(itemName + " Blueprint"))
                        {
                            missing.Add((itemName, itemUrl));
                        }
                    }
                }
            }
            // retrieve missing item data directly from WFM
            foreach (var m in missing)
            {
                Main.AddLog("Load missing market item: " + m.Name);
                newMarketData[m.Name] = await LoadMarketItem(m.Url);
            }

            // to preserve owned count and mastery status while being cautious about thread safety, make copy of equipment data to update
            var newEquipmentData = (JObject)equipmentData.DeepClone();

            // get/update remaining info
            var (newRelicData, newNameData) = LoadEqmtData(allFiltered.Data, newMarketData, newEquipmentData);


            string marketTimeText;
            string equipmentTimeText;
            // Skip writing timestamp if fallback data files were relied on
            if (!allFiltered.IsFallback && !sheetData.IsFallback && !marketItemsIsFallback)
            {
                newMarketData["timestamp"] = now;
                marketTimeText = now.ToString("MMM dd - HH:mm", Main.culture);
            }
            else
            {
                marketTimeText = "FALLBACK";
            }

            if (!allFiltered.IsFallback)
            {
                newEquipmentData["timestamp"] = now;
                equipmentTimeText = now.ToString("MMM dd - HH:mm", Main.culture);
            }
            else
            {
                equipmentTimeText = "FALLBACK";
            }

            newMarketData["version"] = Main.BuildVersion;

            // swap to new data files. marketItems excluded because ReloadItems does that immediately
            marketData = newMarketData;
            equipmentData = newEquipmentData;
            relicData = newRelicData;
            nameData = newNameData;

            SaveAllJSONs();

            Main.RunOnUIThread(() => 
            {
                MainWindow.INSTANCE.MarketData.Content = marketTimeText;
                MainWindow.INSTANCE.DropData.Content = equipmentTimeText;
            });

            Main.AddLog("Data Update Complete");
            Main.StatusUpdate("Data Update Complete", 0);
        }

        public void SaveAllJSONs()
        {
            SaveDatabase(equipmentDataPath, equipmentData);
            SaveDatabase(relicDataPath, relicData);
            SaveDatabase(nameDataPath, nameData);
            SaveDatabase(marketItemsPath, marketItems);
            SaveDatabase(marketDataPath, marketData);
            
            // 更新OCR字典（增量更新）
            UpdateOcrDictionary();
        }
        
        /// <summary>
        /// 更新OCR字典，增量添加中文字符
        /// </summary>
        public void UpdateOcrDictionary()
        {
            try
            {
                string keysPath = Path.Combine(applicationDirectory, "models", "v4", "ppocr_keys_v1.txt");
                
                // 1. 读取现有字符到HashSet（避免重复）
                var existingChars = new HashSet<string>();
                if (File.Exists(keysPath))
                {
                    foreach (var line in File.ReadAllLines(keysPath))
                    {
                        var trimmed = line.Trim();
                        if (trimmed.Length == 1)
                        {
                            existingChars.Add(trimmed);
                        }
                    }
                }
                
                // 2. 从market_items.json提取中文字符
                string marketItemsPath = Path.Combine(applicationDirectory, "market_items.json");
                if (!File.Exists(marketItemsPath))
                {
                    Main.AddLog("market_items.json不存在，跳过字典更新");
                    return;
                }
                
                var json = File.ReadAllText(marketItemsPath);
                var marketItems = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
                
                int newCharCount = 0;
                foreach (var item in marketItems)
                {
                    if (item.Key == "version") continue;
                    var parts = item.Value.Split('|');
                    if (parts.Length >= 3)
                    {
                        var chineseName = parts[2];
                        foreach (var c in chineseName)
                        {
                            // 只处理中文字符
                            if (c >= 0x4E00 && c <= 0x9FA5)
                            {
                                var charStr = c.ToString();
                                // 增量更新：只添加不存在的字符
                                if (!existingChars.Contains(charStr))
                                {
                                    existingChars.Add(charStr);
                                    newCharCount++;
                                }
                            }
                        }
                    }
                }
                
                // 3. 只有新增字符时才写入文件
                if (newCharCount > 0)
                {
                    using (var writer = new StreamWriter(keysPath, false, System.Text.Encoding.UTF8))
                    {
                        foreach (var c in existingChars)
                        {
                            writer.WriteLine(c);
                        }
                    }
                    Main.AddLog($"字典已更新，新增{newCharCount}个中文字符，总计{existingChars.Count}个字符");
                }
                else
                {
                    Main.AddLog($"字典无需更新，当前共{existingChars.Count}个字符");
                }
            }
            catch (Exception ex)
            {
                Main.AddLog($"更新字典失败: {ex.Message}");
            }
        }

        public bool IsPartVaulted(string name)
        {
            if (name.IndexOf("Prime") < 0)
                return false;
            string eqmt = name.Substring(0, name.IndexOf("Prime") + 5);
            return equipmentData[eqmt]["parts"][name]["vaulted"].ToObject<bool>();
        }

        public bool IsPartMastered(string name)
        {
            if (name.IndexOf("Prime") < 0)
                return false;
            string eqmt = name.Substring(0, name.IndexOf("Prime") + 5);
            return equipmentData[eqmt]["mastered"].ToObject<bool>();
        }

        public string PartsOwned(string name)
        {
            if (name.IndexOf("Prime") < 0)
                return "0";
            string eqmt = name.Substring(0, name.IndexOf("Prime") + 5);
            string owned = equipmentData[eqmt]["parts"][name]["owned"].ToString();
            if (owned == "0")
                return "0";
            return owned;
        }

        public string PartsCount(string name)
        {
            if (name.IndexOf("Prime") < 0)
                return "0";
            string eqmt = name.Substring(0, name.IndexOf("Prime") + 5);
            string count = equipmentData[eqmt]["parts"][name]["count"].ToString();
            if (count == "0")
                return "0";
            return count;
        }

        private static void AddElement(int[,] d, List<int> xList, List<int> yList, int x, int y)
        {
            int loc = 0;
            int temp = d[x, y];
            while (loc < xList.Count && temp > d[xList[loc], yList[loc]])
            {
                loc += 1;
            }

            if (loc == xList.Count)
            {
                xList.Add(x);
                yList.Add(y);
                return;
            }

            xList.Insert(loc, x);
            yList.Insert(loc, y);
        }

        readonly char[,] ReplacementList = null;

        public int GetDifference(char c1, char c2)
        {
            if (c1 == c2 || c1 == '?' || c2 == '?')
            {
                return 0;
            }

            for (int i = 0; i < ReplacementList.GetLength(0) - 1; i++)
            {
                if ((c1 == ReplacementList[i, 0] || c2 == ReplacementList[i, 0]) &&
                    (c1 == ReplacementList[i, 1] || c2 == ReplacementList[i, 1]))
                {
                    return 0;
                }
            }

            return 1;
        }

        public int LevenshteinDistance(string s, string t)
        {
            // 检查字符串是否包含中文
            bool containsChinese = System.Text.RegularExpressions.Regex.IsMatch(s, @"[\u4e00-\u9fa5]") || System.Text.RegularExpressions.Regex.IsMatch(t, @"[\u4e00-\u9fa5]");
            
            // 检查字符串是否包含韩文
            bool containsKorean = System.Text.RegularExpressions.Regex.IsMatch(s, @"[\uac00-\ud7af]") || System.Text.RegularExpressions.Regex.IsMatch(t, @"[\uac00-\ud7af]");
            
            if (containsKorean)
            {
                // for korean
                return LevenshteinDistanceKorean(s, t);
            }
            else if (containsChinese)
            {
                // for chinese
                return LevenshteinDistanceChinese(s, t);
            }
            else
            {
                return LevenshteinDistanceDefault(s, t);
            }
        }

        public static int LevenshteinDistanceDefault(string s, string t)
        {
            // Levenshtein Distance determines how many character changes it takes to form a known result
            // For example: Nuvo Prime is closer to Nova Prime (2) then Ash Prime (4)
            // For more info see: https://en.wikipedia.org/wiki/Levenshtein_distance
            s = s.ToLower(Main.culture);
            t = t.ToLower(Main.culture);
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            if (n == 0 || m == 0)
                return n + m;

            d[0, 0] = 0;

            int count = 0;
            for (int i = 1; i <= n; i++)
                d[i, 0] = (s[i - 1] == ' ' ? count : ++count);

            count = 0;
            for (int j = 1; j <= m; j++)
                d[0, j] = (t[j - 1] == ' ' ? count : ++count);

            for (int i = 1; i <= n; i++)
                for (int j = 1; j <= m; j++)
                {
                    // deletion of s
                    int opt1 = d[i - 1, j];
                    if (s[i - 1] != ' ')
                        opt1++;

                    // deletion of t
                    int opt2 = d[i, j - 1];
                    if (t[j - 1] != ' ')
                        opt2++;

                    // swapping s to t
                    int opt3 = d[i - 1, j - 1];
                    if (t[j - 1] != s[i - 1])
                        opt3++;
                    d[i, j] = Math.Min(Math.Min(opt1, opt2), opt3);
                }



            return d[n, m];
        }

        // This isn't used anymore?!
        public static bool IsKorean(String str)
        {
            // Safeguard for empty strings that will give false positives and/or crashes
            if (string.IsNullOrEmpty(str)) return false;
            char c = str[0];
            if (0x1100 <= c && c <= 0x11FF) return true;
            if (0x3130 <= c && c <= 0x318F) return true;
            if (0xAC00 <= c && c <= 0xD7A3) return true;
            return false;
        }

        public string GetLocaleNameData(string s)
        {
            string localeName = "";

            lock (marketItemsLock)
            {
                if (marketItems != null) // Add null check
                {
                    foreach (var marketItem in marketItems)
                    {
                        if (marketItem.Key == "version")
                            continue;
                        string[] split = marketItem.Value.ToString().Split('|');
                        if (split[0] == s)
                        {
                            localeName = split.Length > 2 ? split[2] : "";
                            break;
                        }
                    }
                    
                    // 如果没有找到精确匹配，尝试从部件名称中提取
                    if (string.IsNullOrEmpty(localeName) && s.Contains("Prime"))
                    {
                        foreach (var marketItem in marketItems)
                        {
                            if (marketItem.Key == "version")
                                continue;
                            string[] split = marketItem.Value.ToString().Split('|');
                            if (split[0].StartsWith(s + " ") || split[0] == s)
                            {
                                if (split.Length > 2)
                                {
                                    localeName = split[2];
                                    string[] suffixes = { " 蓝图", " 枪管", " 枪机", " 枪托", " 机体", " 系统", " 头部神经光元", " 外甲", " 刀刃", " 握柄", " 架子", " 链条", " 弓弦", " 上弓臂", " 下弓臂", " 镖袋", " 外壳", " 伪装", " Barrel", " Receiver", " Stock", " Chassis", " Systems", " Neuroptics", " Harness", " Blade", " Grip", " String", " Upper Limb", " Lower Limb", " Pouch", " Carapace", " Blueprint" };
                                    foreach (var suffix in suffixes)
                                    {
                                        if (localeName.EndsWith(suffix))
                                        {
                                            localeName = localeName.Substring(0, localeName.Length - suffix.Length);
                                            break;
                                        }
                                    }
                                }
                                break;
                            }
                        }
                    }
                }
            }

            return localeName;
        }
        
        public string GetEnglishNameFromLocale(string localeName)
        {
            string englishName = "";
            
            lock (marketItemsLock)
            {
                if (marketItems != null)
                {
                    foreach (var marketItem in marketItems)
                    {
                        if (marketItem.Key == "version")
                            continue;
                        string[] split = marketItem.Value.ToString().Split('|');
                        if (split.Length > 2 && split[2] == localeName)
                        {
                            englishName = split[0];
                            break;
                        }
                    }
                }
            }
            
            return englishName;
        }
        
        public string FindBestEnglishNameFromChinese(string chineseName)
        {
            string bestMatch = "";
            int bestDistance = 9999;
            
            // 标准化中文名称，移除空格
            string normalizedChinese = chineseName.Replace(" ", "").Replace("Prime", " Prime");
            
            lock (marketItemsLock)
            {
                if (marketItems != null)
                {
                    foreach (var marketItem in marketItems)
                    {
                        if (marketItem.Key == "version")
                            continue;
                        string[] split = marketItem.Value.ToString().Split('|');
                        if (split.Length > 2)
                        {
                            string localeName = split[2];
                            // 标准化数据库中的中文名称
                            string normalizedLocale = localeName.Replace(" ", "");
                            
                            int distance = LevenshteinDistance(normalizedLocale, normalizedChinese);
                            if (distance < bestDistance)
                            {
                                bestDistance = distance;
                                bestMatch = split[0];
                            }
                            
                            // 也尝试用原始名称匹配
                            distance = LevenshteinDistance(localeName, chineseName);
                            if (distance < bestDistance)
                            {
                                bestDistance = distance;
                                bestMatch = split[0];
                            }
                        }
                    }
                }
            }
            
            if (bestDistance <= 10)
            {
                Main.AddLog($"中文匹配: \"{chineseName}\" -> \"{bestMatch}\" (距离: {bestDistance})");
                return bestMatch;
            }
            
            return "";
        }
        private protected static string e = "A?s/,;j_<Z3Q4z&)";

        public int LevenshteinDistanceKorean(string s, string t)
        {
            // NameData s 를 한글명으로 가져옴
            s = GetLocaleNameData(s);

            // i18n korean edit distance algorithm
            s = " " + s.Replace("설계도", "").Replace(" ", "");
            t = " " + t.Replace("설계도", "").Replace(" ", "");

            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            if (n == 0 || m == 0)
                return n + m;
            int i, j;

            for (i = 1; i < s.Length; i++) d[i, 0] = i * 9;
            for (j = 1; j < t.Length; j++) d[0, j] = j * 9;

            int s1, s2;

            for (i = 1; i < s.Length; i++)
            {
                for (j = 1; j < t.Length; j++)
                {
                    s1 = 0;
                    s2 = 0;

                    char cha = s[i];
                    char chb = t[j];
                    int[] a = new int[3];
                    int[] b = new int[3];
                    a[0] = (((cha - 0xAC00) - (cha - 0xAC00) % 28) / 28) / 21;
                    a[1] = (((cha - 0xAC00) - (cha - 0xAC00) % 28) / 28) % 21;
                    a[2] = (cha - 0xAC00) % 28;

                    b[0] = (((chb - 0xAC00) - (chb - 0xAC00) % 28) / 28) / 21;
                    b[1] = (((chb - 0xAC00) - (chb - 0xAC00) % 28) / 28) % 21;
                    b[2] = (chb - 0xAC00) % 28;

                    if (a[0] != b[0] && a[1] != b[1] && a[2] != b[2])
                    {
                        s1 = 9;
                    }
                    else
                    {
                        for (int k = 0; k < 3; k++)
                        {
                            if (a[k] != b[k])
                            {
                                if (GroupEquals(korean[k], a[k], b[k]))
                                {
                                    s2 += 1;
                                }
                                else
                                {
                                    s1 += 1;
                                }
                            }

                        }
                        s1 *= 3;
                        s2 *= 2;
                    }

                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 9, d[i, j - 1] + 9), d[i - 1, j - 1] + s1 + s2);
                }
            }

            return d[s.Length - 1, t.Length - 1];
        }

        private static readonly List<HashSet<char>> similarChineseCharacters = new List<HashSet<char>>() {
            new HashSet<char>{ '龙', '尤', '尢' },          // 龙形字
            new HashSet<char>{ '骑', '奇', '琦' },          // 骑形字
            new HashSet<char>{ '枪', '抢', '呛' },          // 枪形字
            new HashSet<char>{ '炮', '泡', '跑' },          // 炮形字
            new HashSet<char>{ '剑', '刽', '刿' },          // 剑形字
            new HashSet<char>{ '刀', '刁', '力' },          // 刀形字
            new HashSet<char>{ '甲', '申', '由' },          // 甲形字
            new HashSet<char>{ '乙', '己', '已' },          // 乙形字
            new HashSet<char>{ '天', '夫', '夭' },          // 天形字
            new HashSet<char>{ '日', '曰', '目' },          // 日形字
            new HashSet<char>{ '木', '术', '末' },          // 木形字
            new HashSet<char>{ '水', '永', '氷' },          // 水形字
            new HashSet<char>{ '火', '灭', '灰' },          // 火形字
            new HashSet<char>{ '土', '士', '干' },          // 土形字
            new HashSet<char>{ '金', '全', '余' },          // 金形字
            new HashSet<char>{ '银', '很', '狠' },          // 银形字
            new HashSet<char>{ '铜', '洞', '桐' },          // 铜形字
            new HashSet<char>{ '铁', '跌', '秩' },          // 铁形字
        };

        private static readonly List<HashSet<string>> similarChineseWords = new List<HashSet<string>>() {
            new HashSet<string>{ "蓝图", "蓝途" },          // 常见词
            new HashSet<string>{ "系统", "系通", "细统" },    // 常见词
            new HashSet<string>{ "神经", "神精", "神径" },    // 常见词
            new HashSet<string>{ "骨架", "骨加", "骨驾" },    // 常见词
            new HashSet<string>{ "机体", "机休" },          // 常见词
            new HashSet<string>{ "枪管", "抢管", "枪菅" },    // 常见词
            new HashSet<string>{ "枪托", "抢托", "枪拖" },    // 常见词
            new HashSet<string>{ "扳机", "搬机", "班机" },    // 常见词
            new HashSet<string>{ "弹夹", "弹甲", "弹荚" },    // 常见词
            new HashSet<string>{ "握把", "渥把", "卧把" },    // 常见词
            new HashSet<string>{ "瞄准", "喵准", "描准" },    // 常见词
        };

        public int LevenshteinDistanceChinese(string s, string t)
        {
            // i18n chinese-english mixed edit distance algorithm
            // 标准化输入：移除蓝图和多余空格
            s = " " + s.Replace("蓝图", "").Replace(" ", "");
            t = " " + t.Replace("蓝图", "").Replace(" ", "");

            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            if (n == 0 || m == 0)
                return n + m;

            // 初始化代价矩阵
            for (int i = 1; i < s.Length; i++) d[i, 0] = i * 9;
            for (int j = 1; j < t.Length; j++) d[0, j] = j * 9;

            for (int i = 1; i < s.Length; i++)
            {
                for (int j = 1; j < t.Length; j++)
                {
                    char cha = s[i];
                    char chb = t[j];
                    int cost = 9;

                    if (cha == chb)
                    {
                        cost = 0;
                    }
                    else if (AreSimilarChineseCharacters(cha, chb))
                    {
                        // 形近字，降低代价
                        cost = 3;
                    }
                    else if (AreSimilarChineseWords(s, t, i, j))
                    {
                        // 近义或常见词错误，降低代价
                        cost = 4;
                    }
                    else if (IsEnglishCharacter(cha) && IsEnglishCharacter(chb))
                    {
                        // 英文字符，考虑大小写
                        if (char.ToLower(cha) == char.ToLower(chb))
                        {
                            cost = 1; // 大小写差异代价较低
                        }
                        else
                        {
                            cost = 6; // 不同英文字符代价适中
                        }
                    }
                    else if (IsEnglishCharacter(cha) || IsEnglishCharacter(chb))
                    {
                        // 中英文混合，代价适中
                        cost = 7;
                    }

                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 9, d[i, j - 1] + 9), d[i - 1, j - 1] + cost);
                }
            }

            return d[s.Length - 1, t.Length - 1];
        }

        private bool IsEnglishCharacter(char c)
        {
            return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
        }

        private bool AreSimilarChineseCharacters(char c1, char c2)
        {
            foreach (var similarGroup in similarChineseCharacters)
            {
                if (similarGroup.Contains(c1) && similarGroup.Contains(c2))
                {
                    return true;
                }
            }
            return false;
        }

        private bool AreSimilarChineseWords(string s, string t, int i, int j)
        {
            // 检查2-3字符的常见词
            for (int length = 2; length <= 3; length++)
            {
                if (i - length + 1 >= 0 && j - length + 1 >= 0)
                {
                    string word1 = s.Substring(i - length + 1, length);
                    string word2 = t.Substring(j - length + 1, length);
                    
                    foreach (var similarGroup in similarChineseWords)
                    {
                        if (similarGroup.Contains(word1) && similarGroup.Contains(word2))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        private static bool GroupEquals(Dictionary<int, List<int>> group, int ak, int bk)
        {
            foreach (var entry in group)
            {
                if (entry.Value.Contains(ak) && entry.Value.Contains(bk))
                {
                    return true;
                }
            }
            return false;
        }

        public int LevenshteinDistanceSecond(string str1, string str2, int limit = -1)
        {
            int num;
            Boolean maxY;
            int temp;
            Boolean maxX;
            string s = str1.ToLower(Main.culture);
            string t = str2.ToLower(Main.culture);
            int n = s.Length;
            int m = t.Length;
            if (!(n == 0 || m == 0))
            {
                int[,] d = new int[n + 1 + 1 - 1, m + 1 + 1 - 1];
                List<int> activeX = new List<int>();
                List<int> activeY = new List<int>();
                d[0, 0] = 1;
                activeX.Add(0);
                activeY.Add(0);
                do
                {
                    int currX = activeX[0];
                    activeX.RemoveAt(0);
                    int currY = activeY[0];
                    activeY.RemoveAt(0);

                    temp = d[currX, currY];
                    if (limit != -1 && temp > limit)
                    {
                        return temp;
                    }

                    maxX = currX == n;
                    maxY = currY == m;
                    if (!maxX)
                    {
                        temp = d[currX, currY] + 1;
                        if (temp < d[currX + 1, currY] || d[currX + 1, currY] == 0)
                        {
                            d[currX + 1, currY] = temp;
                            AddElement(d, activeX, activeY, currX + 1, currY);
                        }
                    }

                    if (!maxY)
                    {
                        temp = d[currX, currY] + 1;
                        if (temp < d[currX, currY + 1] || d[currX, currY + 1] == 0)
                        {
                            d[currX, currY + 1] = temp;
                            AddElement(d, activeX, activeY, currX, currY + 1);
                        }
                    }

                    if (!maxX && !maxY)
                    {
                        temp = d[currX, currY] + GetDifference(s[currX], t[currY]);
                        if (temp < d[currX + 1, currY + 1] || d[currX + 1, currY + 1] == 0)
                        {
                            d[currX + 1, currY + 1] = temp;
                            AddElement(d, activeX, activeY, currX + 1, currY + 1);
                        }
                    }
                } while (!(maxX && maxY));

                num = d[n, m] - 1;
            }
            else
            {
                num = n + m;
            }

            return num;
        }

        //public string ClosestAutoComplete(string searchQuery) {
        //	return GetPartNameHuman(searchQuery, out _);
        //}

        // 中文常见识别错误映射表
        private static readonly Dictionary<string, string> ChineseRecognitionFixes = new Dictionary<string, string>
        {
            
        };
        
        // 中文名称直接映射表（不在market_items.json中的物品）
        private static readonly Dictionary<string, string> ChineseToEnglishDirectMap = new Dictionary<string, string>
        {
            { "Forma 蓝图", "Forma Blueprint" },
            { "Forma蓝图", "Forma Blueprint" },
            { "蓝图", "Blueprint" },
        };

        public string GetPartName(string name, out int low, bool suppressLogging, out bool multipleLowest)
        { // Checks the Levenshtein Distance of a string and returns the index in Names() of the closest part
            string lowest = null;
            string lowest_unfiltered = null;
            low = 9999;
            multipleLowest = false;
            
            // 处理多行文本，将换行符替换为空格
            string normalizedName = name.Replace("\n", " ").Replace("\r", "");
            
            // 检查名称是否包含中文
            bool containsChinese = System.Text.RegularExpressions.Regex.IsMatch(normalizedName, @"[\u4e00-\u9fa5]");
            
            // 特殊处理：直接匹配中文名称映射表
            if (containsChinese)
            {
                foreach (var map in ChineseToEnglishDirectMap)
                {
                    if (normalizedName.Contains(map.Key))
                    {
                        // 在nameData中查找对应的完整部件名
                        foreach (KeyValuePair<string, JToken> prop in nameData)
                        {
                            if (prop.Key == map.Value)
                            {
                                low = 0;
                                lowest = prop.Value.ToObject<string>();
                                lowest_unfiltered = prop.Key;
                                if (!suppressLogging)
                                    Main.AddLog("Found part(" + low + "): \"" + lowest_unfiltered + "\" from \"" + name + "\" (Direct Chinese map)");
                                return lowest;
                            }
                        }
                    }
                }
            }
            
            // 处理 Prime 后面的内容
            if (normalizedName.Contains("Prime"))
            {
                int primeIndex = normalizedName.IndexOf("Prime");
                string afterPrime = normalizedName.Substring(primeIndex + 5).Trim();
                if (!string.IsNullOrEmpty(afterPrime))
                {
                    // 强制标记为包含中文，确保使用中文识别规则
                    containsChinese = true;
                    // 保持原始内容，让后续的中文识别修复处理
                }
            }
            
            // 应用中文识别错误修复
            string fixedName = normalizedName;
            if (containsChinese)
            {
                foreach (var fix in ChineseRecognitionFixes)
                {
                    if (fixedName.Contains(fix.Key))
                    {
                        fixedName = fixedName.Replace(fix.Key, fix.Value);
                    }
                }
                
                // 尝试从中文映射到英文名称
                string englishFromChinese = FindBestEnglishNameFromChinese(fixedName);
                if (!string.IsNullOrEmpty(englishFromChinese))
                {
                    // 在nameData中查找对应的完整部件名
                    foreach (KeyValuePair<string, JToken> prop in nameData)
                    {
                        if (prop.Key.StartsWith(englishFromChinese))
                        {
                            int val = LevenshteinDistance(prop.Key, englishFromChinese);
                            if (val < low)
                            {
                                low = val;
                                lowest = prop.Value.ToObject<string>();
                                lowest_unfiltered = prop.Key;
                                multipleLowest = false;
                            }
                        }
                    }
                    
                    if (lowest != null)
                    {
                        if (!suppressLogging)
                            Main.AddLog("Found part(" + low + "): \"" + lowest_unfiltered + "\" from \"" + name + "\" (Chinese mapped to: " + englishFromChinese + ")");
                        return lowest;
                    }
                }
            }
            
            // 特殊处理：当识别到主要Prime武器或战甲名称时，优先匹配相关部件
            string primeName = fixedName;
            if (primeName.EndsWith(" Prime"))
            {
                primeName = primeName.Substring(0, primeName.Length - 6);
            }
            
            foreach (KeyValuePair<string, JToken> prop in nameData)
            {
                int val = LevenshteinDistance(prop.Key, fixedName);
                
                // 对于中文语言环境，优先使用本地化名称进行匹配
                if ((_settings.Locale == "zh" || _settings.Locale == "zh-CN" || _settings.Locale == "zh-hans"))
                {
                    string localeName = GetLocaleNameData(prop.Value.ToObject<string>());
                    if (!string.IsNullOrEmpty(localeName))
                    {
                        int localeVal = LevenshteinDistance(localeName, fixedName);
                        // 如果包含中文，优先使用本地化名称的匹配结果
                        if (containsChinese || localeVal < val)
                        {
                            val = localeVal;
                        }
                    }
                }
                
                // 优先匹配包含相同Prime名称的部件
                if (prop.Key.Contains(primeName + " Prime"))
                {
                    val = Math.Max(0, val - 5); // 降低包含相同Prime名称部件的距离值
                }
                
                if (val < low)
                {
                    low = val;
                    lowest = prop.Value.ToObject<string>();
                    lowest_unfiltered = prop.Key;
                    multipleLowest = false;
                }
                else if (val == low)
                {
                    multipleLowest = true;
                }

                if (val == low && lowest.StartsWith("Gara") && prop.Key.StartsWith("Ivara")) //If both
                {
                    lowest = prop.Value.ToObject<string>();
                    lowest_unfiltered = prop.Key;
                }
            }

            if (!suppressLogging)
                Main.AddLog("Found part(" + low + "): \"" + lowest_unfiltered + "\" from \"" + name + "\" (fixed: \"" + fixedName + "\"");
            return lowest;
        }

        public string GetPartNameHuman(string name, out int low)
        { // Checks the Levenshtein Distance of a string and returns the index in Names() of the closest part optimized for human searching
            string lowest = null;
            string lowest_unfiltered = null;
            low = 9999;
            string nameLower = name.ToLower(Main.culture);
            
            // 检测搜索词是否包含中文
            bool hasChinese = false;
            foreach (char c in name)
            {
                if (c >= 0x4E00 && c <= 0x9FA5)
                {
                    hasChinese = true;
                    break;
                }
            }
            
            if (hasChinese)
            {
                // 中文搜索：在中文名称中搜索
                lock (marketItemsLock)
                {
                    if (marketItems != null)
                    {
                        foreach (var marketItem in marketItems)
                        {
                            if (marketItem.Key == "version")
                                continue;
                            string[] split = marketItem.Value.ToString().Split('|');
                            if (split.Length > 2)
                            {
                                string englishName = split[0];
                                string chineseName = split[2];
                                
                                // 检查中文名称是否包含搜索词
                                if (chineseName.Contains(name))
                                {
                                    // 优先返回包含匹配，距离基于搜索词长度
                                    int val = chineseName.Length - name.Length;
                                    if (val < low)
                                    {
                                        low = val;
                                        lowest = englishName;
                                        lowest_unfiltered = chineseName;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                // 英文搜索：在英文名称中搜索
                foreach (KeyValuePair<string, JToken> prop in nameData)
                {
                    string valueLower = prop.Value.ToString().ToLower(Main.culture);
                    if (valueLower.Contains(nameLower))
                    {
                        // 优先返回包含匹配，距离基于搜索词长度
                        int val = valueLower.Length - nameLower.Length;
                        if (val < low)
                        {
                            low = val;
                            lowest = prop.Value.ToObject<string>();
                            lowest_unfiltered = prop.Value.ToString();
                        }
                    }
                }
                
                // 同时在英文名称中搜索（market_items.json）
                lock (marketItemsLock)
                {
                    if (marketItems != null)
                    {
                        foreach (var marketItem in marketItems)
                        {
                            if (marketItem.Key == "version")
                                continue;
                            string[] split = marketItem.Value.ToString().Split('|');
                            if (split.Length > 0)
                            {
                                string englishName = split[0];
                                string englishNameLower = englishName.ToLower(Main.culture);
                                
                                // 检查英文名称是否包含搜索词
                                if (englishNameLower.Contains(nameLower))
                                {
                                    int val = englishNameLower.Length - nameLower.Length;
                                    if (val < low)
                                    {
                                        low = val;
                                        lowest = englishName;
                                        lowest_unfiltered = split.Length > 2 ? split[2] : englishName;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            // 如果没有找到包含匹配，进行模糊搜索
            if (lowest == null)
            {
                if (hasChinese)
                {
                    // 中文模糊搜索
                    lock (marketItemsLock)
                    {
                        if (marketItems != null)
                        {
                            foreach (var marketItem in marketItems)
                            {
                                if (marketItem.Key == "version")
                                    continue;
                                string[] split = marketItem.Value.ToString().Split('|');
                                if (split.Length > 2)
                                {
                                    string englishName = split[0];
                                    string chineseName = split[2];
                                    
                                    int val = LevenshteinDistance(chineseName, name);
                                    if (val < low)
                                    {
                                        low = val;
                                        lowest = englishName;
                                        lowest_unfiltered = chineseName;
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    // 英文模糊搜索
                    foreach (KeyValuePair<string, JToken> prop in nameData)
                    {
                        int val = LevenshteinDistance(prop.Value.ToString().ToLower(Main.culture), nameLower);
                        if (val < low)
                        {
                            low = val;
                            lowest = prop.Value.ToObject<string>();
                            lowest_unfiltered = prop.Value.ToString();
                        }
                    }
                }
            }
            Main.AddLog("Found part(" + low + "): \"" + lowest_unfiltered + "\" from \"" + name + "\"");
            return lowest;
        }

        public static string GetSetName(string name)
        {
            string result = name.ToLower(Main.culture);

            if (result.Contains("kavasa"))
            {
                return "Kavasa Prime Kubrow Collar Set";
            }

            result = result.Replace("lower limb", "");
            result = result.Replace("upper limb", "");
            result = result.Replace("neuroptics", "");
            result = result.Replace("chassis", "");
            result = result.Replace("systems", "");
            result = result.Replace("carapace", "");
            result = result.Replace("cerebrum", "");
            result = result.Replace("blueprint", "");
            result = result.Replace("harness", "");
            result = result.Replace("blade", "");
            result = result.Replace("pouch", "");
            result = result.Replace("head", "");
            result = result.Replace("barrel", "");
            result = result.Replace("receiver", "");
            result = result.Replace("stock", "");
            result = result.Replace("disc", "");
            result = result.Replace("grip", "");
            result = result.Replace("string", "");
            result = result.Replace("handle", "");
            result = result.Replace("ornament", "");
            result = result.Replace("wings", "");
            result = result.Replace("blades", "");
            result = result.Replace("hilt", "");
            result = result.Replace("link", "");
            result = result.TrimEnd();
            result = Main.culture.TextInfo.ToTitleCase(result);
            result += " Set";
            return result;
        }

        public string GetRelicName(string string1)
        {
            string lowest = null;
            int low = 999;
            int temp;
            string eraName = null;
            JObject job = null;

            foreach (KeyValuePair<string, JToken> era in relicData)
            {
                if (!era.Key.Contains("timestamp"))
                {
                    temp = LevenshteinDistanceSecond(string1, era.Key + "??RELIC", low);
                    if (temp < low)
                    {
                        job = era.Value.ToObject<JObject>();
                        eraName = era.Key;
                        low = temp;
                    }
                }
            }

            low = 999;
            foreach (KeyValuePair<string, JToken> relic in job)
            {
                temp = LevenshteinDistanceSecond(string1, eraName + relic.Key + "RELIC", low);
                if (temp < low)
                {
                    lowest = eraName + " " + relic.Key;
                    low = temp;
                }
            }

            return lowest;
        }

        private void LogChanged(object sender, string line)
        {
            Main.AddLog("LogChanged received: " + line);
            if (autoThread != null && !autoThread.IsCompleted) 
            {
                Main.AddLog("Auto thread already running, skipping");
                return;
            }
            if (autoThread != null)
            {
                Main.AddLog("Disposing existing auto thread");
                autoThread.Dispose();
                autoThread = null;
            }

            if (line.Contains("Pause countdown done") || line.Contains("Got rewards"))
            {
                Main.AddLog("Detected reward event, starting AutoTriggered");
                autoThread = Task.Factory.StartNew(AutoTriggered);
                Overlay.rewardsDisplaying = true;
            }

            //abort if autolist and autocsv disabled, or line doesn't contain end-of-session message or timer finished message
            bool hasEndMessage = line.Contains("MatchingService::EndSession") || line.Contains("Relic timer closed");
            bool hasAutoFeature = _settings.AutoList || _settings.AutoCSV || _settings.AutoCount;
            
            Main.AddLog($"Log check - HasEndMessage: {hasEndMessage}, HasAutoFeature: {hasAutoFeature}, AutoList: {_settings.AutoList}, AutoCSV: {_settings.AutoCSV}, AutoCount: {_settings.AutoCount}");
            
            if (!hasEndMessage || !hasAutoFeature) 
            {
                Main.AddLog("No end session message or auto features disabled, skipping");
                return;
            }

            if (Main.listingHelper.PrimeRewards == null || Main.listingHelper.PrimeRewards.Count == 0)
            {
                Main.AddLog("No prime rewards found, skipping");
                return;
            }

            Main.AddLog("Processing end session with " + Main.listingHelper.PrimeRewards.Count + " rewards");
            Task.Run(async () =>
            {
                if (_settings.AutoList && string.IsNullOrEmpty(inGameName))
                    if (!await IsJWTvalid())
                    {
                        Disconnect();
                    }

                Overlay.rewardsDisplaying = false;
                string csv = "";
                Main.AddLog("Looping through rewards");
                Main.AddLog("AutoList: " + _settings.AutoList + ", AutoCSV: " + _settings.AutoCSV + ", AutoCount: " + _settings.AutoCount);
                foreach (var rewardscreen in Main.listingHelper.PrimeRewards)
                {
                    string rewards = "";
                    for (int i = 0; i < rewardscreen.Count; i++)
                    {
                        rewards += rewardscreen[i];
                        if (i + 1 < rewardscreen.Count)
                            rewards += " || ";
                    }
                    Main.AddLog(rewards + ", detected choice: " + Main.listingHelper.SelectedRewardIndex);


                    if (_settings.AutoCSV)
                    {
                        if (csv.Length == 0 && !File.Exists(applicationDirectory + @"\rewardExport.csv"))
                            csv += "Timestamp,ChosenIndex,Reward_0_Name,Reward_0_Plat,Reward_0_Ducats,Reward_1_Name,Reward_1_Plat,Reward_1_Ducats,Reward_2_Name,Reward_2_Plat,Reward_2_Ducats,Reward_3_Name,Reward_3_Plat,Reward_3_Ducats" + Environment.NewLine;
                        csv += DateTime.UtcNow.ToString("yyyy-MM-dd HH-mm-ssff", Main.culture) + "," + Main.listingHelper.SelectedRewardIndex;
                        for (int i = 0; i < 4; i++)
                        {
                            if (i < rewardscreen.Count)
                            {
                                JObject job = Main.dataBase.marketData.GetValue(rewardscreen[i]).ToObject<JObject>();
                                string plat = job["plat"].ToObject<string>();
                                string ducats = job["ducats"].ToObject<string>();
                                csv += "," + rewardscreen[i] + "," + plat + "," + ducats;
                            }
                            else
                            {
                                csv += ",\"\",0,0"; //fill empty slots with "",0,0
                            }
                        }
                        csv += Environment.NewLine;
                    }

                    if (_settings.AutoCount)
                    {
                        Main.RunOnUIThread(() =>
                        {
                            Main.autoCount.viewModel.addItem(new AutoAddSingleItem(rewardscreen, Main.listingHelper.SelectedRewardIndex, Main.autoCount.viewModel));
                        });
                    }

                    if (_settings.AutoList)
                    {

                        var rewardCollection = Task.Run(() => Main.listingHelper.GetRewardCollection(rewardscreen)).Result;
                        if (rewardCollection.PrimeNames.Count == 0)
                            continue;

                        Main.listingHelper.ScreensList.Add(new KeyValuePair<string, RewardCollection>("", rewardCollection));
                    } else
                    {
                        Main.listingHelper.SelectedRewardIndex = 0; //otherwise done by GetRewardCollection, but that calls WFM API
                    }

                }

                if (_settings.AutoCount)
                {
                    Main.AddLog("Opening AutoCount interface");
                    Main.RunOnUIThread(() =>
                    {
                        AutoCount.ShowAutoCount();
                    });
                }

                if (_settings.AutoCSV)
                {
                    Main.AddLog("appending rewardExport.csv");
                    File.AppendAllText(applicationDirectory + @"\rewardExport.csv", csv);
                }

                if (_settings.AutoList)
                {
                    Main.AddLog("Opening AutoList interface");
                    Main.RunOnUIThread(() =>
                    {
                        if (Main.listingHelper.ScreensList.Count == 1)
                            Main.listingHelper.SetScreen(0);
                        Main.listingHelper.Show();
                        Main.listingHelper.Topmost = true;
                        Main.listingHelper.Topmost = false;
                    });
                }

                Main.AddLog("Clearing listingHelper.PrimeRewards");
                Main.RunOnUIThread(() =>
                {
                    Main.listingHelper.PrimeRewards.Clear();
                });

            });
        }

        public void AutoTriggered()
        {
            try
            {
                Main.AddLog("AutoTriggered called");
                var watch = Stopwatch.StartNew();
                long stop = watch.ElapsedMilliseconds + 5000;
                long wait = watch.ElapsedMilliseconds;
                long fixedStop = watch.ElapsedMilliseconds + ApplicationSettings.GlobalReadonlySettings.FixedAutoDelay;
                const int maxRetries = 3;
                const int retryDelayMs = 500;
                int retryCount = 0;
                bool success = false;

                Main.AddLog("Updating window");
                _window.UpdateWindow();

                if (ApplicationSettings.GlobalReadonlySettings.ThemeSelection == WFtheme.AUTO)
                {
                    Main.AddLog("Using AUTO theme detection");
                    while (watch.ElapsedMilliseconds < stop && !success)
                    {
                        if (watch.ElapsedMilliseconds <= wait) continue;
                        wait += ApplicationSettings.GlobalReadonlySettings.AutoDelay;
                        Main.AddLog("Checking theme weighted");
                        OCR.GetThemeWeighted(out double diff);
                        Main.AddLog("Theme difference: " + diff);
                        if (!(diff > 40)) continue;
                        Main.AddLog("Theme detected, waiting for delay");
                        while (watch.ElapsedMilliseconds < wait) ;
                        
                        while (retryCount < maxRetries && !success)
                        {
                            retryCount++;
                            Main.AddLog($"Auto processing attempt {retryCount}/{maxRetries}");
                            success = OCR.ProcessRewardScreen();
                            if (!success && retryCount < maxRetries)
                            {
                                Main.AddLog($"Processing failed, retrying in {retryDelayMs}ms...");
                                System.Threading.Thread.Sleep(retryDelayMs);
                            }
                        }
                        break;
                    }
                } else
                {
                    Main.AddLog("Using fixed delay: " + fixedStop + "ms");
                    while (watch.ElapsedMilliseconds < fixedStop) ;
                    
                    while (retryCount < maxRetries && !success)
                    {
                        retryCount++;
                        Main.AddLog($"Auto processing attempt {retryCount}/{maxRetries}");
                        success = OCR.ProcessRewardScreen();
                        if (!success && retryCount < maxRetries)
                        {
                            Main.AddLog($"Processing failed, retrying in {retryDelayMs}ms...");
                            System.Threading.Thread.Sleep(retryDelayMs);
                        }
                    }
                }
                watch.Stop();
                Main.AddLog($"AutoTriggered completed in {watch.ElapsedMilliseconds}ms, success: {success}, retries: {retryCount}");
            }
            catch (Exception ex)
            {
                Main.AddLog("AUTO FAILED");
                Main.AddLog(ex.ToString());
                Main.StatusUpdate("Auto Detection Failed", 0);
            }
        }

        /// <summary>
        ///	Get's the user's login JWT to authenticate future API calls.
        /// </summary>
        /// <param name="email">Users email</param>
        /// <param name="password">Users password</param>
        /// <exception cref="Exception">Connection exception JSON formated</exception>
        /// <returns>A task to be awaited</returns>
        public async Task GetUserLogin(string email, string password)
        {
            var request = new HttpRequestMessage()
            {
                RequestUri = new Uri("https://api.warframe.market/v1/auth/signin"),
                Method = HttpMethod.Post,
            };
            var content = JsonConvert.SerializeObject(new
            {
                email,
                password,
                device_id = "wfinfo",
                auth_type = "header"
            });
            request.Content = new StringContent(content, System.Text.Encoding.UTF8, "application/json");
            request.Headers.Add("Authorization", "JWT");
            request.Headers.Add("language", "en");
            request.Headers.Add("accept", "application/json");
            request.Headers.Add("platform", "pc");
            var response = await client.SendAsync(request);
            var responseBody = await response.Content.ReadAsStringAsync();
            Regex rgxBody = new Regex("\"check_code\": \".*?\"");
            string censoredResponse = rgxBody.Replace(responseBody, "\"check_code\": \"REDACTED\"");
            Main.AddLog(censoredResponse);
            if (response.IsSuccessStatusCode)
            {
                SetJWT(response.Headers);
                await OpenWebSocket();
            }
            else
            {
                Regex rgxEmail = new Regex("[a-zA-Z0-9]");
                string censoredEmail = rgxEmail.Replace(email, "*");
                throw new Exception("GetUserLogin, " + responseBody + $"Email: {censoredEmail}, Pw length: {password.Length}");
            }
            request.Dispose();
        }


        // Some vibe-coded reflection modification for userAgent
        public static void SetUserAgent(ClientWebSocketOptions options, string userAgent)
        {
            try
            {
                options.SetRequestHeader("User-Agent", userAgent);
                return;
            }
            catch (System.ArgumentException)
            {
                //Debug.WriteLine(ex.ToString());
                // Fallback to reflection if User-Agent is not settable
                var field = options.GetType().GetField("_requestHeaders", BindingFlags.NonPublic | BindingFlags.Instance);
                if (field != null)
                {
                    var headers = field.GetValue(options) as System.Collections.Specialized.NameValueCollection
                                  ?? new System.Collections.Specialized.NameValueCollection();
                    headers["User-Agent"] = userAgent;
                    field.SetValue(options, headers);
                }
            }
        }

        // Listener to track messages coming back
        private async Task StartWebSocketListener()
        {
            var buffer = new byte[1024 * 4];

            try
            {
                while (marketSocket.State == WebSocketState.Open && !marketSocketCancellation.Token.IsCancellationRequested)
                {
                    var result = await marketSocket.ReceiveAsync(new ArraySegment<byte>(buffer), marketSocketCancellation.Token).ConfigureAwait(false);

                    if (result.MessageType == WebSocketMessageType.Text)
                    {
                        var message = Encoding.UTF8.GetString(buffer, 0, result.Count);
                        Debug.WriteLine($"Received: {message}");
                        await HandleWebSocketMessage(message).ConfigureAwait(false);

                        // Update last successful communication time
                        _lastConnectionTime = DateTime.UtcNow;
                    }
                    else if (result.MessageType == WebSocketMessageType.Close)
                    {
                        Debug.WriteLine("WebSocket close message received");
                        break;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // Expected during intentional shutdown
                Debug.WriteLine("WebSocket listener cancelled");
            }
            catch (WebSocketException wsEx)
            {
                Main.AddLog($"WebSocket connection error: {wsEx.Message}");
            }
            catch (ObjectDisposedException)
            {
                // Expected during shutdown
                Debug.WriteLine("WebSocket was disposed");
            }
            catch (Exception ex)
            {
                Main.AddLog($"Unexpected error in WebSocket listener: {ex.Message}");
            }
            finally
            {
                // Connection lost - trigger reconnection if not intentional
                if (!_intentionalDisconnect && IsJwtLoggedIn())
                {
                    Main.AddLog("WebSocket connection lost unexpectedly, starting reconnection attempts");
                    _ = Task.Run(StartReconnectionProcess);
                }
            }
        }

        private async Task StartReconnectionProcess()
        {
            lock (_reconnectionLock)
            {
                if (_reconnectionInProgress || _intentionalDisconnect)
                {
                    return;
                }
                _reconnectionInProgress = true;
                _reconnectionAttempts = 0;
            }

            try
            {
                while (_reconnectionAttempts < _reconnectionDelays.Length && !_intentionalDisconnect)
                {
                    _reconnectionAttempts++;
                    var delay = _reconnectionDelays[_reconnectionAttempts - 1];

                    Main.AddLog($"Attempting reconnection #{_reconnectionAttempts} in {delay / 1000} seconds...");

                    // Wait before attempting reconnection
                    await Task.Delay(delay);

                    // Check if we should still reconnect
                    if (_intentionalDisconnect || !IsJwtLoggedIn())
                    {
                        Main.AddLog("Reconnection cancelled - user disconnected or logged out");
                        break;
                    }

                    try
                    {
                        // Reset websocket state for reconnection
                        _isWebSocketAuthenticated = false;

                        // Clean up existing websocket
                        if (marketSocket != null)
                        {
                            try
                            {
                                if (marketSocket.State == WebSocketState.Open)
                                {
                                    await marketSocket.CloseAsync(WebSocketCloseStatus.NormalClosure, "Reconnecting", CancellationToken.None);
                                }
                                marketSocket.Dispose();
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }

                        // Attempt reconnection
                        bool reconnected = await OpenWebSocket();

                        if (reconnected)
                        {
                            Main.AddLog($"WebSocket reconnected successfully after {_reconnectionAttempts} attempts");
                            _reconnectionAttempts = 0;
                            break;
                        }
                        else
                        {
                            Main.AddLog($"Reconnection attempt #{_reconnectionAttempts} failed");
                        }
                    }
                    catch (Exception ex)
                    {
                        Main.AddLog($"Reconnection attempt #{_reconnectionAttempts} error: {ex.Message}");
                    }
                }

                // If all attempts failed
                if (_reconnectionAttempts >= _reconnectionDelays.Length && !_intentionalDisconnect)
                {
                    Main.AddLog("All reconnection attempts failed. Please check your connection and try logging in again.");

                    // Optional: Show user notification about connection failure
                    // You might want to update UI state here
                }
            }
            finally
            {
                lock (_reconnectionLock)
                {
                    _reconnectionInProgress = false;
                }
            }
        }

        // Add this method to handle incoming websocket messages
        private static async Task HandleWebSocketMessage(string message)
        {
            try
            {
                // Make JSON parsing async by running it on a background thread
                var messageObj = await Task.Run(() =>
                    JsonConvert.DeserializeObject<JObject>(message)
                ).ConfigureAwait(false);

                // Check for authentication success response
                var route = messageObj["route"]?.ToString();
                var payload = messageObj["payload"];

                if (route == "@wfm|cmd/auth/signIn" || route?.Contains("auth") == true)
                {
                    // Check if authentication was successful
                    var success = payload?["success"]?.ToObject<bool>() ??
                                 (payload?["error"] == null); // No error means success

                    if (success)
                    {
                        Main.AddLog("WebSocket authentication successful");
                        Main.dataBase._isWebSocketAuthenticated = true;
                        Main.dataBase._authenticationCompletionSource?.SetResult(true);
                    }
                    else
                    {
                        var error = payload?["error"]?.ToString() ?? "Unknown authentication error";
                        Main.AddLog($"WebSocket authentication failed: {error}");
                        Main.dataBase._authenticationCompletionSource?.SetResult(false);
                    }
                }

                // Handle status change messages from the server (only if authenticated)
                if (Main.dataBase._isWebSocketAuthenticated)
                {
                    var statusPayload = messageObj["payload"]?["status"]?.ToString();
                    if (!string.IsNullOrEmpty(statusPayload))
                    {
                        await Main.UpdateMarketStatusAsync(statusPayload).ConfigureAwait(false);
                    }
                }
            }
            catch (Exception e)
            {
                // Async logging
                await Task.Run(() =>
                    Main.AddLog($"Error handling websocket message: {e.Message}")
                ).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Attempts to connect the user's account to the websocket
        /// </summary>
        /// <returns>A task to be awaited</returns>
        public async Task<bool> OpenWebSocket()
        {
            Main.AddLog("Connecting to websocket");

            // Reset reconnection state for new connection attempts
            _intentionalDisconnect = false;

            // If already connected and authenticated, return true
            if (marketSocket != null && marketSocket.State == WebSocketState.Open && _isWebSocketAuthenticated)
            {
                return true;
            }

            // Clean up existing websocket if needed
            if (marketSocket != null)
            {
                try
                {
                    var currentState = marketSocket.State;
                    if (currentState == WebSocketState.Open || currentState == WebSocketState.Connecting)
                    {
                        await marketSocket.CloseAsync(WebSocketCloseStatus.NormalClosure, "Reconnecting", CancellationToken.None);
                    }
                }
                catch (InvalidOperationException)
                {
                    // Expected when websocket is not in correct state
                }
                catch (Exception ex)
                {
                    Main.AddLog($"Non-critical error closing existing websocket: {ex.Message}");
                }
                finally
                {
                    try
                    {
                        marketSocket.Dispose();
                    }
                    catch (Exception ex)
                    {
                        Main.AddLog($"Error disposing websocket: {ex.Message}");
                    }
                    marketSocket = null;
                }
            }

            // Create new websocket
            marketSocket = new ClientWebSocket();

            // Reset authentication state
            _isWebSocketAuthenticated = false;
            _authenticationCompletionSource = new TaskCompletionSource<bool>();

            try
            {
                marketSocket.Options.AddSubProtocol("wfm");
                marketSocket.Options.SetRequestHeader("Authorization", "Bearer " + JWT);
                SetUserAgent(marketSocket.Options, "WFInfo/" + Main.BuildVersion);

                Uri marketSocketUri = new Uri("wss://warframe.market/socket-v2");
                marketSocketOpenEvent.Reset();

                await marketSocket.ConnectAsync(marketSocketUri, CancellationToken.None);
                marketSocketOpenEvent.Set();

                if (marketSocket.State == WebSocketState.Open)
                {
                    Debug.WriteLine("Opening reading socket");
                    _webSocketListenerTask = Task.Run(StartWebSocketListener);

                    // Send authentication
                    bool authSuccess = await AuthenticateWebSocket();

                    if (authSuccess)
                    {
                        Main.AddLog("WebSocket connected and authenticated successfully");
                        _lastConnectionTime = DateTime.UtcNow;

                        // Send initial status update after a small delay
                        _ = Task.Delay(500).ContinueWith(async _ =>
                        {
                            if (_process.IsRunning && !_process.GameIsStreamed)
                            {
                                await SetWebsocketStatus("ingame");
                            }
                            else
                            {
                                await SetWebsocketStatus("online");
                            }
                        });
                    }

                    return authSuccess;
                }
            }
            catch (Exception ex)
            {
                Main.AddLog($"Error connecting to websocket: {ex.Message}");
                marketSocket?.Dispose();
                marketSocket = null;
                return false;
            }

            return false;
        }

        private async Task<bool> AuthenticateWebSocket()
        {
            try
            {
                bool authMessageSent = await SendMessage(
                    JsonConvert.SerializeObject(new
                    {
                        route = "@wfm|cmd/auth/signIn",
                        payload = new
                        {
                            token = JWT,
                            deviceId = "wfinfo"
                        }
                    })
                );

                if (!authMessageSent)
                {
                    Main.AddLog("Failed to send authentication message");
                    return false;
                }

                // Wait for authentication completion
                var timeoutTask = Task.Delay(TimeSpan.FromSeconds(10));
                var completedTask = await Task.WhenAny(_authenticationCompletionSource.Task, timeoutTask);

                if (completedTask == timeoutTask)
                {
                    Main.AddLog("WebSocket authentication timed out");
                    return false;
                }

                return await _authenticationCompletionSource.Task;
            }
            catch (Exception ex)
            {
                Main.AddLog($"WebSocket authentication error: {ex.Message}");
                return false;
            }
        }


        /// <summary>
        /// Sets the JWT to be used for future calls
        /// </summary>
        /// <param name="headers">Response headers from the original Login call</param>
        public void SetJWT(HttpResponseHeaders headers)
        {
            foreach (var item in headers)
            {
                //Debug.WriteLine(item.Key);
                if (!item.Key.ToLower(Main.culture).Contains("authorization")) continue;
                var temp = item.Value.First();
                // Split the second part of expression ("JWT ..." or "Bearer ...")
                JWT = temp.Split(' ')[1];
                return;
            }
        }

        /// <summary>
        /// Lists an item under an account. Expected to be called after being logged in thus no login attempts.
        /// </summary>
        /// <param name="primeItem">Human friendly for prime item</param>
        /// <param name="platinum">The amount of platinum the user entered for the listing</param>
        /// <param name="quantity">The quantity of items the user listed.</param>
        /// <returns>The success of the method</returns>
        public async Task<bool> ListItem(string primeItem, int platinum, int quantity)
        {
            try
            {
                using (var request = new HttpRequestMessage()
                {
                    RequestUri = new Uri("https://api.warframe.market/v2/order"),
                    Method = HttpMethod.Post,
                })
                {
                    var itemId = PrimeItemToItemID(primeItem);
                    var json = JsonConvert.SerializeObject(new
                    {
                        type = "sell",
                        itemId,
                        platinum,
                        quantity,
                        visible = true
                    });
                    request.Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
                    request.Headers.Add("Authorization", "Bearer " + JWT);
                    request.Headers.Add("language", "en");
                    request.Headers.Add("accept", "application/json");
                    request.Headers.Add("platform", "pc");
                    request.Headers.Add("auth_type", "header");

                    var response = await client.SendAsync(request);
                    var responseBody = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode) throw new Exception(responseBody);
                    //SetJWT(response.Headers);
                    return true;
                }
            }
            catch (Exception e)
            {
                Main.AddLog($"ListItem: {e.Message} ");
                return false;
            }

        }

        /// <summary>
        /// Updates a listing with given variables
        /// </summary>
        /// <param name="listingId">The listingID of which the listing is going to be updated</param>
        /// <param name="platinum">The new platinum value</param>
        /// <param name="quantity">The new quantity</param>
        /// <returns>The success of the method</returns>
        public async Task<bool> UpdateListing(string listingId, int platinum, int quantity)
        {
            try
            {
                using (var request = new HttpRequestMessage()
                {
                    RequestUri = new Uri("https://api.warframe.market/v2/order/" + listingId),
                    Method = HttpMethod.Put,
                })
                {
                    var json = JsonConvert.SerializeObject(new
                    {
                        platinum,
                        quantity,
                        visible = true
                    });
                    request.Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
                    request.Headers.Add("Authorization", "Bearer " + JWT);
                    request.Headers.Add("language", "en");
                    request.Headers.Add("accept", "application/json");
                    request.Headers.Add("platform", "pc");
                    request.Headers.Add("auth_type", "header");

                    var response = await client.SendAsync(request);
                    var responseBody = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode) throw new Exception(responseBody);

                    //SetJWT(response.Headers);
                    request.Dispose();
                }
                return true;
            }
            catch (Exception e)
            {
                Main.AddLog($"updateListing: {e.Message} ");
                return false;
            }
        }

        /// <summary>
        /// Converts the human friendly name to warframe.market's ID
        /// </summary>
        /// <param name="primeItem">Human friendly name of prime item</param>
        /// <returns>Warframe.market prime item ID</returns>
        public string PrimeItemToItemID(string primeItem)
        {
            lock (marketItemsLock)
            {
                if (marketItems != null) // Add null check
                {
                    foreach (var marketItem in marketItems)
                    {
                        if (marketItem.Value.ToString().Split('|').First().Equals(primeItem, StringComparison.OrdinalIgnoreCase))
                        {
                            return marketItem.Key;
                        }
                    }
                }
            }
            throw new Exception($"PrimeItemToItemID, Prime item \"{primeItem}\" does not exist in marketItem");
        }

        private readonly object _statusUpdateLock = new object();
        private DateTime _lastStatusUpdate = DateTime.MinValue;
        private string _lastStatusSent = "";
        private volatile bool _statusUpdateInProgress = false;
        public async Task SetWebsocketStatus(string status)
        {
            if (!_isWebSocketAuthenticated)
            {
                Debug.WriteLine("Cannot set websocket status: Not authenticated");
                return;
            }

            // Prevent simultaneous calls
            lock (_statusUpdateLock)
            {
                if (_statusUpdateInProgress)
                {
                    Debug.WriteLine($"Status update already in progress, skipping: {status}");
                    return;
                }

                // Prevent duplicate status within short timeframe
                var now = DateTime.UtcNow;
                if (_lastStatusSent == status && (now - _lastStatusUpdate).TotalMilliseconds < 500)
                {
                    Debug.WriteLine($"Skipping duplicate status update: {status}");
                    return;
                }

                _statusUpdateInProgress = true;
                _lastStatusUpdate = now;
                _lastStatusSent = status;
            }

            try
            {
                var payload = new { route = "@wfm|cmd/status/set", payload = new { status = status } };
                string message = JsonConvert.SerializeObject(payload);

                bool success = await SendMessage(message);
                if (success)
                {
                    Debug.WriteLine($"WebSocket status set to: {status}");
                }
                else
                {
                    Main.AddLog($"Failed to set websocket status to: {status}");
                }
            }
            finally
            {
                lock (_statusUpdateLock)
                {
                    _statusUpdateInProgress = false;
                }
            }
        }

        private SemaphoreSlim _sendSemaphore = new SemaphoreSlim(1, 1);

        // Update your SendMessage method to use the semaphore
        private async Task<bool> SendMessage(string message)
        {
            if (marketSocket == null)
            {
                Main.AddLog("Cannot send message: WebSocket is null");
                return false;
            }

            if (marketSocket.State != WebSocketState.Open)
            {
                Main.AddLog($"Cannot send message: WebSocket state is {marketSocket.State}");
                return false;
            }

            bool acquired = false;
            try
            {
                // Acquire semaphore with timeout to prevent indefinite blocking
                acquired = await _sendSemaphore.WaitAsync(TimeSpan.FromSeconds(10));
                if (!acquired)
                {
                    Main.AddLog("Failed to acquire send semaphore within timeout");
                    return false;
                }

                // Double-check websocket state after acquiring semaphore
                if (marketSocket == null || marketSocket.State != WebSocketState.Open)
                {
                    Main.AddLog("WebSocket state changed while waiting for semaphore");
                    return false;
                }

                var messageBytes = Encoding.UTF8.GetBytes(message);
                var buffer = new ArraySegment<byte>(messageBytes);

                // Send with timeout using CancellationToken
                using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15)))
                {
                    await marketSocket.SendAsync(buffer, WebSocketMessageType.Text, true, cts.Token);
                }

                Debug.WriteLine($"WebSocket message sent successfully: {message}");
                return true;
            }
            catch (OperationCanceledException)
            {
                Main.AddLog("WebSocket send operation timed out");
                return false;
            }
            catch (WebSocketException wsEx)
            {
                Main.AddLog($"WebSocket error while sending message: {wsEx.Message}");
                return false;
            }
            catch (ObjectDisposedException)
            {
                Main.AddLog("Cannot send message: WebSocket has been disposed");
                return false;
            }
            catch (InvalidOperationException ex)
            {
                Main.AddLog($"Invalid WebSocket operation: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                Main.AddLog($"Unexpected error while sending WebSocket message: {ex.Message}");
                return false;
            }
            finally
            {
                if (acquired)
                {
                    _sendSemaphore.Release();
                }
            }
        }
        /// <summary>
        /// Disconnects the user from websocket and sets JWT to null
        /// </summary>
        // Add this field to track the listener task
        private Task _webSocketListenerTask;

        public void Disconnect()
        {
            try
            {
                // Mark as intentional disconnect to prevent auto-reconnection
                _intentionalDisconnect = true;

                // Stop any ongoing reconnection attempts
                _reconnectionTimer?.Dispose();
                _reconnectionTimer = null;

                // Send invisible status first while everything is still operational
                if (marketSocket != null && marketSocket.State == WebSocketState.Open && _isWebSocketAuthenticated && IsJwtLoggedIn())
                {
                    try
                    {
                        var task = SetWebsocketStatus("invisible");
                        task.Wait(2000);
                    }
                    catch (Exception ex)
                    {
                        Main.AddLog($"Could not send invisible status: {ex.Message}");
                    }
                }

                // Reset authentication state
                _isWebSocketAuthenticated = false;

                // Complete the authentication task safely
                if (_authenticationCompletionSource != null && !_authenticationCompletionSource.Task.IsCompleted)
                {
                    _authenticationCompletionSource.TrySetResult(false);
                }
                _authenticationCompletionSource = null;

                // Cancel background operations
                marketSocketCancellation?.Cancel();

                // Dispose the send semaphore to prevent further operations
                try
                {
                    _sendSemaphore?.Dispose();
                    _sendSemaphore = new SemaphoreSlim(1, 1);
                }
                catch
                {
                    // Suppress semaphore disposal exceptions
                }

                // Wait for the listener task to complete
                if (_webSocketListenerTask != null && !_webSocketListenerTask.IsCompleted)
                {
                    try
                    {
                        _webSocketListenerTask.Wait(2000);
                    }
                    catch
                    {
                        // Suppress listener shutdown exceptions
                    }
                }

                // Dispose the websocket
                if (marketSocket != null)
                {
                    try
                    {
                        marketSocket.Dispose();
                    }
                    catch
                    {
                        // Suppress disposal exceptions
                    }
                    finally
                    {
                        marketSocket = null;
                        _webSocketListenerTask = null;
                    }
                }

                // Clear user data
                JWT = null;
                rememberMe = false;
                inGameName = string.Empty;

                // Clean up other resources
                try { marketSocketOpenEvent?.Reset(); } catch { }
                try
                {
                    marketSocketCancellation?.Dispose();
                    marketSocketCancellation = new CancellationTokenSource();
                }
                catch { }

                Main.AddLog("WebSocket disconnected successfully");
            }
            catch (Exception ex)
            {
                Main.AddLog($"Error during disconnect: {ex.Message}");
            }
            finally
            {
                // Reset the intentional disconnect flag after cleanup
                _intentionalDisconnect = false;
            }
        }

        public string GetUrlName(string primeName)
        {
            lock (marketItemsLock)
            {
                if (marketItems != null) // Add null check
                {
                    foreach (var marketItem in marketItems)
                    {
                        string[] vals = marketItem.Value.ToString().Split('|');
                        if (vals.Length > 2 && vals[0].Equals(primeName, StringComparison.OrdinalIgnoreCase))
                        {
                            return vals[1];
                        }
                    }
                }
            }
            throw new Exception($"GetUrlName, Prime item \"{primeName}\" does not exist in marketItem");
        }

        /// <summary>
        /// Tries to get the top listings of a prime item
        /// </summary>
        /// <param name="primeName"></param>
        /// <returns></returns>
        public async Task<JObject> GetTopListings(string primeName)
        {
            var urlName = GetUrlName(primeName);

            try
            {
                using (var request = new HttpRequestMessage()
                {
                    RequestUri = new Uri("https://api.warframe.market/v2/orders/item/" + urlName + "/top"),
                    Method = HttpMethod.Get
                })
                {
                    request.Headers.Add("Authorization", "Bearer " + JWT);
                    request.Headers.Add("language", "en");
                    request.Headers.Add("accept", "application/json");
                    request.Headers.Add("platform", "pc");
                    request.Headers.Add("auth_type", "header");
                    var response = await client.SendAsync(request);
                    var body = await response.Content.ReadAsStringAsync();
                    var payload = JsonConvert.DeserializeObject<JObject>(body);
                    if (body.Length < 3)
                        throw new Exception("No sell orders found: " + payload);
                    //Debug.WriteLine(body);

                    return JsonConvert.DeserializeObject<JObject>(body);
                }
            }
            catch (Exception e)
            {
                Main.AddLog("GetTopListings: " + e.Message);
                return null;
            }
        }

        /// <summary>
        /// Tries to get the profile page with the current JWT token
        /// </summary>
        /// <returns>bool of which answers the question "Is the user JWT valid?"</returns>
        public async Task<bool> IsJWTvalid()
        {
            if (JWT == null)
                return false;

            try
            {
                using (var request = new HttpRequestMessage()
                {
                    RequestUri = new Uri("https://api.warframe.market/v2/me"),
                    Method = HttpMethod.Get,
                })
                {
                    request.Headers.Add("Authorization", "Bearer " + JWT);
                    var response = await client.SendAsync(request);
                    if (response.StatusCode != HttpStatusCode.OK)
                    {
                        Main.AddLog($"JWT is invalidated or expired");
                        return false;
                    } else
                    {
                        //SetJWT(response.Headers);
                        var profile = JsonConvert.DeserializeObject<JObject>(await response.Content.ReadAsStringAsync());
                        profile["data"]["checkCode"] = "REDACTED"; // remove the code that can compromise an account.
                        Debug.WriteLine($"JWT check response: {profile["data"]}");
                        return true;
                    }
                }
            }
            catch (Exception e)
            {
                Main.AddLog($"IsJWTvalid: {e.Message} ");
                return false;
            }

        }

        /// <summary>
        /// Queries the current account for the amount of the CURRENT listed items
        /// To get the amount of a listing use:
        /// var listing = await Main.dataBase.GetCurrentListing(primeItem);
        /// var amount = (int) listing?["quantity"];
        /// To get the ID of a listing use:
        /// var listing = await Main.dataBase.GetCurrentListing(primeItem);
        /// var amount = (int) listing?["id"];
        /// </summary>
        /// <param name="primeName"></param>
        /// <returns>Quantity of prime named listed on the site</returns>
        public async Task<JToken> GetCurrentListing(string primeName)
        {
            try
            {
                if (string.IsNullOrEmpty(inGameName))
                {
                    await SetIngameName();
                }

                using (var request = new HttpRequestMessage()
                {
                    RequestUri = new Uri("https://api.warframe.market/v2/orders/my"),
                    Method = HttpMethod.Get
                })
                {
                    request.Headers.Add("Authorization", "Bearer " + JWT);
                    request.Headers.Add("language", "en");
                    request.Headers.Add("accept", "application/json");
                    request.Headers.Add("platform", "pc");
                    request.Headers.Add("auth_type", "header");
                    var response = await client.SendAsync(request);
                    var body = await response.Content.ReadAsStringAsync();
                    var payload = JsonConvert.DeserializeObject<JObject>(body);
                    var allOrders = (JArray)payload?["data"];
                    string itemID = PrimeItemToItemID(primeName);

                    if (allOrders != null)
                    {
                        foreach (var listing in allOrders)
                        {
                            if ((string)listing["type"] == "sell" && itemID == (string)listing?["itemId"])
                            {
                                request.Dispose();
                                return listing;
                            }
                        }

                        return null; //The requested item was not found, but don't throw
                    }
                    else
                    {
                        throw new Exception("No sell orders found: " + payload);
                    }
                }
            }
            catch (Exception e)
            {
                Main.AddLog("GetCurrentListing: " + e.Message);
                return null;
            }
        }


        public bool GetSocketAliveStatus()
        {
            return marketSocket.State == WebSocketState.Open;
        }

        /// <summary>
        /// Post a review on the developers page
        /// </summary>
        /// <param name="message">The content of the review</param>
        /// <returns></returns>
        public async Task<bool> PostReview(string message = "Thank you for WFinfo!")
        {
            var msg = $"{{\"text\":\"{message}\",\"review_type\":\"1\"}}";
            var developers = new List<string> { "dimon222", "Dapal003", "Kekasi", "D1firehail" };
            foreach (var developer in developers)
            {
                try
                {
                    using (var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri("https://api.warframe.market/v1/profile/" + developer + "/review"),
                        Method = HttpMethod.Post
                    })
                    {
                        request.Headers.Add("Authorization", "JWT " + JWT);
                        request.Headers.Add("language", "en");
                        request.Headers.Add("accept", "application/json");
                        request.Headers.Add("platform", "pc");
                        request.Headers.Add("auth_type", "header");
                        request.Content = new StringContent(msg, System.Text.Encoding.UTF8, "application/json");
                        var response = await client.SendAsync(request);
                        var body = await response.Content.ReadAsStringAsync();
                        Debug.WriteLine($"Body: {body}, Content: {msg}");
                    }
                }
                catch (Exception e)
                {
                    Main.AddLog("PostReview: " + e.Message);
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Gets the user's ingame name needed to make listings
        /// </summary>
        /// <returns></returns>
        public async Task SetIngameName()
        {
            using (var request = new HttpRequestMessage()
            {
                RequestUri = new Uri("https://api.warframe.market/v2/me"),
                Method = HttpMethod.Get
            })
            {
                request.Headers.Add("Authorization", "Bearer " + JWT);
                request.Headers.Add("language", "en");
                request.Headers.Add("accept", "application/json");
                request.Headers.Add("platform", "pc");
                request.Headers.Add("auth_type", "header");
                var response = await client.SendAsync(request);
                //setJWT(response.Headers);
                var profile = JsonConvert.DeserializeObject<JObject>(await response.Content.ReadAsStringAsync());
                inGameName = profile["data"]?.Value<string>("ingameName");
            }
        }

    }
}
