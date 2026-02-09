using System;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Linq;
using System.Collections.Generic;
using System.Net.Http;

namespace SherpaOnnxConfig
{
    public partial class MainForm : Form
    {
        private Label? statusLabel;
        private ComboBox? languageComboBox;
        private ComboBox? engineFilterComboBox;
        private ComboBox? voiceComboBox;
        private Button? refreshButton;
        private Button? downloadButton;
        private Button? testVoiceButton;
        private Button? installVoiceButton;
        private Button? uninstallVoiceButton;
        private RichTextBox? outputTextBox;
        private TextBox? testTextInput;

        private SherpaModelsCatalog? sherpaCatalog = null;
        private const string SapiTokensPath = @"SOFTWARE\Microsoft\Speech\Voices\Tokens";
        private static readonly string ModelsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OpenSpeech",
            "models"
        );

        private static readonly string ConfigDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OpenSpeech"
        );

        private static readonly string EnginesConfigPath = Path.Combine(ConfigDir, "engines_config.json");

        public MainForm()
        {
            InitializeComponent();
            LoadCatalogsAsync();
        }

        private void InitializeComponent()
        {
            this.Text = "SherpaOnnx SAPI5 Voice Installer";
            this.Size = new Size(760, 680);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.BackColor = Color.FromArgb(245, 245, 245);

            // Title
            Label titleLabel = new Label
            {
                Location = new Point(20, 15),
                Size = new Size(720, 25),
                Text = "SherpaOnnx SAPI5 Voice Installer",
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102)
            };

            // Status label
            statusLabel = new Label
            {
                Location = new Point(20, 45),
                Size = new Size(720, 20),
                Text = "Status: Loading voice catalog...",
                ForeColor = Color.FromArgb(100, 100, 100)
            };

            // Language selection
            Label languageLabel = new Label
            {
                Location = new Point(20, 80),
                Size = new Size(100, 20),
                Text = "Language:",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            languageComboBox = new ComboBox
            {
                Location = new Point(120, 78),
                Size = new Size(200, 25),
                DropDownStyle = ComboBoxStyle.DropDown,  // Allow typing to search
                Font = new Font("Segoe UI", 9F)
            };
            languageComboBox.Items.Add("All Languages");
            languageComboBox.SelectedIndex = 0;
            languageComboBox.SelectedIndexChanged += FilterVoices;
            languageComboBox.TextChanged += (s, e) => FilterVoices();

            // Engine filter
            Label engineLabel = new Label
            {
                Location = new Point(350, 80),
                Size = new Size(100, 20),
                Text = "Engine Type:",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            engineFilterComboBox = new ComboBox
            {
                Location = new Point(450, 78),
                Size = new Size(160, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9F)
            };
            engineFilterComboBox.Items.AddRange(new object[] { "All Engines", "SherpaOnnx (Offline)", "Azure (Cloud)" });
            engineFilterComboBox.SelectedIndex = 0;
            engineFilterComboBox.SelectedIndexChanged += FilterVoices;

            refreshButton = new Button
            {
                Location = new Point(620, 76),
                Size = new Size(80, 30),
                Text = "Refresh",
                FlatStyle = FlatStyle.Flat
            };
            refreshButton.Click += (s, e) => LoadCatalogsAsync();

            // Voice selection
            GroupBox voiceGroup = new GroupBox
            {
                Location = new Point(20, 115),
                Size = new Size(720, 110),
                Text = "Available Voices"
            };

            voiceComboBox = new ComboBox
            {
                Location = new Point(15, 25),
                Size = new Size(545, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9F)
            };
            voiceComboBox.SelectedIndexChanged += VoiceComboBox_SelectedIndexChanged;

            downloadButton = new Button
            {
                Location = new Point(570, 23),
                Size = new Size(135, 30),
                Text = "Download Model",
                BackColor = Color.FromArgb(255, 140, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Enabled = false
            };
            downloadButton.FlatAppearance.BorderSize = 0;
            downloadButton.Click += DownloadButton_Click;

            Label voiceHint = new Label
            {
                Location = new Point(15, 55),
                Size = new Size(690, 45),
                Text = "Select a voice to download, test and install. SherpaOnnx models are downloaded and cached locally.\nAzure voices require an API key in engines_config.json.",
                ForeColor = Color.FromArgb(120, 120, 120),
                Font = new Font("Segoe UI", 8F)
            };

            Label modelInfoLabel = new Label
            {
                Location = new Point(15, 85),
                Size = new Size(690, 20),
                Text = "",
                Font = new Font("Segoe UI", 8F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 100, 200)
            };

            voiceGroup.Controls.Add(voiceComboBox);
            voiceGroup.Controls.Add(downloadButton);
            voiceGroup.Controls.Add(voiceHint);
            voiceGroup.Controls.Add(modelInfoLabel);

            // Test group
            GroupBox testGroup = new GroupBox
            {
                Location = new Point(20, 235),
                Size = new Size(720, 100),
                Text = "Test Voice"
            };

            testTextInput = new TextBox
            {
                Location = new Point(15, 25),
                Size = new Size(595, 25),
                Text = "The quick brown fox jumps over the lazy dog.",
                Font = new Font("Segoe UI", 9F)
            };

            testVoiceButton = new Button
            {
                Location = new Point(620, 23),
                Size = new Size(85, 30),
                Text = "▶ Test",
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            testVoiceButton.FlatAppearance.BorderSize = 0;
            testVoiceButton.Click += TestVoiceButton_Click;

            Label hintLabel = new Label
            {
                Location = new Point(15, 55),
                Size = new Size(690, 35),
                Text = "Tests the selected voice. SherpaOnnx voices must be downloaded first. Azure voices require SAPI5 registration.",
                ForeColor = Color.FromArgb(120, 120, 120),
                Font = new Font("Segoe UI", 8F)
            };

            testGroup.Controls.Add(testTextInput);
            testGroup.Controls.Add(testVoiceButton);
            testGroup.Controls.Add(hintLabel);

            // Install/Uninstall group
            GroupBox installGroup = new GroupBox
            {
                Location = new Point(20, 345),
                Size = new Size(720, 70),
                Text = "SAPI5 Registration"
            };

            installVoiceButton = new Button
            {
                Location = new Point(15, 25),
                Size = new Size(140, 35),
                Text = "Install to SAPI5",
                BackColor = Color.FromArgb(0, 180, 80),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            installVoiceButton.FlatAppearance.BorderSize = 0;
            installVoiceButton.Click += InstallVoiceButton_Click;

            uninstallVoiceButton = new Button
            {
                Location = new Point(165, 25),
                Size = new Size(140, 35),
                Text = "Uninstall",
                BackColor = Color.FromArgb(200, 50, 50),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            uninstallVoiceButton.FlatAppearance.BorderSize = 0;
            uninstallVoiceButton.Click += UninstallVoiceButton_Click;

            Label installHint = new Label
            {
                Location = new Point(315, 25),
                Size = new Size(390, 35),
                Text = "Registers the voice with Windows SAPI5. Requires Administrator privileges.\nVoices must be downloaded before installation.",
                ForeColor = Color.FromArgb(120, 120, 120),
                Font = new Font("Segoe UI", 8F)
            };

            installGroup.Controls.Add(installVoiceButton);
            installGroup.Controls.Add(uninstallVoiceButton);
            installGroup.Controls.Add(installHint);

            // Output
            outputTextBox = new RichTextBox
            {
                Location = new Point(20, 425),
                Size = new Size(720, 230),
                ReadOnly = true,
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.FromArgb(200, 200, 200),
                Font = new Font("Consolas", 9F),
                BorderStyle = BorderStyle.FixedSingle,
                Text = "Welcome to SherpaOnnx SAPI5 Voice Installer!\r\n\r\n" +
                       "Loading voice catalog...\r\n\r\n"
            };

            this.Controls.Add(titleLabel);
            this.Controls.Add(statusLabel);
            this.Controls.Add(languageLabel);
            this.Controls.Add(languageComboBox);
            this.Controls.Add(engineLabel);
            this.Controls.Add(engineFilterComboBox);
            this.Controls.Add(refreshButton);
            this.Controls.Add(voiceGroup);
            this.Controls.Add(testGroup);
            this.Controls.Add(installGroup);
            this.Controls.Add(outputTextBox);
        }

        private HashSet<string> allLanguages = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private async void LoadCatalogsAsync()
        {
            try
            {
                statusLabel!.Text = "Status: Loading SherpaOnnx catalog...";

                string catalogPath = Path.Combine(AppContext.BaseDirectory, "merged_models.json");
                if (!File.Exists(catalogPath))
                {
                    catalogPath = "C:\\github\\SherpaOnnxAzureSAPI-installer\\ConfigApp\\merged_models.json";
                }

                if (File.Exists(catalogPath))
                {
                    string json = await File.ReadAllTextAsync(catalogPath);
                    sherpaCatalog = JsonSerializer.Deserialize<SherpaModelsCatalog>(json,
                        new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

                    AppendOutput($"Loaded {sherpaCatalog?.Count ?? 0} SherpaOnnx models from catalog.", Color.FromArgb(100, 255, 100));

                    // Extract all unique languages from catalog
                    allLanguages.Clear();
                    allLanguages.Add("All Languages");

                    if (sherpaCatalog != null)
                    {
                        foreach (var kvp in sherpaCatalog)
                        {
                            try
                            {
                                var model = kvp.Value;

                                // Process languages in SherpaLanguage format
                                if (model.Languages != null)
                                {
                                    foreach (var lang in model.Languages)
                                    {
                                        string langName = lang.LanguageName ?? "";
                                        if (!string.IsNullOrEmpty(langName))
                                            allLanguages.Add(langName);
                                    }
                                }
                                // Process languages in SherpaLanguage2 format
                                else if (model.LanguageList != null)
                                {
                                    foreach (var lang in model.LanguageList)
                                    {
                                        string langName = lang.Language_Name ?? "";
                                        if (!string.IsNullOrEmpty(langName))
                                            allLanguages.Add(langName);
                                    }
                                }
                                // Process language from LanguageData object
                                else if (model.LanguageData != null)
                                {
                                    if (model.LanguageData is JsonElement jsonElement)
                                    {
                                        if (jsonElement.ValueKind == JsonValueKind.Array)
                                        {
                                            foreach (var item in jsonElement.EnumerateArray())
                                            {
                                                if (item.ValueKind == JsonValueKind.Object)
                                                {
                                                    // Use TryGetProperty to avoid exceptions on missing keys
                                                    if (item.TryGetProperty("language_name", out var langNameProp))
                                                        allLanguages.Add(langNameProp.GetString() ?? "");

                                                    if (item.TryGetProperty("Language Name", out var langNameProp2))
                                                        allLanguages.Add(langNameProp2.GetString() ?? "");
                                                }
                                                else if (item.ValueKind == JsonValueKind.String)
                                                {
                                                    allLanguages.Add(item.GetString() ?? "");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // Skip models with problematic structures
                                continue;
                            }
                        }
                    }
                }
                else
                {
                    AppendOutput("Catalog file not found. Using built-in model list.", Color.FromArgb(255, 200, 100));
                }

                // Populate language dropdown dynamically
                PopulateLanguageDropdown();

                FilterVoices();
                statusLabel.Text = $"Status: Ready - {voiceComboBox?.Items.Count ?? 0} voices available";
            }
            catch (Exception ex)
            {
                AppendOutput($"Error loading catalog: {ex.Message}", Color.FromArgb(255, 100, 100));
                statusLabel!.Text = "Status: Error loading catalog";
            }
        }

        private void PopulateLanguageDropdown()
        {
            // Clear existing items except the first selection
            int selectedIndex = languageComboBox!.SelectedIndex;
            languageComboBox.Items.Clear();

            // Add all discovered languages
            foreach (var lang in allLanguages.OrderBy(l => l))
            {
                languageComboBox.Items.Add(lang);
            }

            languageComboBox.SelectedIndex = 0;

            AppendOutput($"Found {allLanguages.Count} unique languages in catalog.", Color.FromArgb(100, 200, 255));
        }

        private List<VoiceInfo> allVoices = new List<VoiceInfo>();

        private void FilterVoices(object? sender = null, EventArgs? e = null)
        {
            allVoices.Clear();
            voiceComboBox!.Items.Clear();

            string languageFilter = languageComboBox!.SelectedItem?.ToString()?.ToLower() ?? "";
            string engineFilter = engineFilterComboBox!.SelectedItem?.ToString()?.ToLower() ?? "";

            // Add SherpaOnnx voices from catalog
            if (sherpaCatalog != null && (engineFilter.Contains("all") || engineFilter.Contains("sherpa")))
            {
                foreach (var kvp in sherpaCatalog)
                {
                    var model = kvp.Value;

                    // Check language filter - handle both language formats
                    bool langMatch = languageFilter.Contains("all");

                    // Try different language field formats
                    List<SherpaLanguage>? languages = model.Languages;
                    List<SherpaLanguage2>? languages2 = model.LanguageList;
                    bool hasLanguageData = model.LanguageData != null;

                    // Check languages in SherpaLanguage format
                    if (languages != null && languages.Count > 0)
                    {
                        foreach (var lang in languages)
                        {
                            string langName = (lang.LanguageName ?? "").ToLower();

                            // Match if filter contains language name or if language name contains filter
                            if (!langMatch && !string.IsNullOrEmpty(languageFilter) && !string.IsNullOrEmpty(langName))
                            {
                                if (languageFilter.Contains("english") && langName.Contains("english")) langMatch = true;
                                else if (languageFilter.Contains("spanish") && langName.Contains("spanish")) langMatch = true;
                                else if (languageFilter.Contains("french") && langName.Contains("french")) langMatch = true;
                                else if (languageFilter.Contains("german") && langName.Contains("german")) langMatch = true;
                                else if (languageFilter.Contains("italian") && langName.Contains("italian")) langMatch = true;
                                else if (languageFilter.Contains("portuguese") && langName.Contains("portuguese")) langMatch = true;
                                else if (languageFilter.Contains("chinese") && langName.Contains("chinese")) langMatch = true;
                                else if (languageFilter.Contains("japanese") && langName.Contains("japanese")) langMatch = true;
                                else if (languageFilter.Contains("korean") && langName.Contains("korean")) langMatch = true;
                                else if (langName.Contains(languageFilter)) langMatch = true;
                                else if (languageFilter.Contains(langName)) langMatch = true;
                            }
                        }
                    }
                    // Check languages in SherpaLanguage2 format
                    else if (languages2 != null && languages2.Count > 0)
                    {
                        foreach (var lang in languages2)
                        {
                            string langName = (lang.Language_Name ?? "").ToLower();

                            if (!langMatch && !string.IsNullOrEmpty(languageFilter) && !string.IsNullOrEmpty(langName))
                            {
                                if (languageFilter.Contains("english") && langName.Contains("english")) langMatch = true;
                                else if (languageFilter.Contains("spanish") && langName.Contains("spanish")) langMatch = true;
                                else if (languageFilter.Contains("french") && langName.Contains("french")) langMatch = true;
                                else if (languageFilter.Contains("german") && langName.Contains("german")) langMatch = true;
                                else if (languageFilter.Contains("italian") && langName.Contains("italian")) langMatch = true;
                                else if (languageFilter.Contains("portuguese") && langName.Contains("portuguese")) langMatch = true;
                                else if (languageFilter.Contains("chinese") && langName.Contains("chinese")) langMatch = true;
                                else if (languageFilter.Contains("japanese") && langName.Contains("japanese")) langMatch = true;
                                else if (languageFilter.Contains("korean") && langName.Contains("korean")) langMatch = true;
                                else if (langName.Contains(languageFilter)) langMatch = true;
                                else if (languageFilter.Contains(langName)) langMatch = true;
                            }
                        }
                    }
                    // LanguageData handling is done in the hasLanguageData block above
                    // Handle LanguageData object (dynamic language field)
                    else if (hasLanguageData && model.LanguageData is JsonElement jsonElem)
                    {
                        if (jsonElem.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var item in jsonElem.EnumerateArray())
                            {
                                string langName = "";
                                if (item.ValueKind == JsonValueKind.Object)
                                {
                                    if (item.TryGetProperty("language_name", out var prop))
                                        langName = prop.GetString() ?? "";
                                    else if (item.TryGetProperty("Language Name", out var prop2))
                                        langName = prop2.GetString() ?? "";
                                }
                                else if (item.ValueKind == JsonValueKind.String)
                                {
                                    langName = item.GetString() ?? "";
                                }

                                if (!langMatch && !string.IsNullOrEmpty(languageFilter) && !string.IsNullOrEmpty(langName))
                                {
                                    string langLower = langName.ToLower();
                                    if (languageFilter.Contains("english") && langLower.Contains("english")) langMatch = true;
                                    else if (languageFilter.Contains("spanish") && langLower.Contains("spanish")) langMatch = true;
                                    else if (languageFilter.Contains("french") && langLower.Contains("french")) langMatch = true;
                                    else if (languageFilter.Contains("german") && langLower.Contains("german")) langMatch = true;
                                    else if (languageFilter.Contains("italian") && langLower.Contains("italian")) langMatch = true;
                                    else if (languageFilter.Contains("portuguese") && langLower.Contains("portuguese")) langMatch = true;
                                    else if (languageFilter.Contains("chinese") && langLower.Contains("chinese")) langMatch = true;
                                    else if (languageFilter.Contains("japanese") && langLower.Contains("japanese")) langMatch = true;
                                    else if (languageFilter.Contains("korean") && langLower.Contains("korean")) langMatch = true;
                                    else if (langLower.Contains(languageFilter)) langMatch = true;
                                    else if (languageFilter.Contains(langLower)) langMatch = true;
                                }
                            }
                        }
                        else if (jsonElem.ValueKind == JsonValueKind.String)
                        {
                            string langName = jsonElem.GetString() ?? "";
                            if (!langMatch && !string.IsNullOrEmpty(languageFilter) && !string.IsNullOrEmpty(langName))
                            {
                                string langLower = langName.ToLower();
                                if (languageFilter.Contains("english") && langLower.Contains("en")) langMatch = true;
                                else if (languageFilter.Contains("spanish") && langLower.Contains("es")) langMatch = true;
                                else if (languageFilter.Contains("french") && langLower.Contains("fr")) langMatch = true;
                                else if (languageFilter.Contains("german") && langLower.Contains("de")) langMatch = true;
                                else if (languageFilter.Contains("italian") && langLower.Contains("it")) langMatch = true;
                                else if (languageFilter.Contains("portuguese") && langLower.Contains("pt")) langMatch = true;
                                else if (languageFilter.Contains("chinese") && langLower.Contains("zh")) langMatch = true;
                                else if (languageFilter.Contains("japanese") && langLower.Contains("ja")) langMatch = true;
                                else if (languageFilter.Contains("korean") && langLower.Contains("ko")) langMatch = true;
                                else if (langLower.Contains(languageFilter)) langMatch = true;
                                else if (languageFilter.Contains(langLower)) langMatch = true;
                            }
                        }
                    }

                    if (!langMatch) continue;

                    // Check engine filter
                    bool engineMatch = engineFilter.Contains("all") || engineFilter.Contains("sherpa");
                    string modelType = model.GetModelType() ?? "";
                    if (engineFilter.Contains("azure") && !engineFilter.Contains("all") && !engineFilter.Contains("sherpa"))
                        continue;

                    // Get language display name
                    string langStr = "Unknown";
                    if (languages != null && languages.Count > 0)
                        langStr = languages[0].LanguageName ?? "Unknown";
                    else if (languages2 != null && languages2.Count > 0)
                        langStr = languages2[0].Language_Name ?? "Unknown";
                    else if (hasLanguageData && model.LanguageData is JsonElement jsonElem)
                    {
                        if (jsonElem.ValueKind == JsonValueKind.Array && jsonElem.GetArrayLength() > 0)
                        {
                            var first = jsonElem[0];
                            if (first.ValueKind == JsonValueKind.Object)
                            {
                                if (first.TryGetProperty("language_name", out var prop))
                                    langStr = prop.GetString() ?? "";
                                else if (first.TryGetProperty("Language Name", out var prop2))
                                    langStr = prop2.GetString() ?? "";
                            }
                            else if (first.ValueKind == JsonValueKind.String)
                            {
                                langStr = first.GetString() ?? "";
                            }
                        }
                        else if (jsonElem.ValueKind == JsonValueKind.String)
                        {
                            langStr = jsonElem.GetString() ?? "";
                        }
                    }

                    // Determine engine type for display
                    string engineType = model.Id?.Contains("mms") == true ? "MMS" : "SherpaOnnx";

                    // Convert URL format if needed (for MMS models)
                    string modelUrl = model.GetUrl() ?? "";
                    if (modelUrl.Contains("/resolve/main/") && model.Id?.Contains("mms") == true)
                    {
                        // Convert resolve URL to tree URL for listing
                        modelUrl = modelUrl.Replace("/resolve/main/", "/tree/main/");
                    }

                    allVoices.Add(new VoiceInfo
                    {
                        Id = model.Id ?? kvp.Key,
                        Name = model.Name ?? modelType,
                        Language = langStr,
                        EngineType = engineType,
                        IsOffline = true,
                        ModelUrl = modelUrl,
                        ModelSize = model.FilesizeMb,
                        SampleRate = model.SampleRate,
                        ModelType = modelType,
                        Source = "sherpa"
                    });
                }
            }

            // Add Azure voices
            if (engineFilter.Contains("all") || engineFilter.Contains("azure"))
            {
                foreach (var azureVoice in GetAzureVoices())
                {
                    bool langMatch = languageFilter.Contains("all") ||
                                   azureVoice.Language.ToLower().Contains(languageFilter);
                    if (langMatch)
                    {
                        allVoices.Add(azureVoice);
                    }
                }
            }

            // Populate dropdown
            foreach (var voice in allVoices.OrderBy(v => v.Language).ThenBy(v => v.Name))
            {
                string status = voice.IsDownloaded() ? "[✓]" : "[↓]";
                string source = voice.Source == "azure" ? "[Azure]" : $"[{voice.EngineType}]";
                voiceComboBox.Items.Add($"{voice.Id} - {voice.Language} {source} {status}");
            }

            if (voiceComboBox.Items.Count > 0)
                voiceComboBox.SelectedIndex = 0;

            statusLabel!.Text = $"Status: {voiceComboBox.Items.Count} voice(s) available";
        }

        private void VoiceComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            var voice = GetSelectedVoice();
            if (voice != null)
            {
                downloadButton!.Enabled = !voice.IsOffline || !voice.IsDownloaded();
            }
        }

        private VoiceInfo? GetSelectedVoice()
        {
            if (voiceComboBox?.SelectedItem == null) return null;
            string item = voiceComboBox.SelectedItem.ToString()!;
            string id = item.Split(' ')[0];
            return allVoices.FirstOrDefault(v => v.Id == id);
        }

        private List<VoiceInfo> GetAzureVoices()
        {
            return new List<VoiceInfo>
            {
                new VoiceInfo { Id = "jenny", Name = "Jenny Neural", Language = "English (US)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Female" },
                new VoiceInfo { Id = "guy", Name = "Guy Neural", Language = "English (US)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Male" },
                new VoiceInfo { Id = "libby", Name = "Libby Neural", Language = "English (UK)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Female" },
                new VoiceInfo { Id = "sonia", Name = "Sonia Neural", Language = "English (UK)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Female" },
                new VoiceInfo { Id = "elena", Name = "Elena Neural", Language = "Spanish (Spain)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Female" },
                new VoiceInfo { Id = "dario", Name = "Dario Neural", Language = "Spanish (Spain)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Male" },
                new VoiceInfo { Id = "denise", Name = "Denise Neural", Language = "French (France)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Female" },
                new VoiceInfo { Id = "henri", Name = "Henri Neural", Language = "French (France)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Male" },
                new VoiceInfo { Id = "katja", Name = "Katja Neural", Language = "German (Germany)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Female" },
                new VoiceInfo { Id = "conrad", Name = "Conrad Neural", Language = "German (Germany)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Male" },
                new VoiceInfo { Id = "elisa", Name = "Elsa Neural", Language = "Italian (Italy)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Female" },
                new VoiceInfo { Id = "diego", Name = "Diego Neural", Language = "Italian (Italy)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Male" },
                new VoiceInfo { Id = "francisca", Name = "Francisca Neural", Language = "Portuguese (Brazil)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Female" },
                new VoiceInfo { Id = "antonio", Name = "Antonio Neural", Language = "Portuguese (Brazil)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Male" },
                new VoiceInfo { Id = "xiaoyi", Name = "Xiaoyi Neural", Language = "Chinese (Mainland)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Female" },
                new VoiceInfo { Id = "xiaochen", Name = "Xiaochen Neural", Language = "Chinese (Mainland)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Male" },
                new VoiceInfo { Id = "nanami", Name = "Nanami Neural", Language = "Japanese (Japan)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Female" },
                new VoiceInfo { Id = "keita", Name = "Keita Neural", Language = "Japanese (Japan)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Male" },
                new VoiceInfo { Id = "sunhi", Name = "SunHi Neural", Language = "Korean (Korea)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Female" },
                new VoiceInfo { Id = "insun", Name = "InSuk Neural", Language = "Korean (Korea)", EngineType = "Azure", IsOffline = false, Source = "azure", Gender = "Male" },
            };
        }

        private async void DownloadButton_Click(object? sender, EventArgs e)
        {
            var voice = GetSelectedVoice();
            if (voice == null || voice.Source != "sherpa")
            {
                AppendOutput("ERROR: Can only download SherpaOnnx models", Color.FromArgb(255, 100, 100));
                return;
            }

            AppendOutput($"\r\n=== Downloading Model: {voice.Id} ===", Color.FromArgb(255, 140, 0));
            if (voice.ModelSize > 0)
                AppendOutput($"Size: {voice.ModelSize:F2} MB", Color.FromArgb(200, 200, 200));
            statusLabel!.Text = $"Status: Downloading {voice.Id}...";

            try
            {
                string modelDir = Path.Combine(ModelsDir, voice.Id);
                Directory.CreateDirectory(modelDir);

                if (voice.ModelUrl == null)
                {
                    AppendOutput("\rERROR: No download URL available for this model", Color.FromArgb(255, 100, 100));
                    return;
                }

                // Detect URL type and download accordingly
                if (voice.ModelUrl.Contains("huggingface.co"))
                {
                    await DownloadHuggingFaceModel(voice, modelDir);
                }
                else if (voice.ModelUrl.EndsWith(".tar.bz2") || voice.ModelUrl.EndsWith(".tar.gz") || voice.ModelUrl.Contains("tar.bz2"))
                {
                    await DownloadTarArchive(voice, modelDir);
                }
                else
                {
                    AppendOutput($"\rERROR: Unknown URL format: {voice.ModelUrl}", Color.FromArgb(255, 100, 100));
                    return;
                }

                AppendOutput($"\r✓ Model downloaded to {modelDir}", Color.FromArgb(100, 255, 100));
                statusLabel.Text = $"Status: {voice.Id} downloaded";
                FilterVoices(); // Refresh status
            }
            catch (Exception ex)
            {
                AppendOutput($"\rERROR: {ex.Message}", Color.FromArgb(255, 100, 100));
                statusLabel.Text = "Status: Download failed";
            }
        }

        private async System.Threading.Tasks.Task DownloadTarArchive(VoiceInfo voice, string modelDir)
        {
            string tarFile = Path.Combine(modelDir, "model.tar.bz2");

            AppendOutput($"\rDownloading archive from {voice.ModelUrl}...", Color.FromArgb(150, 200, 255));
            using (var client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromMinutes(30);
                var response = await client.GetAsync(voice.ModelUrl);
                response.EnsureSuccessStatusCode();

                using (var fs = File.Create(tarFile))
                {
                    await response.Content.CopyToAsync(fs);
                }
            }

            AppendOutput($"Extracting...", Color.FromArgb(150, 200, 255));

            // Extract using tar
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "tar",
                Arguments = $"-xf \"{tarFile}\" -C \"{modelDir}\"",
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = Process.Start(psi)!)
            {
                process.WaitForExit();
            }

            // Clean up tar file
            File.Delete(tarFile);
        }

        private async System.Threading.Tasks.Task DownloadHuggingFaceModel(VoiceInfo voice, string modelDir)
        {
            // For MMS models, the URL is like https://huggingface.co/willwade/mms-tts-multilingual-models-onnx/tree/main/abi
            // We need to convert to API URL to list files: https://huggingface.co/api/models/willwade/mms-tts-multilingual-models-onnx/tree/main/abi
            // Or download known files directly: tokens.txt and model.onnx

            string baseUrl = voice.ModelUrl;

            // Convert tree URL to direct file URLs
            // e.g., https://huggingface.co/willwade/mms-tts-multilingual-models-onnx/tree/main/abi
            // becomes https://huggingface.co/willwade/mms-tts-multilingual-models-onnx/resolve/main/abi/tokens.txt

            string resolveBaseUrl = baseUrl.Replace("/tree/", "/resolve/");

            AppendOutput($"\rDownloading MMS model files from HuggingFace...", Color.FromArgb(150, 200, 255));

            using (var client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromMinutes(10);
                client.DefaultRequestHeaders.Add("User-Agent", "SherpaOnnxConfig");

                // Download tokens.txt
                string tokensUrl = $"{resolveBaseUrl}/tokens.txt";
                AppendOutput($"\r  Downloading tokens.txt...", Color.FromArgb(200, 200, 200));

                try
                {
                    var tokensResponse = await client.GetAsync(tokensUrl);
                    if (tokensResponse.IsSuccessStatusCode)
                    {
                        string tokensPath = Path.Combine(modelDir, "tokens.txt");
                        File.WriteAllText(tokensPath, await tokensResponse.Content.ReadAsStringAsync());
                        AppendOutput($"    ✓ tokens.txt", Color.FromArgb(150, 255, 150));
                    }
                    else
                    {
                        AppendOutput($"    ✗ tokens.txt ({tokensResponse.StatusCode})", Color.FromArgb(255, 150, 150));
                    }
                }
                catch (Exception ex)
                {
                    AppendOutput($"    ✗ tokens.txt: {ex.Message}", Color.FromArgb(255, 150, 150));
                }

                // Download model.onnx
                string modelUrl = $"{resolveBaseUrl}/model.onnx";
                AppendOutput($"\r  Downloading model.onnx...", Color.FromArgb(200, 200, 200));

                try
                {
                    var modelResponse = await client.GetAsync(modelUrl);
                    if (modelResponse.IsSuccessStatusCode)
                    {
                        string modelPath = Path.Combine(modelDir, "model.onnx");
                        using (var fs = File.Create(modelPath))
                        {
                            await modelResponse.Content.CopyToAsync(fs);
                        }
                        AppendOutput($"    ✓ model.onnx", Color.FromArgb(150, 255, 150));
                    }
                    else
                    {
                        AppendOutput($"    ✗ model.onnx ({modelResponse.StatusCode})", Color.FromArgb(255, 150, 150));
                    }
                }
                catch (Exception ex)
                {
                    AppendOutput($"    ✗ model.onnx: {ex.Message}", Color.FromArgb(255, 150, 150));
                }

                // Try downloading model.int8.onnx if exists (common for MMS)
                string modelInt8Url = $"{resolveBaseUrl}/model.int8.onnx";
                try
                {
                    var int8Response = await client.GetAsync(modelInt8Url);
                    if (int8Response.IsSuccessStatusCode)
                    {
                        AppendOutput($"\r  Found model.int8.onnx, downloading...", Color.FromArgb(200, 200, 200));
                        string int8Path = Path.Combine(modelDir, "model.int8.onnx");
                        using (var fs = File.Create(int8Path))
                        {
                            await int8Response.Content.CopyToAsync(fs);
                        }
                        AppendOutput($"    ✓ model.int8.onnx", Color.FromArgb(150, 255, 150));
                    }
                }
                catch
                {
                    // int8 model is optional, ignore errors
                }
            }
        }

        private async void TestVoiceButton_Click(object? sender, EventArgs e)
        {
            var voice = GetSelectedVoice();
            if (voice == null)
            {
                AppendOutput("ERROR: No voice selected", Color.FromArgb(255, 100, 100));
                return;
            }

            string testText = testTextInput!.Text.Trim();
            if (string.IsNullOrEmpty(testText))
                testText = "The quick brown fox jumps over the lazy dog.";

            AppendOutput($"\r\n=== Testing Voice: {voice.Id} ===", Color.FromArgb(100, 150, 255));
            AppendOutput($"Text: \"{testText}\"", Color.FromArgb(200, 200, 200));
            statusLabel!.Text = $"Status: Testing {voice.Id}...";

            try
            {
                // Test using SAPI5 - this works for both SherpaOnnx and Azure voices
                AppendOutput("\r  Testing via SAPI5...", Color.FromArgb(150, 200, 255));

                bool voiceInstalled = await System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        // Check if voice is installed in SAPI5
                        using (var synthesizer = new System.Speech.Synthesis.SpeechSynthesizer())
                        {
                            var installedVoices = synthesizer.GetInstalledVoices();
                            var friendlyName = GetFriendlyVoiceName(voice);

                            var matchedVoice = installedVoices.FirstOrDefault(v =>
                                v.VoiceInfo.Name == friendlyName);

                            if (matchedVoice == null)
                            {
                                return false; // Voice not installed
                            }

                            // Try to select and speak
                            synthesizer.SelectVoice(friendlyName);

                            // Use async speak with timeout
                            var tcs = new System.Threading.Tasks.TaskCompletionSource<bool>();
                            synthesizer.SpeakCompleted += (s, e) => tcs.SetResult(true);

                            synthesizer.SpeakAsync(testText);

                            // Wait for completion or timeout (10 seconds)
                            if (!tcs.Task.Wait(TimeSpan.FromSeconds(10)))
                            {
                                synthesizer.SpeakAsyncCancelAll();
                                return false;
                            }

                            return true;
                        }
                    }
                    catch
                    {
                        return false;
                    }
                });

                if (!voiceInstalled)
                {
                    AppendOutput("\r  ERROR: Voice not installed in SAPI5. Click 'Install Voice' first.", Color.FromArgb(255, 100, 100));
                    AppendOutput("\r  Note: SherpaOnnx voices must be installed to SAPI5 before testing.", Color.FromArgb(200, 200, 100));
                    statusLabel.Text = "Status: Test failed";
                    return;
                }

                AppendOutput("\r  ✓ Test completed successfully!", Color.FromArgb(100, 255, 100));
                statusLabel.Text = "Status: Test complete";
            }
            catch (Exception ex)
            {
                AppendOutput($"\rERROR: {ex.Message}", Color.FromArgb(255, 100, 100));
                statusLabel.Text = "Status: Test failed";
            }
        }

        private async void InstallVoiceButton_Click(object? sender, EventArgs e)
        {
            var voice = GetSelectedVoice();
            if (voice == null)
            {
                AppendOutput("ERROR: No voice selected", Color.FromArgb(255, 100, 100));
                return;
            }

            AppendOutput($"\r\n=== Installing Voice: {voice.Id} ===", Color.FromArgb(100, 255, 100));
            statusLabel!.Text = $"Status: Installing {voice.Id}...";

            try
            {
                string dllPath = "C:\\github\\SherpaOnnxAzureSAPI-installer\\NativeTTSWrapper\\x64\\Release\\NativeTTSWrapper.dll";
                if (!File.Exists(dllPath))
                {
                    AppendOutput($"\rERROR: DLL not found at {dllPath}", Color.FromArgb(255, 100, 100));
                    return;
                }

                // Step 1: Register DLL
                AppendOutput($"\rStep 1: Registering NativeTTSWrapper.dll...", Color.FromArgb(200, 200, 200));
                await RegisterDll(dllPath);

                // Step 2: Create registry keys
                AppendOutput($"\rStep 2: Creating SAPI5 registry entries...", Color.FromArgb(200, 200, 200));
                await CreateVoiceRegistryKeys(voice);

                // Step 3: Update engines_config.json
                AppendOutput($"\rStep 3: Updating engines_config.json...", Color.FromArgb(200, 200, 200));
                UpdateEnginesConfig(voice);

                AppendOutput($"\r✓ Voice '{voice.Id}' installed successfully!", Color.FromArgb(100, 255, 100));
                statusLabel.Text = $"Status: {voice.Id} installed";
            }
            catch (Exception ex)
            {
                AppendOutput($"\rERROR: {ex.Message}", Color.FromArgb(255, 100, 100));
                statusLabel.Text = "Status: Installation failed";
            }
        }

        private async void UninstallVoiceButton_Click(object? sender, EventArgs e)
        {
            var voice = GetSelectedVoice();
            if (voice == null)
            {
                AppendOutput("ERROR: No voice selected", Color.FromArgb(255, 100, 100));
                return;
            }

            AppendOutput($"\r\n=== Uninstalling Voice: {voice.Id} ===", Color.FromArgb(255, 150, 100));
            statusLabel!.Text = $"Status: Uninstalling {voice.Id}...";

            try
            {
                await RemoveVoiceRegistryKeys(voice.Id);
                AppendOutput($"\r✓ Voice '{voice.Id}' uninstalled!", Color.FromArgb(255, 200, 100));
                statusLabel.Text = $"Status: {voice.Id} uninstalled";
            }
            catch (Exception ex)
            {
                AppendOutput($"\rERROR: {ex.Message}", Color.FromArgb(255, 100, 100));
                statusLabel.Text = "Status: Uninstallation failed";
            }
        }

        private System.Threading.Tasks.Task RegisterDll(string dllPath)
        {
            var tcs = new System.Threading.Tasks.TaskCompletionSource<bool>();

            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "regsvr32",
                    Arguments = $"\"{dllPath}\" /s",
                    UseShellExecute = true,
                    Verb = "runas",
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                Process? process = Process.Start(psi);
                if (process != null)
                {
                    process.EnableRaisingEvents = true;
                    process.Exited += (s, e) =>
                    {
                        tcs.SetResult(process.ExitCode == 0);
                        process.Dispose();
                    };
                }
                else
                {
                    tcs.SetResult(false);
                }
            }
            catch
            {
                tcs.SetResult(false);
            }

            return tcs.Task;
        }

        private async System.Threading.Tasks.Task CreateVoiceRegistryKeys(VoiceInfo voice)
        {
            await System.Threading.Tasks.Task.Run(() =>
            {
                using (RegistryKey? rootKey = Registry.LocalMachine.OpenSubKey(SapiTokensPath, true))
                {
                    if (rootKey == null)
                        throw new Exception("Run as Administrator");

                    using (RegistryKey voiceKey = rootKey.CreateSubKey(voice.Id))
                    {
                        // Use friendly name for display in SAPI apps
                        string displayName = GetFriendlyVoiceName(voice);
                        voiceKey.SetValue("", displayName);

                        using (RegistryKey attrKey = voiceKey.CreateSubKey("Attributes"))
                        {
                            string langCode = GetLanguageCode(voice.Language);
                            attrKey.SetValue("Language", langCode);
                            attrKey.SetValue("Gender", voice.Gender ?? "Female");
                            attrKey.SetValue("Age", "Adult");
                            attrKey.SetValue("Name", displayName);  // Use friendly name instead of ID
                            attrKey.SetValue("Vendor", "OpenAssistive");
                            attrKey.SetValue("Description", $"{voice.Name} TTS Voice");
                        }
                    }
                }
            });

            AppendOutput($"\r  Created: HKLM\\{SapiTokensPath}\\{voice.Id}", Color.FromArgb(150, 255, 150));
        }

        private async System.Threading.Tasks.Task RemoveVoiceRegistryKeys(string voiceId)
        {
            await System.Threading.Tasks.Task.Run(() =>
            {
                using (RegistryKey? rootKey = Registry.LocalMachine.OpenSubKey(SapiTokensPath, true))
                {
                    if (rootKey != null)
                        rootKey.DeleteSubKeyTree(voiceId, false);
                }
            });
        }

        private void UpdateEnginesConfig(VoiceInfo voice)
        {
            try
            {
                // Ensure config directory exists
                Directory.CreateDirectory(ConfigDir);

                // Read existing config or create new one
                Dictionary<string, object> config;
                string configPath = EnginesConfigPath;

                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    config = JsonSerializer.Deserialize<Dictionary<string, object>>(json) ?? new Dictionary<string, object>();
                }
                else
                {
                    // Create default config structure
                    config = new Dictionary<string, object>
                    {
                        ["engines"] = new Dictionary<string, object>(),
                        ["voices"] = new Dictionary<string, object>(),
                        ["settings"] = new Dictionary<string, object>
                        {
                            ["defaultEngine"] = "sherpa-amy",
                            ["fallbackEngine"] = "sherpa-amy"
                        }
                    };
                }

                // Get engines section
                var engines = config.ContainsKey("engines")
                    ? (Dictionary<string, object>)config["engines"]
                    : new Dictionary<string, object>();

                // Get voices section
                var voices = config.ContainsKey("voices")
                    ? (Dictionary<string, object>)config["voices"]
                    : new Dictionary<string, object>();

                // Create engine ID for this voice
                string engineId = $"sherpa-{voice.Id}";

                // Build model paths using ModelsDir (LocalApplicationData)
                string modelDir = Path.Combine(ModelsDir, voice.Id);
                string modelPath = Path.Combine(modelDir, "model.onnx");
                string tokensPath = Path.Combine(modelDir, "tokens.txt");

                // For models with espeak-ng-data
                string dataDir = "";
                if (Directory.Exists(Path.Combine(modelDir, "espeak-ng-data")))
                {
                    dataDir = Path.Combine(modelDir, "espeak-ng-data");
                }

                // Add/update engine configuration
                var engineConfig = new Dictionary<string, object>
                {
                    ["type"] = "sherpaonnx",
                    ["config"] = new Dictionary<string, object>
                    {
                        ["modelPath"] = modelPath.Replace("\\", "/"),
                        ["tokensPath"] = tokensPath.Replace("\\", "/"),
                        ["noiseScale"] = 0.667,
                        ["noiseScaleW"] = 0.8,
                        ["lengthScale"] = 1.15,
                        ["numThreads"] = 1,
                        ["provider"] = "cpu",
                        ["debug"] = true
                    }
                };

                if (!string.IsNullOrEmpty(dataDir))
                {
                    ((Dictionary<string, object>)engineConfig["config"])["dataDir"] = dataDir.Replace("\\", "/");
                }

                engines[engineId] = engineConfig;
                voices[voice.Id] = engineId;

                // Update config
                config["engines"] = engines;
                config["voices"] = voices;

                // Write back to file
                var options = new JsonSerializerOptions { WriteIndented = true };
                string jsonOutput = JsonSerializer.Serialize(config, options);
                File.WriteAllText(configPath, jsonOutput);

                AppendOutput($"\r  ✓ Updated engines_config.json at {configPath}", Color.FromArgb(100, 255, 100));
                AppendOutput($"\r  ✓ Engine ID: {engineId}", Color.FromArgb(100, 255, 100));
            }
            catch (Exception ex)
            {
                AppendOutput($"\r  ERROR updating config: {ex.Message}", Color.FromArgb(255, 100, 100));
            }
        }

        private string GetLanguageCode(string language)
        {
            if (language.Contains("US") || language.Contains("United States")) return "409";
            if (language.Contains("UK") || language.Contains("United Kingdom")) return "809";
            if (language.Contains("Spanish")) return "c0a";
            if (language.Contains("French")) return "40c";
            if (language.Contains("German")) return "407";
            if (language.Contains("Italian")) return "410";
            if (language.Contains("Portuguese")) return "416";
            if (language.Contains("Chinese")) return "804";
            if (language.Contains("Japanese")) return "411";
            if (language.Contains("Korean")) return "412";
            return "409"; // Default to en-US
        }

        private string GetFriendlyVoiceName(VoiceInfo voice)
        {
            // If voice has a proper name, use it directly
            if (!string.IsNullOrEmpty(voice.Name) && voice.Name != voice.Id)
            {
                // For names like "Hausa (MMS)", use as-is
                if (voice.Name.Contains("(") && voice.Name.Contains(")"))
                {
                    return voice.Name;
                }

                // Capitalize first letter for simple names like "amy"
                if (voice.Name.Length > 0 && char.IsLower(voice.Name[0]))
                {
                    return char.ToUpper(voice.Name[0]) + voice.Name.Substring(1);
                }

                return voice.Name;
            }

            // Fallback: create friendly name from ID
            string friendlyId = voice.Id;

            // Capitalize first letter
            if (friendlyId.Length > 0 && char.IsLower(friendlyId[0]))
            {
                friendlyId = char.ToUpper(friendlyId[0]) + friendlyId.Substring(1);
            }

            // Add engine type suffix if not already present
            if (!friendlyId.Contains("Sherpa") && !friendlyId.Contains("Azure"))
            {
                friendlyId += $" ({voice.EngineType})";
            }

            return friendlyId;
        }

        private void AppendOutput(string text, Color color)
        {
            if (outputTextBox != null)
            {
                outputTextBox.SelectionStart = outputTextBox.TextLength;
                outputTextBox.SelectionLength = 0;
                outputTextBox.SelectionColor = color;
                outputTextBox.AppendText(text + "\r\n");
                outputTextBox.SelectionStart = outputTextBox.TextLength;
                outputTextBox.ScrollToCaret();
            }
        }

        // Data classes
        private class SherpaModelsCatalog : Dictionary<string, SherpaModel> { }

        private class SherpaModel
        {
            [System.Text.Json.Serialization.JsonPropertyName("id")]
            public string? Id { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("model_type")]
            public string? ModelType { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("type")]
            public string? Type { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("developer")]
            public string? Developer { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("name")]
            public string? Name { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("language")]
            public object? LanguageData { get; set; }  // Can be array or string

            [System.Text.Json.Serialization.JsonPropertyName("languages")]
            public List<SherpaLanguage>? Languages { get; set; }

            // For MMS models - language is an array with different format
            public List<SherpaLanguage2>? LanguageList { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("quality")]
            public string? Quality { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("sample_rate")]
            public int SampleRate { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("num_speakers")]
            public int NumSpeakers { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("url")]
            public string? Url { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("compression")]
            public bool Compression { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("filesize_mb")]
            public double FilesizeMb { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("region")]
            public string? Region { get; set; }

            // Helper properties
            public string? GetUrl() => Url;
            public string? GetModelType() => ModelType ?? Type;
        }

        private class SherpaLanguage
        {
            [System.Text.Json.Serialization.JsonPropertyName("lang_code")]
            public string? LangCode { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("language_name")]
            public string? LanguageName { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("country")]
            public string? Country { get; set; }
        }

        private class SherpaLanguage2
        {
            [System.Text.Json.Serialization.JsonPropertyName("Iso Code")]
            public string? Iso_Code { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("Language Name")]
            public string? Language_Name { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("Country")]
            public string? Country { get; set; }
        }

        private class VoiceInfo
        {
            public string Id { get; set; } = "";
            public string Name { get; set; } = "";
            public string Language { get; set; } = "";
            public string EngineType { get; set; } = "";
            public bool IsOffline { get; set; }
            public bool IsDownloaded() => IsOffline && Directory.Exists(Path.Combine(ModelsDir, Id));
            public string? ModelUrl { get; set; }
            public double ModelSize { get; set; }
            public int SampleRate { get; set; }
            public string? ModelType { get; set; }
            public string? Source { get; set; }
            public string? Gender { get; set; }
        }
    }
}
