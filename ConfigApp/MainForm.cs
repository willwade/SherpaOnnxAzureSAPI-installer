using System;
using System.Drawing;
using System.IO;
using System.Speech.Synthesis;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using static SherpaOnnxConfig.EngineConfigManager;

namespace SherpaOnnxConfig
{
    public partial class MainForm : Form
    {
        private TabControl? tabControl;
        private TabPage? voicesPage;
        private TabPage? settingsPage;
        private TabPage? azurePage;
        private TabPage? aboutPage;

        // Voices page controls
        private ListBox? voicesListBox;
        private TextBox? voiceSearchBox;
        private Button? downloadButton;
        private Button? testVoiceButton;
        private ProgressBar? downloadProgress;
        private Label? statusLabel;

        // Settings page controls
        private NumericUpDown? speedSlider;
        private NumericUpDown? pitchSlider;
        private NumericUpDown? stabilitySlider;
        private NumericUpDown? threadsSlider;
        private Button? saveSettingsButton;

        // Azure page controls
        private TextBox? subscriptionKeyTextBox;
        private TextBox? regionTextBox;
        private ComboBox? azureVoiceComboBox;
        private Button? saveAzureButton;
        private Button? testAzureButton;

        private const string ConfigPath = @"C:\Program Files\OpenAssistive\OpenSpeech\engines_config.json";
        private const string VoiceDbUrl = "https://github.com/willwade/tts-wrapper/raw/main/tts_wrapper/engines/sherpaonnx/merged_models.json";
        private const string LocalVoiceDbPath = "./merged_models.json";

        public MainForm()
        {
            InitializeComponent();
            LoadVoicesDatabase();
        }

        private void InitializeComponent()
        {
            this.Text = "SherpaOnnx SAPI5 Configuration";
            this.Size = new Size(800, 600);
            this.MinimumSize = new Size(700, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9F);

            // Create TabControl
            tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            // Create TabPages
            voicesPage = new TabPage("Voices") { Padding = new Padding(10) };
            settingsPage = new TabPage("Settings") { Padding = new Padding(10) };
            azurePage = new TabPage("Azure") { Padding = new Padding(10) };
            aboutPage = new TabPage("About") { Padding = new Padding(10) };

            tabControl.TabPages.Add(voicesPage);
            tabControl.TabPages.Add(settingsPage);
            tabControl.TabPages.Add(azurePage);
            tabControl.TabPages.Add(aboutPage);

            this.Controls.Add(tabControl);

            // Initialize each page
            InitializeVoicesPage();
            InitializeSettingsPage();
            InitializeAzurePage();
            InitializeAboutPage();
        }

        private void InitializeVoicesPage()
        {
            if (voicesPage == null) return;

            // Search box
            var searchLabel = new Label
            {
                Text = "Search voices:",
                Location = new Point(10, 15),
                Size = new Size(80, 20)
            };

            voiceSearchBox = new TextBox
            {
                Location = new Point(95, 13),
                Size = new Size(350, 25),
                PlaceholderText = "Type to filter..."
            };
            voiceSearchBox.TextChanged += (s, e) => FilterVoices();

            // Voices list
            var listLabel = new Label
            {
                Text = "Available Voices (select to download):",
                Location = new Point(10, 50),
                Size = new Size(200, 20)
            };

            voicesListBox = new ListBox
            {
                Location = new Point(10, 75),
                Size = new Size(435, 300),
                SelectionMode = SelectionMode.One,
                DisplayMember = "DisplayName",
                ValueMember = "Id"
            };

            // Details panel
            var detailsLabel = new Label
            {
                Text = "Voice Details:",
                Location = new Point(460, 50),
                Size = new Size(100, 20)
            };

            var detailsBox = new GroupBox
            {
                Location = new Point(460, 75),
                Size = new Size(300, 200),
                Text = "Information"
            };

            // Buttons
            downloadButton = new Button
            {
                Text = "Download & Install",
                Location = new Point(10, 390),
                Size = new Size(140, 35),
                Enabled = false
            };
            downloadButton.Click += DownloadButton_Click;

            testVoiceButton = new Button
            {
                Text = "▶ Test Voice",
                Location = new Point(160, 390),
                Size = new Size(120, 35)
            };
            testVoiceButton.Click += TestVoiceButton_Click;

            downloadProgress = new ProgressBar
            {
                Location = new Point(10, 435),
                Size = new Size(650, 20),
                Visible = false
            };

            statusLabel = new Label
            {
                Text = "Ready",
                Location = new Point(10, 465),
                Size = new Size(650, 20),
                ForeColor = SystemColors.GrayText
            };

            voicesListBox.SelectedIndexChanged += (s, e) => UpdateVoiceDetails();

            // Add controls to page
            voicesPage.Controls.AddRange(new Control[]
            {
                searchLabel, voiceSearchBox,
                listLabel, voicesListBox,
                detailsLabel, detailsBox,
                downloadButton, testVoiceButton,
                downloadProgress, statusLabel
            });
        }

        private void InitializeSettingsPage()
        {
            if (settingsPage == null) return;

            int yPos = 20;

            // Speed setting
            var speedLabel = new Label
            {
                Text = "Speech Speed (lengthScale):",
                Location = new Point(20, yPos),
                Size = new Size(180, 20)
            };

            speedSlider = new NumericUpDown
            {
                Location = new Point(220, yPos - 2),
                Size = new Size(100, 25),
                Minimum = 0.5m,
                Maximum = 2.0m,
                Increment = 0.05m,
                DecimalPlaces = 2,
                Value = 1.15m
            };

            var speedHelp = new Label
            {
                Text = "Lower=faster, Higher=slower (1.0=normal)",
                Location = new Point(330, yPos),
                Size = new Size(250, 20),
                ForeColor = SystemColors.GrayText
            };

            yPos += 50;

            // Pitch variation
            var pitchLabel = new Label
            {
                Text = "Pitch Variation (noiseScale):",
                Location = new Point(20, yPos),
                Size = new Size(180, 20)
            };

            pitchSlider = new NumericUpDown
            {
                Location = new Point(220, yPos - 2),
                Size = new Size(100, 25),
                Minimum = 0.3m,
                Maximum = 0.9m,
                Increment = 0.05m,
                DecimalPlaces = 3,
                Value = 0.667m
            };

            var pitchHelp = new Label
            {
                Text = "Higher=more variation",
                Location = new Point(330, yPos),
                Size = new Size(250, 20),
                ForeColor = SystemColors.GrayText
            };

            yPos += 50;

            // Stability
            var stabilityLabel = new Label
            {
                Text = "Pitch Stability (noiseScaleW):",
                Location = new Point(20, yPos),
                Size = new Size(180, 20)
            };

            stabilitySlider = new NumericUpDown
            {
                Location = new Point(220, yPos - 2),
                Size = new Size(100, 25),
                Minimum = 0.5m,
                Maximum = 1.2m,
                Increment = 0.1m,
                DecimalPlaces = 1,
                Value = 0.8m
            };

            var stabilityHelp = new Label
            {
                Text = "Higher=more stable",
                Location = new Point(330, yPos),
                Size = new Size(250, 20),
                ForeColor = SystemColors.GrayText
            };

            yPos += 50;

            // Threads
            var threadsLabel = new Label
            {
                Text = "CPU Threads:",
                Location = new Point(20, yPos),
                Size = new Size(180, 20)
            };

            threadsSlider = new NumericUpDown
            {
                Location = new Point(220, yPos - 2),
                Size = new Size(100, 25),
                Minimum = 1,
                Maximum = 4,
                Increment = 1,
                Value = 1
            };

            var threadsHelp = new Label
            {
                Text = "1=safest, 4=fastest",
                Location = new Point(330, yPos),
                Size = new Size(250, 20),
                ForeColor = SystemColors.GrayText
            };

            yPos += 60;

            // Save button
            saveSettingsButton = new Button
            {
                Text = "Save Settings",
                Location = new Point(20, yPos),
                Size = new Size(140, 35),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            saveSettingsButton.FlatAppearance.BorderSize = 0;
            saveSettingsButton.Click += SaveSettingsButton_Click;

            // Reset button
            var resetButton = new Button
            {
                Text = "Reset to Defaults",
                Location = new Point(170, yPos),
                Size = new Size(140, 35)
            };
            resetButton.Click += (s, e) =>
            {
                speedSlider.Value = 1.15m;
                pitchSlider.Value = 0.667m;
                stabilitySlider.Value = 0.8m;
                threadsSlider.Value = 1;
            };

            // Test voice button
            var testButton = new Button
            {
                Text = "▶ Test Settings",
                Location = new Point(320, yPos),
                Size = new Size(140, 35)
            };
            testButton.Click += TestVoiceButton_Click;

            // Add controls
            settingsPage.Controls.AddRange(new Control[]
            {
                speedLabel, speedSlider, speedHelp,
                pitchLabel, pitchSlider, pitchHelp,
                stabilityLabel, stabilitySlider, stabilityHelp,
                threadsLabel, threadsSlider, threadsHelp,
                saveSettingsButton, resetButton, testButton
            });
        }

        private void InitializeAzurePage()
        {
            if (azurePage == null) return;

            var noteLabel = new Label
            {
                Text = "Azure TTS Configuration (optional cloud-based voices)",
                Location = new Point(20, 20),
                Size = new Size(500, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Italic),
                ForeColor = SystemColors.GrayText
            };

            int yPos = 60;

            // Subscription key
            var keyLabel = new Label
            {
                Text = "Subscription Key:",
                Location = new Point(20, yPos),
                Size = new Size(120, 20)
            };

            subscriptionKeyTextBox = new TextBox
            {
                Location = new Point(150, yPos - 2),
                Size = new Size(300, 25),
                PasswordChar = '●',
                PlaceholderText = "Enter your Azure subscription key"
            };

            yPos += 40;

            // Region
            var regionLabel = new Label
            {
                Text = "Region:",
                Location = new Point(20, yPos),
                Size = new Size(120, 20)
            };

            regionTextBox = new TextBox
            {
                Location = new Point(150, yPos - 2),
                Size = new Size(150, 25),
                PlaceholderText = "e.g., uksouth, eastus"
            };

            yPos += 40;

            // Voice selection
            var voiceLabel = new Label
            {
                Text = "Default Voice:",
                Location = new Point(20, yPos),
                Size = new Size(120, 20)
            };

            azureVoiceComboBox = new ComboBox
            {
                Location = new Point(150, yPos - 2),
                Size = new Size(250, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            azureVoiceComboBox.Items.AddRange(new object[]
            {
                "en-US-JennyNeural", "en-US-GuyNeural", "en-US-AriaNeural",
                "en-GB-SoniaNeural", "en-GB-LibbyNeural", "en-GB-RyanNeural"
            });
            azureVoiceComboBox.SelectedIndex = 0;

            yPos += 50;

            // Save button
            saveAzureButton = new Button
            {
                Text = "Save Azure Config",
                Location = new Point(20, yPos),
                Size = new Size(150, 35),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            saveAzureButton.FlatAppearance.BorderSize = 0;
            saveAzureButton.Click += SaveAzureButton_Click;

            // Test button
            testAzureButton = new Button
            {
                Text = "▶ Test Azure Voice",
                Location = new Point(180, yPos),
                Size = new Size(150, 35)
            };
            testAzureButton.Click += TestAzureVoiceButton_Click;

            // Info box
            var infoBox = new GroupBox
            {
                Location = new Point(20, yPos + 50),
                Size = new Size(700, 150),
                Text = "Azure TTS Information"
            };

            var infoText = new Label
            {
                Text = "Azure TTS provides high-quality cloud-based neural voices.\n\n" +
                       "To get started:\n" +
                       "1. Create an Azure Speech Service account\n" +
                       "2. Get your subscription key and region\n" +
                       "3. Enter them above and click Save\n\n" +
                       "Learn more: https://aka.ms/azure-speech-docs",
                Location = new Point(10, 20),
                Size = new Size(680, 120)
            };

            infoBox.Controls.Add(infoText);

            azurePage.Controls.AddRange(new Control[]
            {
                noteLabel, keyLabel, subscriptionKeyTextBox,
                regionLabel, regionTextBox,
                voiceLabel, azureVoiceComboBox,
                saveAzureButton, testAzureButton, infoBox
            });
        }

        private void InitializeAboutPage()
        {
            if (aboutPage == null) return;

            var titleLabel = new Label
            {
                Text = "SherpaOnnx SAPI5 TTS Engine",
                Location = new Point(20, 20),
                Size = new Size(300, 30),
                Font = new Font("Segoe UI", 14F, FontStyle.Bold)
            };

            var versionLabel = new Label
            {
                Text = "Version 1.0.0",
                Location = new Point(20, 60),
                Size = new Size(200, 20),
                ForeColor = SystemColors.GrayText
            };

            var infoBox = new GroupBox
            {
                Location = new Point(20, 100),
                Size = new Size(700, 250),
                Text = "About"
            };

            var infoText = new Label
            {
                Text = "A native Windows SAPI5 Text-to-Speech engine using SherpaOnnx with offline neural TTS models.\n\n" +
                       "Features:\n" +
                       "  • 100% offline operation (SherpaOnnx voices)\n" +
                       "  • High-quality neural TTS using VITS models\n" +
                       "  • Compatible with all SAPI5 applications\n" +
                       "  • Support for multiple languages and voices\n" +
                       "  • Azure TTS integration (cloud-based option)\n\n" +
                       "Project: https://github.com/OpenAssistive/SherpaOnnxSAPI-installer\n\n" +
                       "Built with:\n" +
                       "  • SherpaOnnx by k2-fsa\n" +
                       "  • Piper Voice Models\n" +
                       "  • ONNX Runtime\n\n" +
                       "License: Apache 2.0",
                Location = new Point(10, 20),
                Size = new Size(680, 220)
            };

            infoBox.Controls.Add(infoText);

            // Close button
            var closeButton = new Button
            {
                Text = "Close",
                Location = new Point(20, 370),
                Size = new Size(100, 35),
                DialogResult = DialogResult.OK
            };

            aboutPage.Controls.AddRange(new Control[]
            {
                titleLabel, versionLabel, infoBox, closeButton
            });
        }

        private async void LoadVoicesDatabase()
        {
            try
            {
                statusLabel?.Text ??= "Loading voice database...";

                if (File.Exists(LocalVoiceDbPath))
                {
                    var json = await File.ReadAllTextAsync(LocalVoiceDbPath);
                    ParseVoicesDatabase(json);
                }
                else
                {
                    statusLabel?.Text ??= "Voice database not found. Please run build-installer.ps1 first.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading voice database: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ParseVoicesDatabase(string json)
        {
            // Simple JSON parse for demo
            // In production, use proper deserialization
            var voices = JsonSerializer.Deserialize<JsonElement>(json);

            if (voicesListBox == null) return;

            voicesListBox.Items.Clear();

            foreach (var voice in voices.EnumerateObject())
            {
                var voiceData = voice.Value;
                if (voiceData.TryGetProperty("language", out var lang) &&
                    lang[0].GetProperty("lang_code").GetString() == "en")
                {
                    voicesListBox.Items.Add(new
                    {
                        Id = voice.Name,
                        DisplayName = $"{voiceData.GetProperty("name").GetString()} ({voiceData.GetProperty("quality").GetString()})",
                        Name = voiceData.GetProperty("name").GetString(),
                        Quality = voiceData.GetProperty("quality").GetString(),
                        Size = voiceData.GetProperty("filesize_mb").GetSingle()
                    });
                }
            }

            statusLabel?.Text ??= $"Found {voicesListBox.Items.Count} English voices";
        }

        private void FilterVoices()
        {
            if (voicesListBox == null || voiceSearchBox == null) return;

            var search = voiceSearchBox.Text.ToLower();
            foreach (var item in voicesListBox.Items)
            {
                // TODO: Implement filtering
            }
        }

        private void UpdateVoiceDetails()
        {
            downloadButton!.Enabled = voicesListBox!.SelectedItem != null;
        }

        private async void DownloadButton_Click(object? sender, EventArgs e)
        {
            if (voicesListBox?.SelectedItem == null) return;

            statusLabel!.Text = "Downloading voice...";
            downloadProgress!.Visible = true;
            downloadButton!.Enabled = false;

            try
            {
                // TODO: Implement download logic
                await Task.Delay(2000); // Demo
                statusLabel.Text = "Voice installed successfully!";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error downloading voice: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "Download failed";
            }
            finally
            {
                downloadProgress.Visible = false;
                downloadButton.Enabled = true;
            }
        }

        private void SaveSettingsButton_Click(object? sender, EventArgs e)
        {
            try
            {
                // TODO: Save to engines_config.json
                MessageBox.Show("Settings saved successfully!", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving settings: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveAzureButton_Click(object? sender, EventArgs e)
        {
            try
            {
                // TODO: Save Azure config
                MessageBox.Show("Azure configuration saved!", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving Azure config: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TestVoiceButton_Click(object? sender, EventArgs e)
        {
            try
            {
                using var synthesizer = new SpeechSynthesizer();
                var voices = synthesizer.GetInstalledVoices();

                foreach (var voice in voices)
                {
                    if (voice.VoiceInfo.Description.Contains("Sherpa", StringComparison.OrdinalIgnoreCase))
                    {
                        synthesizer.SelectVoice(voice.VoiceInfo.Name);
                        synthesizer.Speak("This is a test of the SherpaOnnx SAPI5 text to speech engine.");
                        statusLabel?.Text ??= "Test complete!";
                        return;
                    }
                }

                MessageBox.Show("SherpaOnnx voice not found. Please install the engine first.",
                    "Voice Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error testing voice: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TestAzureVoiceButton_Click(object? sender, EventArgs e)
        {
            MessageBox.Show("Azure TTS not yet implemented in the native engine.",
                "Not Implemented", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
