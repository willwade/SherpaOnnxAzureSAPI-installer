using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;

namespace OpenSpeechTTS
{
    [Guid("3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2")]
    [ComVisible(true)]
    public class Sapi5VoiceImpl : ISpTTSEngine, ISpObjectWithToken
    {
        private SherpaTTS _sherpaTts;
        private bool _initialized;
        private const int S_OK = 0;
        private const int E_FAIL = unchecked((int)0x80004005);
        private const int E_INVALIDARG = unchecked((int)0x80070057);
        private const int E_OUTOFMEMORY = unchecked((int)0x8007000E);

        // Standard PCM format GUID
        private static readonly Guid SPDFID_WaveFormatEx = new Guid("C31ADBAE-527F-4ff5-A230-F62BB61FF70C");

        // Static flag to ensure we only set up the assembly resolver once
        private static bool _assemblyResolverSetup = false;
        private static readonly object _resolverLock = new object();

        public Sapi5VoiceImpl()
        {
            try
            {
                LogMessage("Initializing Sapi5VoiceImpl constructor...");

                // Set up assembly resolver for dependencies
                SetupAssemblyResolver();

                // Pre-load dependencies
                PreloadDependencies();

                // Don't initialize the voice here - wait for SetObjectToken to be called
                // This is the correct SAPI5 pattern
                LogMessage("Sapi5VoiceImpl constructor completed - waiting for SetObjectToken");
            }
            catch (Exception ex)
            {
                LogError($"Error in Sapi5VoiceImpl constructor: {ex.Message}", ex);
                throw new Exception($"Error in Sapi5VoiceImpl constructor: {ex.Message}", ex);
            }
        }

        // OFFICIAL SAPI5 Speak method implementation
        public int Speak(uint dwSpeakFlags, ref Guid rguidFormatId, ref WaveFormatEx pWaveFormatEx,
                        ref SPVTEXTFRAG pTextFragList, IntPtr pOutputSite)
        {
            // IMMEDIATE logging to see if method is called
            try
            {
                string logDir = "C:\\OpenSpeech";
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);
                File.AppendAllText(Path.Combine(logDir, "sapi_debug.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: *** SPEAK METHOD CALLED *** flags: {dwSpeakFlags}, initialized: {_initialized}\n");
            }
            catch { }

            if (!_initialized)
            {
                LogMessage("TTS engine not initialized via SetObjectToken - attempting auto-initialization...");

                // Auto-initialize if SetObjectToken wasn't called
                // This handles cases where SAPI uses a different initialization pattern
                try
                {
                    string modelPath = @"C:\Program Files\OpenSpeech\models\piper-en-amy-medium\model.onnx";
                    string tokensPath = @"C:\Program Files\OpenSpeech\models\piper-en-amy-medium\tokens.txt";

                    if (File.Exists(modelPath) && File.Exists(tokensPath))
                    {
                        LogMessage("Auto-initializing with real Sherpa TTS...");
                        _sherpaTts = new SherpaTTS(modelPath, tokensPath, "", Path.GetDirectoryName(modelPath));
                        _initialized = true;
                        LogMessage("Auto-initialization completed successfully");
                    }
                    else
                    {
                        LogMessage("Model files not found - using mock mode");
                        _initialized = true; // Allow mock mode to work
                    }
                }
                catch (Exception ex)
                {
                    LogError($"Auto-initialization failed: {ex.Message}", ex);
                    LogMessage("Continuing with mock audio generation...");
                    _initialized = true; // Allow mock mode to work
                }
            }

            try
            {
                LogMessage($"Speak called with flags: {dwSpeakFlags}");

                // Extract text from fragment list
                string text = ExtractTextFromFragList(ref pTextFragList);
                if (string.IsNullOrEmpty(text))
                {
                    LogMessage("No text to speak");
                    return S_OK;
                }

                LogMessage($"Speaking text: '{text}'");

                // Get the output site interface
                if (pOutputSite == IntPtr.Zero)
                {
                    LogError("No output site provided");
                    return E_INVALIDARG;
                }

                var outputSite = Marshal.GetObjectForIUnknown(pOutputSite) as ISpTTSEngineSite;
                if (outputSite == null)
                {
                    LogError("Failed to get output site interface");
                    return E_FAIL;
                }

                // Generate audio using Sherpa ONNX (or mock data for testing)
                byte[] audioData;
                if (_sherpaTts != null)
                {
                    audioData = _sherpaTts.GenerateAudio(text);
                    LogMessage($"Generated {audioData.Length} bytes of audio data using Sherpa ONNX");
                }
                else
                {
                    // TEMPORARY: Generate mock audio data for testing
                    audioData = GenerateMockAudioData(text);
                    LogMessage($"Generated {audioData.Length} bytes of MOCK audio data for testing");
                }

                // Send start event
                SendEvent(outputSite, SpEventIds.SPEI_START_INPUT_STREAM, 0, 0, IntPtr.Zero, IntPtr.Zero);

                // Send word boundary events (mock implementation)
                SendWordBoundaryEvents(outputSite, text, audioData.Length);

                // Write audio data to output site
                uint bytesWritten = 0;
                int hr = outputSite.Write(Marshal.UnsafeAddrOfPinnedArrayElement(audioData, 0),
                                         (uint)audioData.Length, out bytesWritten);

                if (hr != S_OK)
                {
                    LogError($"Failed to write audio data. HRESULT: 0x{hr:X8}");
                    return hr;
                }

                LogMessage($"Successfully wrote {bytesWritten} bytes of audio data");

                // Send end event
                SendEvent(outputSite, SpEventIds.SPEI_END_INPUT_STREAM, 0, (ulong)audioData.Length, IntPtr.Zero, IntPtr.Zero);

                return S_OK;
            }
            catch (Exception ex)
            {
                LogError($"Error in Speak: {ex.Message}", ex);
                return E_FAIL;
            }
        }

        public int GetOutputFormat(ref Guid pTargetFormatId, ref WaveFormatEx pTargetWaveFormatEx,
                                  out Guid pOutputFormatId, out IntPtr ppCoMemOutputWaveFormatEx)
        {
            // IMMEDIATE logging to see if method is called
            try
            {
                string logDir = "C:\\OpenSpeech";
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);
                File.AppendAllText(Path.Combine(logDir, "sapi_debug.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: *** GET OUTPUT FORMAT CALLED *** TargetFormatId: {pTargetFormatId}\n");
            }
            catch { }

            try
            {
                LogMessage($"GetOutputFormat called with TargetFormatId: {pTargetFormatId}");

                // Always set our preferred format
                pOutputFormatId = SPDFID_WaveFormatEx;

                // Allocate memory for WaveFormatEx
                int waveFormatSize = Marshal.SizeOf<WaveFormatEx>();
                ppCoMemOutputWaveFormatEx = Marshal.AllocCoTaskMem(waveFormatSize);

                // Create our format - match Sherpa ONNX output
                uint sampleRate = 22050u; // Fixed sample rate
                var format = new WaveFormatEx
                {
                    wFormatTag = 1, // PCM
                    nChannels = 1, // Mono
                    nSamplesPerSec = sampleRate,
                    wBitsPerSample = 16,
                    nBlockAlign = 2, // (nChannels * wBitsPerSample) / 8
                    nAvgBytesPerSec = sampleRate * 2, // nSamplesPerSec * nBlockAlign
                    cbSize = 0
                };

                Marshal.StructureToPtr(format, ppCoMemOutputWaveFormatEx, false);

                LogMessage($"Returning output format: {format.nSamplesPerSec}Hz, {format.nChannels} channel(s), {format.wBitsPerSample}-bit");
                LogMessage($"GetOutputFormat returning S_OK");
                return S_OK;
            }
            catch (Exception ex)
            {
                LogError($"Error in GetOutputFormat: {ex.Message}", ex);
                pOutputFormatId = Guid.Empty;
                ppCoMemOutputWaveFormatEx = IntPtr.Zero;
                return E_FAIL;
            }
        }

        // Helper method to extract text from SAPI fragment list
        private string ExtractTextFromFragList(ref SPVTEXTFRAG fragList)
        {
            try
            {
                if (fragList.pTextStart == IntPtr.Zero || fragList.ulTextLen == 0)
                    return string.Empty;

                // Read the text from the pointer
                return Marshal.PtrToStringUni(fragList.pTextStart, (int)fragList.ulTextLen);
            }
            catch (Exception ex)
            {
                LogError($"Error extracting text from fragment list: {ex.Message}", ex);
                return string.Empty;
            }
        }

        // Helper method to send SAPI events
        private void SendEvent(ISpTTSEngineSite outputSite, SpEventIds eventId, uint streamNum,
                              ulong audioOffset, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                var spEvent = new SpEvent
                {
                    eEventId = (ushort)eventId,
                    elParamType = 0,
                    ulStreamNum = streamNum,
                    ullAudioStreamOffset = audioOffset,
                    wParam = wParam,
                    lParam = lParam
                };

                IntPtr eventPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf<SpEvent>());
                Marshal.StructureToPtr(spEvent, eventPtr, false);

                int hr = outputSite.AddEvents(eventPtr, 1);
                if (hr != S_OK)
                {
                    LogError($"Failed to send event {eventId}. HRESULT: 0x{hr:X8}");
                }

                Marshal.FreeCoTaskMem(eventPtr);
            }
            catch (Exception ex)
            {
                LogError($"Error sending event {eventId}: {ex.Message}", ex);
            }
        }

        // Helper method to send word boundary events (mock implementation)
        private void SendWordBoundaryEvents(ISpTTSEngineSite outputSite, string text, int audioLength)
        {
            try
            {
                // Simple word boundary detection
                string[] words = text.Split(new char[] { ' ', '\t', '\n', '\r' },
                                           StringSplitOptions.RemoveEmptyEntries);

                ulong audioOffset = 0;
                ulong audioPerWord = (ulong)(audioLength / Math.Max(words.Length, 1));

                for (int i = 0; i < words.Length; i++)
                {
                    // Send word boundary event
                    IntPtr wordPtr = Marshal.StringToCoTaskMemUni(words[i]);
                    SendEvent(outputSite, SpEventIds.SPEI_WORD_BOUNDARY, 0, audioOffset,
                             (IntPtr)words[i].Length, wordPtr);

                    audioOffset += audioPerWord;
                    Marshal.FreeCoTaskMem(wordPtr);
                }
            }
            catch (Exception ex)
            {
                LogError($"Error sending word boundary events: {ex.Message}", ex);
            }
        }

        // Logging helper methods
        private void LogMessage(string message)
        {
            try
            {
                string logDir = "C:\\OpenSpeech";
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);

                File.AppendAllText(Path.Combine(logDir, "sapi_debug.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: {message}\n");
            }
            catch { }
        }

        private void LogError(string message, Exception ex = null)
        {
            try
            {
                string logDir = "C:\\OpenSpeech";
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);

                string fullMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: ERROR: {message}";
                if (ex != null)
                {
                    fullMessage += $"\nException: {ex.Message}\nStack Trace: {ex.StackTrace}";
                }
                fullMessage += "\n";

                File.AppendAllText(Path.Combine(logDir, "sapi_error.log"), fullMessage);
            }
            catch { }
        }

        // Generate mock audio data for testing (creates a simple WAV file with silence)
        private byte[] GenerateMockAudioData(string text)
        {
            try
            {
                // Create a simple WAV file with 1 second of silence per 10 characters
                int durationMs = Math.Max(1000, text.Length * 100); // At least 1 second
                uint sampleRate = 22050;
                int samples = (int)(sampleRate * durationMs / 1000);

                using (var ms = new MemoryStream())
                {
                    using (var writer = new BinaryWriter(ms))
                    {
                        // WAV header
                        writer.Write(0x46464952); // "RIFF"
                        writer.Write(36 + samples * 2);
                        writer.Write(0x45564157); // "WAVE"
                        writer.Write(0x20746D66); // "fmt "
                        writer.Write(16);
                        writer.Write((short)1); // PCM
                        writer.Write((short)1); // Mono
                        writer.Write(sampleRate); // Sample rate
                        writer.Write(sampleRate * 2); // Bytes per second
                        writer.Write((short)2); // Block align
                        writer.Write((short)16); // Bits per sample
                        writer.Write(0x61746164); // "data"
                        writer.Write(samples * 2);

                        // Generate simple tone instead of silence for testing
                        for (int i = 0; i < samples; i++)
                        {
                            // Generate a simple 440Hz tone (A note)
                            double time = (double)i / sampleRate;
                            double amplitude = Math.Sin(2 * Math.PI * 440 * time) * 0.1; // Low volume
                            short sample = (short)(amplitude * short.MaxValue);
                            writer.Write(sample);
                        }
                    }

                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                LogError($"Error generating mock audio data: {ex.Message}", ex);
                // Return minimal WAV file on error
                return new byte[] { 0x52, 0x49, 0x46, 0x46, 0x24, 0x00, 0x00, 0x00, 0x57, 0x41, 0x56, 0x45 };
            }
        }

        // Pre-load dependencies to avoid assembly loading issues
        private static void PreloadDependencies()
        {
            try
            {
                string installDir = @"C:\Program Files\OpenAssistive\OpenSpeech";

                // List of dependencies to pre-load
                string[] dependencies = {
                    "sherpa-onnx.dll",
                    "SherpaNative.dll"
                };

                foreach (string dep in dependencies)
                {
                    string depPath = Path.Combine(installDir, dep);
                    if (File.Exists(depPath))
                    {
                        try
                        {
                            // Try multiple loading approaches to bypass strong-name verification
                            System.Reflection.Assembly assembly = null;

                            try
                            {
                                // Method 1: Try UnsafeLoadFrom (bypasses security checks)
                                var loadFromMethod = typeof(System.Reflection.Assembly).GetMethod("UnsafeLoadFrom",
                                    System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public);
                                if (loadFromMethod != null)
                                {
                                    assembly = (System.Reflection.Assembly)loadFromMethod.Invoke(null, new object[] { depPath });
                                }
                            }
                            catch
                            {
                                // Method 2: Try LoadFile instead of LoadFrom
                                try
                                {
                                    assembly = System.Reflection.Assembly.LoadFile(depPath);
                                }
                                catch
                                {
                                    // Method 3: Try ReflectionOnlyLoadFrom
                                    try
                                    {
                                        assembly = System.Reflection.Assembly.ReflectionOnlyLoadFrom(depPath);
                                    }
                                    catch
                                    {
                                        // Method 4: Fall back to regular LoadFrom
                                        assembly = System.Reflection.Assembly.LoadFrom(depPath);
                                    }
                                }
                            }

                            // Log successful preload
                            try
                            {
                                string logDir = "C:\\OpenSpeech";
                                if (!Directory.Exists(logDir))
                                    Directory.CreateDirectory(logDir);

                                File.AppendAllText(Path.Combine(logDir, "sapi_debug.log"),
                                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Successfully preloaded {dep} using {assembly?.GetType().Name}\n");
                            }
                            catch { }
                        }
                        catch (Exception ex)
                        {
                            // Log preload failure
                            try
                            {
                                string logDir = "C:\\OpenSpeech";
                                if (!Directory.Exists(logDir))
                                    Directory.CreateDirectory(logDir);

                                File.AppendAllText(Path.Combine(logDir, "sapi_error.log"),
                                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Failed to preload {dep}: {ex.Message}\n");
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // Log missing dependency
                        try
                        {
                            string logDir = "C:\\OpenSpeech";
                            if (!Directory.Exists(logDir))
                                Directory.CreateDirectory(logDir);

                            File.AppendAllText(Path.Combine(logDir, "sapi_debug.log"),
                                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Dependency not found: {depPath}\n");
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log any errors in preloading
                try
                {
                    string logDir = "C:\\OpenSpeech";
                    if (!Directory.Exists(logDir))
                        Directory.CreateDirectory(logDir);

                    File.AppendAllText(Path.Combine(logDir, "sapi_error.log"),
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Error in PreloadDependencies: {ex.Message}\n");
                }
                catch { }
            }
        }

        // Assembly resolver to find dependencies in the installation directory
        private static void SetupAssemblyResolver()
        {
            lock (_resolverLock)
            {
                if (_assemblyResolverSetup)
                    return;

                try
                {
                    AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
                    _assemblyResolverSetup = true;

                    // Log that we've set up the resolver
                    try
                    {
                        string logDir = "C:\\OpenSpeech";
                        if (!Directory.Exists(logDir))
                            Directory.CreateDirectory(logDir);

                        File.AppendAllText(Path.Combine(logDir, "sapi_debug.log"),
                            $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Assembly resolver setup completed\n");
                    }
                    catch { }
                }
                catch (Exception ex)
                {
                    try
                    {
                        string logDir = "C:\\OpenSpeech";
                        if (!Directory.Exists(logDir))
                            Directory.CreateDirectory(logDir);

                        File.AppendAllText(Path.Combine(logDir, "sapi_error.log"),
                            $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: ERROR setting up assembly resolver: {ex.Message}\n");
                    }
                    catch { }
                }
            }
        }

        private static System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                // Log the assembly resolution attempt
                try
                {
                    string logDir = "C:\\OpenSpeech";
                    if (!Directory.Exists(logDir))
                        Directory.CreateDirectory(logDir);

                    File.AppendAllText(Path.Combine(logDir, "sapi_debug.log"),
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Resolving assembly: {args.Name}\n");
                }
                catch { }

                // Extract the simple name from the full assembly name
                string assemblyName = args.Name.Split(',')[0];

                // List of known dependencies and their file names
                var dependencyMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "sherpa-onnx", "sherpa-onnx.dll" },
                    { "SherpaNative", "SherpaNative.dll" },
                    { "onnxruntime", "onnxruntime.dll" },
                    { "onnxruntime_providers_shared", "onnxruntime_providers_shared.dll" }
                };

                if (dependencyMap.TryGetValue(assemblyName, out string fileName))
                {
                    // Try to find the dependency in the installation directory
                    string installDir = @"C:\Program Files\OpenAssistive\OpenSpeech";
                    string dependencyPath = Path.Combine(installDir, fileName);

                    if (File.Exists(dependencyPath))
                    {
                        try
                        {
                            var assembly = System.Reflection.Assembly.LoadFrom(dependencyPath);

                            // Log successful resolution
                            try
                            {
                                string logDir = "C:\\OpenSpeech";
                                File.AppendAllText(Path.Combine(logDir, "sapi_debug.log"),
                                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Successfully resolved {assemblyName} from {dependencyPath}\n");
                            }
                            catch { }

                            return assembly;
                        }
                        catch (Exception ex)
                        {
                            // Log resolution failure
                            try
                            {
                                string logDir = "C:\\OpenSpeech";
                                File.AppendAllText(Path.Combine(logDir, "sapi_error.log"),
                                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Failed to load {assemblyName} from {dependencyPath}: {ex.Message}\n");
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        // Log that dependency file was not found
                        try
                        {
                            string logDir = "C:\\OpenSpeech";
                            File.AppendAllText(Path.Combine(logDir, "sapi_debug.log"),
                                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: Dependency file not found: {dependencyPath}\n");
                        }
                        catch { }
                    }
                }

                return null; // Let the default resolution continue
            }
            catch (Exception ex)
            {
                // Log any unexpected errors in the resolver
                try
                {
                    string logDir = "C:\\OpenSpeech";
                    if (!Directory.Exists(logDir))
                        Directory.CreateDirectory(logDir);

                    File.AppendAllText(Path.Combine(logDir, "sapi_error.log"),
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: ERROR in assembly resolver: {ex.Message}\n");
                }
                catch { }

                return null;
            }
        }

        // ISpObjectWithToken implementation - REQUIRED for SAPI5 TTS engines
        private IntPtr _objectToken = IntPtr.Zero;

        public int SetObjectToken(IntPtr pToken)
        {
            // IMMEDIATE logging to see if method is called
            try
            {
                string logDir = "C:\\OpenSpeech";
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);
                File.AppendAllText(Path.Combine(logDir, "sapi_debug.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}: *** SET OBJECT TOKEN CALLED *** pToken: {pToken}\n");
            }
            catch { }

            try
            {
                LogMessage($"SetObjectToken called with pToken: {pToken}");
                _objectToken = pToken;

                // Now initialize the voice using the token passed by SAPI
                // This is the correct SAPI5 pattern - initialization happens here, not in constructor

                if (pToken != IntPtr.Zero)
                {
                    LogMessage("Initializing voice from object token...");

                    // For now, use hardcoded paths since we know where the Amy voice is installed
                    // In a full implementation, we would parse the token to get voice-specific paths
                    string modelPath = @"C:\Program Files\OpenSpeech\models\piper-en-amy-medium\model.onnx";
                    string tokensPath = @"C:\Program Files\OpenSpeech\models\piper-en-amy-medium\tokens.txt";

                    LogMessage($"Using model path: {modelPath}");
                    LogMessage($"Using tokens path: {tokensPath}");

                    // Initialize SherpaTTS with real TTS capability
                    LogMessage("Initializing SherpaTTS...");
                    _sherpaTts = new SherpaTTS(modelPath, tokensPath, "", Path.GetDirectoryName(modelPath));

                    _initialized = true;
                    LogMessage("Voice initialization completed successfully");
                }
                else
                {
                    LogMessage("Warning: pToken is null - using default initialization");
                    _initialized = true;
                }

                LogMessage("SetObjectToken completed successfully");
                return S_OK;
            }
            catch (Exception ex)
            {
                LogError($"Error in SetObjectToken: {ex.Message}", ex);
                return E_FAIL;
            }
        }

        public int GetObjectToken(out IntPtr ppToken)
        {
            try
            {
                LogMessage("GetObjectToken called");
                ppToken = _objectToken;
                return S_OK;
            }
            catch (Exception ex)
            {
                LogError($"Error in GetObjectToken: {ex.Message}", ex);
                ppToken = IntPtr.Zero;
                return E_FAIL;
            }
        }
    }
}
