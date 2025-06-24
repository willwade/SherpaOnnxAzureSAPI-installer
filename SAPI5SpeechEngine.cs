using SensorySoftware.Speech.SpeechManager;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Speech.Synthesis;
using System.Threading;
using System.Threading.Tasks;

namespace SensorySoftware.Speech.SpeechEngines
{
    class Sapi5SpeechEngine : ISpeechEngine
	{
        private const string _engineId = "SAPI5";
		private const int _speakingSemaphoreTimeout = 500; // milliseconds
		
		private SpeechSynthesizer _synthesizer;
		private SpeechChunkQueue _chunkQueue;
		private SemaphoreSlim _speakingSemaphore;
		private object _voiceSynthesizer;
		private int? _deviceId;

        public string Id { get { return _engineId; } }
		public int SortOrder { get { return 4; } }
		public string Name { get { return "SAPI5"; } }

		public IEnumerable<SpeechEngineCapabilities> Capabilities
		{
			get
			{
				yield return SpeechEngineCapabilities.ChangeSpeed;
				yield return SpeechEngineCapabilities.ChangePitch;
				yield return SpeechEngineCapabilities.ForceWordEvents;
				yield return SpeechEngineCapabilities.CreateWav;
				yield return SpeechEngineCapabilities.ChangeDevice;
			}
		}

		public IEnumerable<SpeechEngineVoice> GetVoices(string sensoryFolder)
		{
			_chunkQueue = new SpeechChunkQueue(Name);
			_speakingSemaphore = new SemaphoreSlim(1, 1);

			_synthesizer = new SpeechSynthesizer();
			_synthesizer.SpeakCompleted += OnSpeakCompleted;
            _synthesizer.SpeakProgress += OnSpeakProgress;

            // Cache this value for setting output device later... horrible that it's private!
            _voiceSynthesizer = _synthesizer.GetType().GetProperty("VoiceSynthesizer", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(_synthesizer, null);
			
			foreach (var voiceInfo in _synthesizer.GetInstalledVoices().Where(v => v.Enabled).Select(v => v.VoiceInfo))
                yield return new SpeechEngineVoice(this,
                    string.Empty,
                    voiceInfo.Name,
                    Name + " " + voiceInfo.Description,
                    voiceInfo.Id + (char)2 + voiceInfo.Name,
                    voiceInfo.Id,
                    voiceInfo.Culture.Name,
                    (VoiceGender)voiceInfo.Gender,
                    (VoiceAge)voiceInfo.Age);
		}

        private void OnSpeakProgress(object sender, SpeakProgressEventArgs e)
        {
            _chunkQueue.OnCharacterPosition(e.CharacterPosition + 1 - SSML.HeaderLength); // charPos is zerobased
        }

        private void OnSpeakCompleted(object sender, SpeakCompletedEventArgs e)
		{
			try
			{
				_chunkQueue.OnEnded();
				_speakingSemaphore.Release();
			}
			catch (Exception ex)
			{
				Logger.Instance.Error(ex.ToString());
			}
		}

		public Task Speak(SpeechChunk speechChunk)
		{
			_chunkQueue.Enqueue(speechChunk);

			try
			{
				var voiceId = speechChunk.Voice.InternalId.Split((char)2);
				if (voiceId[0] != _synthesizer.Voice.Id)
				{
					Logger.Instance.InfoFormat("Loading {0} voice '{1}' ({2})", Name, voiceId[0], voiceId[1]);
					_synthesizer.SelectVoice(voiceId[1]);

					// IVONA Welsh and Welsh English voices have duplicate names - the SelectVoice method should take the ID but Microsoft messed this up.
					// Load alternate matching voices until we find the one we want...
					int alternate = 1;
					while (_synthesizer.Voice.Id != voiceId[0])
					{
						Logger.Instance.InfoFormat("Loaded {0} voice '{1}', trying ({2}, {3}, {4}, {5})...", Name, _synthesizer.Voice.Id, speechChunk.Voice.Gender, speechChunk.Voice.Age, alternate, speechChunk.Voice.LanguageCode);
                        _synthesizer.SelectVoiceByHints((System.Speech.Synthesis.VoiceGender)speechChunk.Voice.Gender, (System.Speech.Synthesis.VoiceAge)speechChunk.Voice.Age, alternate++, new CultureInfo(speechChunk.Voice.LanguageCode));

						if (alternate > _synthesizer.GetInstalledVoices().Count)
							throw new ArgumentOutOfRangeException("Unable to match voice " + voiceId[0]);
					}
				}

                var xml = SSML.Create(speechChunk.Text, _synthesizer.Voice.Culture.Name, speechChunk.Pitch, speechChunk.Speed, isEscaped: false);

				// Use this semaphore to ensure the stop event has been fired before we continued, as SAPI5 is asynchronous
				if (!_speakingSemaphore.Wait(_speakingSemaphoreTimeout))
					Logger.Instance.Warn(string.Format("{0} speak wait timeout ({1}ms) exceeded", Name, _speakingSemaphoreTimeout));

				if (speechChunk.WavStream != null)
				{
					_deviceId = null;
                    // , new SpeechAudioFormatInfo(speechChunk.WavSampleRate, AudioBitsPerSample.Sixteen, AudioChannel.Mono)
                    _synthesizer.SetOutputToWaveStream(speechChunk.WavStream);
					_synthesizer.SpeakSsml(xml);
					_synthesizer.SetOutputToNull();
					_chunkQueue.OnEnded();
                }
				else
				{
					var deviceId = DeviceMapper.GetOutputWaveDeviceId(speechChunk);
					if (deviceId != _deviceId)
					{
						Logger.Instance.InfoFormat("Setting {0} device '{1}'", Name, deviceId);
						_synthesizer.SetOutputToDefaultAudioDevice();

						// HACK - Use reflection hack to set output device - why so private?!
						var waveOut = _voiceSynthesizer.GetType().GetField("_waveOut", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(_voiceSynthesizer);
						waveOut.GetType().GetField("_curDevice", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(waveOut, deviceId);
						_deviceId = deviceId;
					}

					_synthesizer.SpeakSsmlAsync(xml);
                }

                return Task.FromResult(true);
            }
			catch
			{
				_chunkQueue.OnEnded();
				throw;
			}
		}

		public void StopSpeaking()
		{
			_synthesizer.SpeakAsyncCancelAll();
		}

		public void Dispose()
		{
			if (_synthesizer != null)
				_synthesizer.Dispose();
		}
	}
}