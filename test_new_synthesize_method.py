#!/usr/bin/env python3
"""
Test script for the new synthesize() method in tts-wrapper v0.10.19

This script tests the new synthesize() method that provides silent audio synthesis
without playback, which should solve the issues we had with the old methods.
"""

import os
import sys
import logging

# Setup basic logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def test_synthesize_method():
    """Test the new synthesize() method"""
    print("🧪 Testing new synthesize() method from tts-wrapper v0.10.19")
    print("=" * 60)
    
    try:
        from tts_wrapper import SherpaOnnxTTS, MicrosoftTTS
        
        # Test 1: SherpaOnnx TTS
        print("\n📦 Test 1: SherpaOnnx TTS with synthesize() method")
        try:
            tts_sherpa = SherpaOnnxTTS()
            
            if hasattr(tts_sherpa, 'synthesize'):
                print("✅ synthesize() method is available")
                print("🎤 Testing synthesize() - should be SILENT...")
                
                audio_bytes = tts_sherpa.synthesize("Testing the new synthesize method with SherpaOnnx.")
                
                if audio_bytes:
                    print(f"✅ Success! Generated {len(audio_bytes)} bytes of audio data")
                    
                    # Check if it has WAV headers
                    if len(audio_bytes) >= 12 and audio_bytes[:4] == b'RIFF' and audio_bytes[8:12] == b'WAVE':
                        print("✅ Audio data already has WAV headers")
                    else:
                        print("ℹ️ Audio data is raw PCM (no WAV headers)")
                else:
                    print("❌ No audio data returned")
            else:
                print("❌ synthesize() method not available")
                
        except Exception as e:
            print(f"❌ SherpaOnnx test failed: {e}")
        
        # Test 2: Azure TTS (if configured)
        print("\n📦 Test 2: Azure TTS with synthesize() method")
        try:
            # Try to load Azure config
            import configparser
            config = configparser.ConfigParser()
            config_path = os.path.join(os.path.dirname(__file__), "settings.cfg")
            
            if os.path.exists(config_path):
                config.read(config_path)
                azure_key = config.get("azureTTS", "key", fallback="")
                azure_location = config.get("azureTTS", "location", fallback="uksouth")
                
                if azure_key:
                    print("✅ Azure TTS configuration found")
                    credentials = (azure_key, azure_location)
                    tts_azure = MicrosoftTTS(credentials=credentials)
                    tts_azure.voice = "en-GB-LibbyNeural"
                    
                    if hasattr(tts_azure, 'synthesize'):
                        print("✅ synthesize() method is available")
                        print("🎤 Testing synthesize() - should be SILENT...")
                        
                        audio_bytes = tts_azure.synthesize("Testing the new synthesize method with Azure TTS.", voice_id="en-GB-LibbyNeural")
                        
                        if audio_bytes:
                            print(f"✅ Success! Generated {len(audio_bytes)} bytes of audio data")
                            
                            # Check if it has WAV headers
                            if len(audio_bytes) >= 12 and audio_bytes[:4] == b'RIFF' and audio_bytes[8:12] == b'WAVE':
                                print("✅ Audio data already has WAV headers")
                            else:
                                print("ℹ️ Audio data is raw PCM (no WAV headers)")
                        else:
                            print("❌ No audio data returned")
                    else:
                        print("❌ synthesize() method not available")
                else:
                    print("⚠️ Azure TTS key not configured, skipping test")
            else:
                print("⚠️ settings.cfg not found, skipping Azure test")
                
        except Exception as e:
            print(f"❌ Azure TTS test failed: {e}")
        
        # Test 3: Streaming version
        print("\n📦 Test 3: Testing synthesize() with streaming=True")
        try:
            tts_sherpa = SherpaOnnxTTS()
            
            if hasattr(tts_sherpa, 'synthesize'):
                print("🎤 Testing synthesize(streaming=True) - should be SILENT...")
                
                audio_stream = tts_sherpa.synthesize("Testing streaming synthesis.", streaming=True)
                
                if audio_stream:
                    print("✅ Got audio stream")
                    total_bytes = 0
                    chunk_count = 0
                    
                    for chunk in audio_stream:
                        if chunk:
                            total_bytes += len(chunk)
                            chunk_count += 1
                    
                    print(f"✅ Processed {chunk_count} chunks, total {total_bytes} bytes")
                else:
                    print("❌ No audio stream returned")
            else:
                print("❌ synthesize() method not available")
                
        except Exception as e:
            print(f"❌ Streaming test failed: {e}")
            
    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("Make sure tts-wrapper v0.10.19 is installed")
    
    print("\n" + "=" * 60)
    print("🏁 Test completed!")

if __name__ == "__main__":
    test_synthesize_method()
