#!/usr/bin/env python3
"""
Test script that exactly mimics AACSpeakHelper's TTS calling method.
This will help us identify why AACSpeakHelper was playing audio.
"""

import sys
import os
from pathlib import Path

def test_aac_speak_helper_method():
    """Test the exact method AACSpeakHelper uses"""
    print("🧪 Testing AACSpeakHelper's exact TTS method")
    print("=" * 60)
    
    try:
        from tts_wrapper import SherpaOnnxTTS
        
        print("📦 Creating SherpaOnnx TTS engine (AACSpeakHelper style)...")
        
        # Create engine exactly like AACSpeakHelper does
        tts = SherpaOnnxTTS()
        
        print("🎤 Testing speak_streamed() exactly like AACSpeakHelper...")
        print("   Parameters: return_bytes=True, wait_for_completion=False")
        print("   Listen carefully - you should hear NOTHING!")
        
        # Call exactly like AACSpeakHelper does (line 204-209 in AACSpeakHelperServer.py)
        audio_bytes = tts.speak_streamed(
            "This mimics AACSpeakHelper exactly. You should NOT hear this.",
            voice_id=None,
            return_bytes=True,
            wait_for_completion=False  # This is different from our previous test!
        )
        
        if audio_bytes:
            print(f"✅ Got {len(audio_bytes)} bytes of audio data")
            print("❓ Did you hear audio? (You shouldn't have!)")
        else:
            print("❌ No audio bytes returned")
            
    except Exception as e:
        print(f"❌ AACSpeakHelper method test failed: {e}")
        import traceback
        traceback.print_exc()

def test_with_wait_for_completion():
    """Test with wait_for_completion=True vs False"""
    print("\n🧪 Testing wait_for_completion parameter difference")
    print("=" * 60)
    
    try:
        from tts_wrapper import SherpaOnnxTTS
        
        print("📦 Test 1: wait_for_completion=True")
        tts1 = SherpaOnnxTTS()
        print("   Listen - should be SILENT...")
        audio_bytes1 = tts1.speak_streamed(
            "Test with wait for completion true.",
            return_bytes=True,
            wait_for_completion=True
        )
        print(f"   Result: {len(audio_bytes1) if audio_bytes1 else 0} bytes")
        
        print("\n📦 Test 2: wait_for_completion=False (AACSpeakHelper style)")
        tts2 = SherpaOnnxTTS()
        print("   Listen - should be SILENT...")
        audio_bytes2 = tts2.speak_streamed(
            "Test with wait for completion false.",
            return_bytes=True,
            wait_for_completion=False
        )
        print(f"   Result: {len(audio_bytes2) if audio_bytes2 else 0} bytes")
        
    except Exception as e:
        print(f"❌ wait_for_completion test failed: {e}")
        import traceback
        traceback.print_exc()

def test_volume_setting():
    """Test if volume setting affects audio playback"""
    print("\n🧪 Testing volume parameter")
    print("=" * 60)
    
    try:
        from tts_wrapper import SherpaOnnxTTS
        
        print("📦 Test with volume=100 (default)")
        tts1 = SherpaOnnxTTS()
        print("   Listen - should be SILENT...")
        audio_bytes1 = tts1.speak_streamed(
            "Test with volume one hundred.",
            return_bytes=True,
            wait_for_completion=False
        )
        print(f"   Result: {len(audio_bytes1) if audio_bytes1 else 0} bytes")
        
        print("\n📦 Test with volume=0")
        tts2 = SherpaOnnxTTS()
        print("   Listen - should be SILENT...")
        audio_bytes2 = tts2.speak_streamed(
            "Test with volume zero.",
            return_bytes=True,
            wait_for_completion=False,
            volume=0
        )
        print(f"   Result: {len(audio_bytes2) if audio_bytes2 else 0} bytes")
        
    except Exception as e:
        print(f"❌ Volume test failed: {e}")
        import traceback
        traceback.print_exc()

def main():
    print("🔬 AACSpeakHelper Method Analysis")
    print("=" * 60)
    print("This script tests the exact method AACSpeakHelper uses")
    print("to identify why it might be playing audio.")
    print()
    
    input("Press Enter to start AACSpeakHelper method test...")
    test_aac_speak_helper_method()
    
    input("\nPress Enter to test wait_for_completion parameter...")
    test_with_wait_for_completion()
    
    input("\nPress Enter to test volume parameter...")
    test_volume_setting()
    
    print("\n" + "=" * 60)
    print("🏁 Analysis Complete!")
    print("=" * 60)
    print("If you heard audio in any test, we've found the issue!")

if __name__ == "__main__":
    main()
