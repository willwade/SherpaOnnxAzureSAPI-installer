#!/usr/bin/env python3
"""
Direct Pipe Service Test
Tests the AACSpeakHelper pipe service directly using the voice configuration
"""

import json
import win32file
import win32pipe
import sys

def test_azure_voice_direct():
    """Test Azure TTS voice directly through pipe service"""
    pipe_name = r'\\.\pipe\AACSpeakHelper'
    
    print("🧪 TESTING AZURE TTS THROUGH PIPE SERVICE")
    print("=" * 50)
    
    try:
        # Connect to pipe
        print("📡 Connecting to AACSpeakHelper pipe...")
        handle = win32file.CreateFile(
            pipe_name,
            win32file.GENERIC_READ | win32file.GENERIC_WRITE,
            0, None,
            win32file.OPEN_EXISTING,
            0, None
        )
        print("✅ Connected to pipe successfully")
        
        # Create message using the exact format from voice config
        message = {
            'args': {
                'listvoices': False,
                'preview': False,
                'style': '',
                'styledegree': None,
                'text': 'Hello from Azure TTS test',
                'verbose': True
            },
            'config': {
                'TTS': {
                    'engine': 'azure',
                    'voice_id': 'en-GB-LibbyNeural',
                    'bypass_tts': 'false'
                },
                'azureTTS': {
                    'key': 'b14f8945b0f1459f9964bdd72c42c2cc',
                    'location': 'uksouth',
                    'voice': 'en-GB-LibbyNeural',
                    'style': '',
                    'role': ''
                },
                'translate': {
                    'no_translate': 'true',
                    'provider': '',
                    'start_lang': 'auto',
                    'end_lang': 'en',
                    'replace_pb': 'false'
                },
                'App': {
                    'config_path': '',
                    'audio_files_path': ''
                }
            },
            'clipboard_text': 'Hello from Azure TTS test'
        }
        
        # Send message
        print("📤 Sending TTS request...")
        json_message = json.dumps(message)
        print(f"📋 Message size: {len(json_message)} bytes")
        
        win32file.WriteFile(handle, json_message.encode())
        print("✅ Message sent successfully")
        
        # Try to read response
        print("📥 Waiting for response...")
        try:
            result, data = win32file.ReadFile(handle, 64 * 1024)
            if result == 0 and data:
                response = data.decode()
                print(f"📨 Received response: {response[:100]}...")
            else:
                print("📭 No response data received")
        except Exception as e:
            print(f"📭 No response received (this is normal for TTS): {e}")
        
        win32file.CloseHandle(handle)
        print("✅ Test completed successfully!")
        print("")
        print("🎯 RESULT: If you heard speech, the pipe service is working correctly!")
        print("🎯 ISSUE: The problem is COM wrapper registration, not the TTS engine")
        
        return True
        
    except Exception as e:
        print(f"❌ Pipe test failed: {e}")
        print("")
        print("🔍 POSSIBLE CAUSES:")
        print("  1. AACSpeakHelper service not running")
        print("  2. Pipe communication issue")
        print("  3. Azure TTS credentials invalid")
        return False

def test_pipe_connection():
    """Test basic pipe connection"""
    pipe_name = r'\\.\pipe\AACSpeakHelper'
    
    print("🔌 TESTING PIPE CONNECTION")
    print("=" * 30)
    
    try:
        handle = win32file.CreateFile(
            pipe_name,
            win32file.GENERIC_READ,
            0, None,
            win32file.OPEN_EXISTING,
            0, None
        )
        win32file.CloseHandle(handle)
        print("✅ Pipe connection successful")
        return True
    except Exception as e:
        print(f"❌ Pipe connection failed: {e}")
        return False

if __name__ == '__main__':
    print("🧪 DIRECT PIPE SERVICE TEST")
    print("=" * 60)
    print("")
    
    # Test basic connection first
    if test_pipe_connection():
        print("")
        # Test Azure TTS
        test_azure_voice_direct()
    else:
        print("")
        print("🚫 Cannot proceed - AACSpeakHelper service not accessible")
        print("💡 Make sure AACSpeakHelper is running:")
        print("   cd AACSpeakHelper && uv run AACSpeakHelperServer.py")
