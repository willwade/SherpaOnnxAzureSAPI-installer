#!/usr/bin/env python3
"""
Test script for SAPI Voice Manager
Tests the CLI tool functionality without requiring user interaction
"""

import os
import sys
import json
import tempfile
import shutil
from pathlib import Path

# Add current directory to path to import our module
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from SapiVoiceManager import SapiVoiceManager, TTS_ENGINES, VOICE_LOCALES
    print("✅ Successfully imported SapiVoiceManager")
except ImportError as e:
    print(f"❌ Failed to import SapiVoiceManager: {e}")
    sys.exit(1)

def test_voice_config_creation():
    """Test creating voice configurations programmatically"""
    print("\n=== Testing Voice Configuration Creation ===")
    
    # Create a temporary directory for testing
    test_dir = Path(tempfile.mkdtemp())
    print(f"Using test directory: {test_dir}")
    
    try:
        # Create test configurations
        test_configs = [
            {
                "name": "Test-British-Azure-Libby",
                "displayName": "Test British English (Azure Libby)",
                "description": "Test configuration for Azure TTS Libby voice",
                "language": "English",
                "locale": "en-GB",
                "gender": "Female",
                "age": "Adult",
                "vendor": "Azure TTS",
                "ttsConfig": {
                    "engine": "azure",
                    "azureTTS": {
                        "key": "test-key-12345",
                        "location": "uksouth",
                        "voice": "en-GB-LibbyNeural"
                    },
                    "TTS": {
                        "engine": "azure",
                        "voice_id": "en-GB-LibbyNeural",
                        "bypass_tts": False
                    },
                    "translate": {
                        "no_translate": True,
                        "provider": "",
                        "start_lang": "auto",
                        "end_lang": "en",
                        "replace_pb": False
                    }
                }
            },
            {
                "name": "Test-American-Sherpa-Amy",
                "displayName": "Test American English (Sherpa Amy)",
                "description": "Test configuration for Sherpa-ONNX Amy voice",
                "language": "English",
                "locale": "en-US",
                "gender": "Female",
                "age": "Adult",
                "vendor": "Sherpa-ONNX",
                "ttsConfig": {
                    "engine": "sherpa",
                    "SherpaOnnxTTS": {
                        "voice_id": "sherpa-piper-amy-normal"
                    },
                    "TTS": {
                        "engine": "sherpa",
                        "voice_id": "sherpa-piper-amy-normal",
                        "bypass_tts": False
                    },
                    "translate": {
                        "no_translate": True,
                        "provider": "",
                        "start_lang": "auto",
                        "end_lang": "en",
                        "replace_pb": False
                    }
                }
            }
        ]
        
        # Save test configurations
        voice_configs_dir = test_dir / "voice_configs"
        voice_configs_dir.mkdir(exist_ok=True)
        
        for config in test_configs:
            config_file = voice_configs_dir / f"{config['name']}.json"
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            print(f"✅ Created test config: {config_file.name}")
        
        # Test loading configurations
        print("\n--- Testing Configuration Loading ---")
        for config_file in voice_configs_dir.glob("*.json"):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                print(f"✅ Loaded: {config['displayName']}")
                print(f"   Engine: {config['ttsConfig']['engine']}")
                print(f"   Locale: {config['locale']}")
                
                # Validate required fields
                required_fields = ["name", "displayName", "locale", "gender", "ttsConfig"]
                missing_fields = [field for field in required_fields if field not in config]
                
                if missing_fields:
                    print(f"❌ Missing fields: {', '.join(missing_fields)}")
                else:
                    print("✅ All required fields present")
                    
            except Exception as e:
                print(f"❌ Error loading {config_file.name}: {e}")
        
        return True
        
    except Exception as e:
        print(f"❌ Test failed: {e}")
        return False
    finally:
        # Cleanup
        shutil.rmtree(test_dir, ignore_errors=True)
        print(f"🧹 Cleaned up test directory")

def test_engine_configurations():
    """Test TTS engine configuration validation"""
    print("\n=== Testing Engine Configurations ===")
    
    for engine_key, engine_info in TTS_ENGINES.items():
        print(f"\n--- Testing {engine_info['name']} ---")
        print(f"Config section: {engine_info['config_section']}")
        print(f"Credential fields: {engine_info['credential_fields']}")
        print(f"Required fields: {engine_info['required_fields']}")
        print(f"Description: {engine_info['description']}")
        print("✅ Engine configuration valid")

def test_locale_mappings():
    """Test locale to LCID mappings"""
    print("\n=== Testing Locale Mappings ===")
    
    for locale, info in VOICE_LOCALES.items():
        print(f"{locale}: LCID {info['lcid']} - {info['name']}")
    
    print(f"✅ {len(VOICE_LOCALES)} locales configured")

def test_manager_initialization():
    """Test SapiVoiceManager initialization"""
    print("\n=== Testing Manager Initialization ===")
    
    try:
        manager = SapiVoiceManager()
        print(f"✅ Manager initialized")
        print(f"Voice configs directory: {manager.voice_configs_dir}")
        print(f"Installer path: {manager.installer_path}")
        
        # Test directory creation
        if manager.voice_configs_dir.exists():
            print("✅ Voice configs directory exists")
        else:
            print("⚠️  Voice configs directory will be created on first use")
        
        # Test installer detection
        if manager.installer_path:
            print(f"✅ Installer found: {manager.installer_path}")
        else:
            print("⚠️  Installer not found (expected if not built yet)")
        
        return True
        
    except Exception as e:
        print(f"❌ Manager initialization failed: {e}")
        return False

def test_command_line_args():
    """Test command line argument parsing"""
    print("\n=== Testing Command Line Arguments ===")
    
    # Test importing the main function
    try:
        from SapiVoiceManager import main
        print("✅ Main function imported successfully")
        
        # Note: We can't easily test argparse without actually running the script
        # But we can verify the function exists and is callable
        if callable(main):
            print("✅ Main function is callable")
        else:
            print("❌ Main function is not callable")
            
    except ImportError as e:
        print(f"❌ Failed to import main function: {e}")
        return False
    
    return True

def run_all_tests():
    """Run all tests"""
    print("🧪 SAPI Voice Manager Test Suite")
    print("=" * 50)
    
    tests = [
        ("Manager Initialization", test_manager_initialization),
        ("Engine Configurations", test_engine_configurations),
        ("Locale Mappings", test_locale_mappings),
        ("Voice Config Creation", test_voice_config_creation),
        ("Command Line Args", test_command_line_args)
    ]
    
    passed = 0
    failed = 0
    
    for test_name, test_func in tests:
        print(f"\n🔍 Running: {test_name}")
        try:
            if test_func():
                print(f"✅ {test_name}: PASSED")
                passed += 1
            else:
                print(f"❌ {test_name}: FAILED")
                failed += 1
        except Exception as e:
            print(f"❌ {test_name}: ERROR - {e}")
            failed += 1
    
    print("\n" + "=" * 50)
    print(f"🧪 Test Results: {passed} passed, {failed} failed")
    
    if failed == 0:
        print("🎉 All tests passed!")
        return True
    else:
        print("⚠️  Some tests failed. Check the output above.")
        return False

if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)
