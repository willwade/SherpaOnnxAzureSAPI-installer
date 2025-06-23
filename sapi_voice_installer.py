#!/usr/bin/env python3
"""
SAPI Voice Installer for AACSpeakHelper
A unified Python installer for managing SAPI voices that communicate with AACSpeakHelper.

This replaces the .NET installer and provides a simpler Python + C++ solution.
"""

import os
import sys
import json
import argparse
import winreg
import subprocess
import configparser
import ctypes
import time
from pathlib import Path

# TTS engine definitions (moved from removed cli_config_creator.py)
TTS_ENGINES = {
    "sherpa": {
        "name": "Sherpa-ONNX",
        "config_section": "SherpaOnnxTTS",
        "credential_fields": [],
        "voice_list": {"English": "en_GB-jenny_dioco-medium"},
    },
    "azure": {
        "name": "Azure TTS",
        "config_section": "azureTTS",
        "credential_fields": ["key", "location"],
        "voice_list": {"English (US)": "en-US-JennyNeural", "English (UK)": "en-GB-LibbyNeural"},
    },
    "google": {
        "name": "Google TTS",
        "config_section": "googleTTS",
        "credential_fields": ["creds"],
        "voice_list": {"English (US)": "en-US-Wavenet-C"},
    },
    "elevenlabs": {
        "name": "ElevenLabs",
        "config_section": "ElevenLabsTTS",
        "credential_fields": ["api_key"],
        "voice_list": {},
    },
    "playht": {
        "name": "PlayHT",
        "config_section": "PlayHTTTS",
        "credential_fields": ["api_key", "user_id"],
        "voice_list": {},
    },
    "polly": {
        "name": "AWS Polly",
        "config_section": "PollyTTS",
        "credential_fields": ["region", "aws_key_id", "aws_access_key"],
        "voice_list": {},
    },
    "watson": {
        "name": "IBM Watson",
        "config_section": "WatsonTTS",
        "credential_fields": ["api_key", "region", "instance_id"],
        "voice_list": {},
    },
    "openai": {
        "name": "OpenAI TTS",
        "config_section": "OpenAITTS",
        "credential_fields": ["api_key"],
        "voice_list": {},
    },
    "witai": {
        "name": "Wit.AI",
        "config_section": "WitAiTTS",
        "credential_fields": ["token"],
        "voice_list": {},
    },
}

# Constants
NATIVE_TTS_WRAPPER_CLSID = "{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
VOICE_CONFIGS_DIR = Path("C:/Program Files/OpenAssistive/OpenSpeech/voice_configs")
AACSPEAKHELPER_CONFIG_PATH = Path("AACSpeakHelper/settings.cfg")
SAPI_REGISTRY_PATH = r"SOFTWARE\Microsoft\SPEECH\Voices\Tokens"

def is_admin():
    """Check if the current process is running with administrator privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def warn_admin_required():
    """Warn user that administrator privileges are required"""
    if not is_admin():
        print("\n⚠️  ADMINISTRATOR PRIVILEGES REQUIRED")
        print("=" * 50)
        print("This installer needs to register COM components with Windows.")
        print("COM registration requires administrator privileges.")
        print("")
        print("Please:")
        print("1. Close this installer")
        print("2. Right-click on PowerShell or Command Prompt")
        print("3. Select 'Run as administrator'")
        print("4. Navigate to this directory and run the installer again")
        print("")
        print("Or run directly as admin:")
        print(f"   uv run {sys.argv[0]}")
        print("=" * 50)
        return False
    return True

def get_config_dir():
    """Get the configuration directory path"""
    if getattr(sys, "frozen", False):
        config_dir = os.path.join(
            os.path.expanduser("~"),
            "AppData",
            "Roaming",
            "Ace Centre",
            "AACSpeakHelper",
        )
    else:
        config_dir = os.path.dirname(os.path.abspath(__file__))

    os.makedirs(config_dir, exist_ok=True)
    return config_dir

def load_config(custom_config_path=None):
    """Load configuration from file"""
    config = configparser.ConfigParser()

    if custom_config_path and os.path.exists(custom_config_path):
        config_path = custom_config_path
    else:
        config_path = os.path.join(get_config_dir(), "settings.cfg")

    if os.path.exists(config_path):
        config.read(config_path)
        print(f"Configuration loaded from {config_path}")
    else:
        print("No existing configuration found. Using default settings.")
        # Create minimal default config for voice discovery
        config["TTS"] = {"engine": "SherpaOnnxTTS"}
        config["azureTTS"] = {"key": "", "location": "uksouth"}
        config["SherpaOnnxTTS"] = {"voice_id": "en_GB-jenny_dioco-medium"}

    return config, config_path

def normalize_voice_data(voices):
    """Normalize voice data from any TTS engine to a consistent format"""
    if not voices:
        return []

    voice_list = []
    for voice in voices:
        # Handle both dict and object formats
        if isinstance(voice, dict):
            normalized_voice = {
                'id': voice.get('id', voice.get('voice_id', '')),
                'name': voice.get('name', voice.get('display_name', '')),
                'language': voice.get('language', voice.get('locale', '')),
                'gender': voice.get('gender', 'Unknown')
            }
        else:
            normalized_voice = {
                'id': getattr(voice, 'id', getattr(voice, 'voice_id', '')),
                'name': getattr(voice, 'name', getattr(voice, 'display_name', '')),
                'language': getattr(voice, 'language', getattr(voice, 'locale', '')),
                'gender': getattr(voice, 'gender', 'Unknown')
            }
        voice_list.append(normalized_voice)
    return voice_list

def get_voices_from_engine(engine_key, config):
    """Get voices from the actual TTS engine using py3-tts-wrapper"""
    try:
        from tts_wrapper import (
            MicrosoftClient, MicrosoftTTS,
            SherpaOnnxClient, SherpaOnnxTTS,
            GoogleTransTTS,
            ElevenLabsClient, ElevenLabsTTS,
            PlayHTClient, PlayHTTTS,
            PollyClient, PollyTTS,
            WatsonClient, WatsonTTS,
            OpenAIClient,
            WitAiClient, WitAiTTS
        )

        engine_config = TTS_ENGINES[engine_key]
        section_name = engine_config["config_section"]

        if engine_key == "azure":
            # Get Azure TTS credentials
            key = config.get(section_name, "key", fallback="")
            location = config.get(section_name, "location", fallback="")
            if not key:
                key = os.getenv("AZURE_TTS_KEY", "")
            if not location:
                location = os.getenv("AZURE_TTS_LOCATION", "uksouth")

            if not key or not location:
                print("❌ Azure TTS credentials not configured. Please configure them first.")
                return None

            print(f"🔍 Testing Azure TTS credentials (region: {location})...")
            client = MicrosoftClient(credentials=(key, location))
            tts = MicrosoftTTS(client)
            voices = tts.get_voices()
            print(f"✅ Retrieved {len(voices)} voices from Azure TTS")
            return normalize_voice_data(voices)

        elif engine_key == "sherpa":
            # SherpaOnnx doesn't need credentials
            print("🔍 Getting SherpaOnnx voices...")
            client = SherpaOnnxClient()
            tts = SherpaOnnxTTS(client)
            voices = tts.get_voices()
            print(f"✅ Retrieved {len(voices) if voices else 0} voices from SherpaOnnx")
            return normalize_voice_data(voices)

        elif engine_key == "google_trans":
            # Google Trans TTS doesn't need credentials
            print("🔍 Getting GoogleTrans voices...")
            try:
                tts = GoogleTransTTS(voice_id="en")
                voices = tts.get_voices()
                print(f"✅ Retrieved {len(voices) if voices else 0} voices from GoogleTrans")
                return normalize_voice_data(voices)
            except Exception as e:
                print(f"❌ Error with GoogleTrans: {e}")
                return None

        elif engine_key == "elevenlabs":
            # Get ElevenLabs credentials
            api_key = config.get(section_name, "api_key", fallback="")
            if not api_key:
                api_key = os.getenv("ELEVENLABS_API_KEY", "")
            if not api_key:
                print("❌ ElevenLabs API key not configured. Please configure it first.")
                return None

            print("🔍 Testing ElevenLabs credentials...")
            client = ElevenLabsClient(credentials=(api_key,))
            tts = ElevenLabsTTS(client)
            voices = tts.get_voices()
            print(f"✅ Retrieved {len(voices)} voices from ElevenLabs")
            return normalize_voice_data(voices)

        elif engine_key == "playht":
            # Get PlayHT credentials
            api_key = config.get(section_name, "api_key", fallback="")
            user_id = config.get(section_name, "user_id", fallback="")
            if not api_key:
                api_key = os.getenv("PLAYHT_API_KEY", "")
            if not user_id:
                user_id = os.getenv("PLAYHT_USER_ID", "")
            if not api_key or not user_id:
                print("❌ PlayHT credentials not configured. Please configure them first.")
                return None

            print("🔍 Testing PlayHT credentials...")
            client = PlayHTClient(credentials=(api_key, user_id))
            tts = PlayHTTTS(client)
            voices = tts.get_voices()
            print(f"✅ Retrieved {len(voices)} voices from PlayHT")
            return normalize_voice_data(voices)

        elif engine_key == "polly":
            # Get Polly credentials
            region = config.get(section_name, "region", fallback="")
            aws_key_id = config.get(section_name, "aws_key_id", fallback="")
            aws_access_key = config.get(section_name, "aws_access_key", fallback="")
            if not region:
                region = os.getenv("POLLY_REGION", "us-east-1")
            if not aws_key_id:
                aws_key_id = os.getenv("POLLY_AWS_KEY_ID", "")
            if not aws_access_key:
                aws_access_key = os.getenv("POLLY_AWS_ACCESS_KEY", "")
            if not aws_key_id or not aws_access_key:
                print("❌ AWS Polly credentials not configured. Please configure them first.")
                return None

            print(f"🔍 Testing AWS Polly credentials (region: {region})...")
            client = PollyClient(credentials=(region, aws_key_id, aws_access_key))
            tts = PollyTTS(client)
            voices = tts.get_voices()
            print(f"✅ Retrieved {len(voices)} voices from AWS Polly")
            return normalize_voice_data(voices)

        elif engine_key == "watson":
            # Get Watson credentials
            api_key = config.get(section_name, "api_key", fallback="")
            region = config.get(section_name, "region", fallback="")
            instance_id = config.get(section_name, "instance_id", fallback="")
            if not api_key:
                api_key = os.getenv("WATSON_API_KEY", "")
            if not region:
                region = os.getenv("WATSON_REGION", "eu-gb")
            if not instance_id:
                instance_id = os.getenv("WATSON_INSTANCE_ID", "")
            if not api_key:
                print("❌ Watson credentials not configured. Please configure them first.")
                return None

            print(f"🔍 Testing Watson credentials (region: {region})...")
            client = WatsonClient(credentials=(api_key, region, instance_id))
            tts = WatsonTTS(client)
            voices = tts.get_voices()
            print(f"✅ Retrieved {len(voices)} voices from Watson")
            return normalize_voice_data(voices)

        elif engine_key == "openai":
            # Get OpenAI credentials
            api_key = config.get(section_name, "api_key", fallback="")
            if not api_key:
                api_key = os.getenv("OPENAI_API_KEY", "")
            if not api_key:
                print("❌ OpenAI API key not configured. Please configure it first.")
                return None

            print("🔍 Testing OpenAI credentials...")
            client = OpenAIClient(credentials=(api_key,))
            voices = client.get_voices()
            print(f"✅ Retrieved {len(voices)} voices from OpenAI")
            return normalize_voice_data(voices)

        elif engine_key == "witai":
            # Get WitAI credentials
            token = config.get(section_name, "token", fallback="")
            if not token:
                token = os.getenv("WITAI_TOKEN", "")
            if not token:
                print("❌ WitAI token not configured. Please configure it first.")
                return None

            print("🔍 Testing WitAI credentials...")
            client = WitAiClient(credentials=(token,))
            tts = WitAiTTS(client)
            voices = tts.get_voices()
            print(f"✅ Retrieved {len(voices)} voices from WitAI")
            return normalize_voice_data(voices)

        return None

    except Exception as e:
        print(f"❌ Error getting voices from {engine_key}: {e}")
        return None

def configure_engine_credentials(config, engine_key):
    """Configure credentials for a specific TTS engine"""
    engine_config = TTS_ENGINES.get(engine_key)
    if not engine_config:
        print(f"❌ Unknown engine: {engine_key}")
        return config

    section_name = engine_config["config_section"]
    credential_fields = engine_config["credential_fields"]

    if not credential_fields:
        print(f"✅ {engine_config['name']} doesn't require credentials")
        return config

    print(f"\n🔧 Configuring {engine_config['name']} credentials:")
    print("-" * 50)

    # Ensure section exists
    if not config.has_section(section_name):
        config.add_section(section_name)

    # Configure each credential field
    for field in credential_fields:
        current_value = config.get(section_name, field, fallback="")

        # Show current value (masked if it looks like a key)
        if current_value and any(keyword in field.lower() for keyword in ['key', 'secret', 'token', 'password']):
            display_value = f"{current_value[:8]}..." if len(current_value) > 8 else current_value
        else:
            display_value = current_value

        prompt = f"Enter {field}"
        if display_value:
            prompt += f" [current: {display_value}]"
        prompt += ": "

        new_value = input(prompt).strip()
        if new_value:
            config.set(section_name, field, new_value)
        elif not current_value:
            print(f"⚠️  Warning: {field} is required for {engine_config['name']}")

    return config

def test_engine_credentials(config, engine_key):
    """Test if engine credentials are working by fetching voices"""
    print(f"\n🧪 Testing {TTS_ENGINES[engine_key]['name']} credentials...")

    voices = get_voices_from_engine(engine_key, config)
    if voices:
        print(f"✅ Credentials working! Found {len(voices)} voices")
        return True
    else:
        print("❌ Credential test failed")
        return False

def get_credential_status(config, engine_key):
    """Get the credential status for an engine"""
    engine_config = TTS_ENGINES.get(engine_key)
    if not engine_config:
        return "❓ Unknown"

    credential_fields = engine_config["credential_fields"]
    if not credential_fields:
        return "✅ No credentials needed"

    section_name = engine_config["config_section"]
    missing_creds = []

    for field in credential_fields:
        value = config.get(section_name, field, fallback="")
        if not value:
            # Check environment variables as fallback
            env_var = f"{engine_key.upper()}_{field.upper()}"
            if not os.getenv(env_var, ""):
                missing_creds.append(field)

    if missing_creds:
        return f"❌ Missing: {', '.join(missing_creds)}"
    else:
        return "✅ Configured"

class SAPIVoiceInstaller:
    """Manages SAPI voice installation and registration"""
    
    def __init__(self):
        self.ensure_directories()
    
    def ensure_directories(self):
        """Create necessary directories"""
        try:
            VOICE_CONFIGS_DIR.mkdir(parents=True, exist_ok=True)
            print(f"✅ Voice configs directory: {VOICE_CONFIGS_DIR}")
        except Exception as e:
            print(f"❌ Failed to create directories: {e}")
            sys.exit(1)
    
    def register_com_wrapper(self, unattended_mode=False):
        """Register both x86 and x64 C++ COM wrapper DLLs with sophisticated error handling"""
        dll_path_x64 = Path("NativeTTSWrapper/x64/Release/NativeTTSWrapper.dll")
        dll_path_x86 = Path("NativeTTSWrapper/Win32/Release/NativeTTSWrapper.dll")

        # Check if running in unattended mode (CI/automated)
        if not unattended_mode:
            unattended_mode = (
                os.getenv("CI") == "true" or
                os.getenv("GITHUB_ACTIONS") == "true" or
                "--unattended" in sys.argv
            )

        if unattended_mode:
            print("[CI MODE] Running COM registration in unattended mode")
        else:
            print("[INTERACTIVE] Running COM registration in interactive mode")

        # Check if at least one DLL exists
        if not dll_path_x64.exists() and not dll_path_x86.exists():
            print(f"❌ No COM wrapper DLLs found:")
            print(f"   x64: {dll_path_x64}")
            print(f"   x86: {dll_path_x86}")
            print("   Please build the C++ project first")
            return False

        # Check for administrator privileges (skip in CI)
        if not unattended_mode and not is_admin():
            print("❌ COM registration requires administrator privileges")
            print("   Please run as administrator")
            return False
        elif unattended_mode:
            print("[CI MODE] Skipping administrator check in unattended mode")
        else:
            print("[OK] Running as Administrator")

        # Register both architectures
        success_count = 0
        total_count = 0

        for dll_path, arch_name in [(dll_path_x64, "x64"), (dll_path_x86, "x86")]:
            if not dll_path.exists():
                print(f"[INFO] Skipping {arch_name} registration - DLL not found: {dll_path}")
                continue

            total_count += 1
            print(f"[INFO] Registering {arch_name} COM wrapper: {dll_path}")

            try:
                # First registration attempt
                result = subprocess.run([
                    "regsvr32", "/s", str(dll_path.absolute())
                ], capture_output=True)

                if result.returncode == 0:
                    print(f"[OK] {arch_name} COM wrapper registered successfully!")
                    success_count += 1
                else:
                    print(f"[WARNING] {arch_name} registration failed (exit code: {result.returncode})")
                    print(f"[INFO] Attempting cleanup and retry for {arch_name}...")

                    # Clean up any stale registrations
                    cleanup_success = self._cleanup_com_registration()
                    if cleanup_success:
                        print(f"[OK] Cleanup completed successfully for {arch_name}")
                    else:
                        print(f"[WARNING] Cleanup had some issues for {arch_name}, but continuing...")

                    # Wait a moment for cleanup to complete
                    time.sleep(2)

                    # Retry registration
                    retry_result = subprocess.run([
                        "regsvr32", "/s", str(dll_path.absolute())
                    ], capture_output=True)

                    if retry_result.returncode == 0:
                        print(f"[OK] {arch_name} COM wrapper registered successfully after cleanup!")
                        success_count += 1
                    else:
                        print(f"[ERROR] {arch_name} registration failed even after cleanup (exit code: {retry_result.returncode})")

                        # Try to get more detailed error information
                        if retry_result.stderr:
                            error_msg = retry_result.stderr.decode('utf-8', errors='ignore').strip()
                            if error_msg:
                                print(f"[ERROR] {arch_name} error details: {error_msg}")

            except Exception as e:
                print(f"❌ Failed to register {arch_name} COM wrapper: {e}")

        # Summary
        if success_count == 0:
            print("❌ Failed to register any COM wrappers")
            print("[INFO] This might be due to:")
            print("        - Missing Visual C++ Redistributables")
            print("        - Antivirus blocking the registration")
            print("        - Insufficient permissions")
            print("        - DLL dependencies not found")
            if not unattended_mode:
                print("        Please restart your computer and try again")
            return False
        elif success_count == total_count:
            print(f"✅ Successfully registered all {success_count} COM wrappers!")
            print("   SAPI voices should now work in both 32-bit and 64-bit applications.")
            return True
        else:
            print(f"⚠️ Partially successful: {success_count}/{total_count} COM wrappers registered")
            print("   Some SAPI applications may not see the voices.")
            return True  # Partial success is still success

    def _verify_com_registration(self):
        """Verify COM registration by checking registry"""
        try:
            # Check for our CLSID in registry
            clsid_path = r"SOFTWARE\Classes\CLSID\{4A8B9C2D-1E3F-4567-8901-234567890ABC}"
            result = subprocess.run([
                "reg", "query", f"HKEY_LOCAL_MACHINE\\{clsid_path}"
            ], capture_output=True, text=True)
            return result.returncode == 0
        except Exception:
            return False

    def _cleanup_com_registration(self):
        """Clean up stale COM registrations and files"""
        try:
            print("[INFO] Cleaning up stale COM registrations...")

            # Unregister any existing COM wrappers (both architectures)
            dll_paths = [
                Path("NativeTTSWrapper/x64/Release/NativeTTSWrapper.dll"),
                Path("NativeTTSWrapper/Win32/Release/NativeTTSWrapper.dll"),
                Path("C:/Program Files/SherpaOnnx Azure SAPI Bridge/NativeTTSWrapper.dll"),
                Path("C:/Program Files (x86)/SherpaOnnx Azure SAPI Bridge/NativeTTSWrapper.dll")
            ]

            for dll_path in dll_paths:
                if dll_path.exists():
                    try:
                        subprocess.run([
                            "regsvr32", "/u", "/s", str(dll_path.absolute())
                        ], capture_output=True)
                        print(f"[INFO] Unregistered: {dll_path}")
                    except Exception:
                        pass  # Ignore errors during cleanup

            # Remove CLSID entries
            clsid_entries = [
                r"HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{4A8B9C2D-1E3F-4567-8901-234567890ABC}",
                r"HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{4A8B9C2D-1E3F-4567-8901-234567890ABD}"
            ]

            for clsid in clsid_entries:
                try:
                    subprocess.run(["reg", "delete", clsid, "/f"], capture_output=True)
                except Exception:
                    pass  # Ignore errors during cleanup

            # Remove AACSpeakHelper voice tokens
            try:
                result = subprocess.run([
                    "reg", "query", r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens"
                ], capture_output=True, text=True)

                if result.returncode == 0:
                    for line in result.stdout.split('\n'):
                        if 'AACSpeakHelper' in line:
                            token_path = line.strip()
                            if token_path:
                                try:
                                    subprocess.run(["reg", "delete", token_path, "/f"], capture_output=True)
                                except Exception:
                                    pass
            except Exception:
                pass

            # Clean up log files
            log_files = [
                Path("C:/OpenSpeech/native_tts_debug.log"),
                Path("AACSpeakHelper").glob("*.log")
            ]

            for log_file in log_files:
                if isinstance(log_file, Path) and log_file.exists():
                    try:
                        log_file.unlink()
                        print(f"[INFO] Removed log file: {log_file}")
                    except Exception:
                        pass
                elif hasattr(log_file, '__iter__'):  # It's a glob result
                    for lf in log_file:
                        try:
                            lf.unlink()
                            print(f"[INFO] Removed log file: {lf}")
                        except Exception:
                            pass

            print("[OK] Cleanup completed")
            return True

        except Exception as e:
            print(f"[ERROR] Cleanup failed: {e}")
            return False

    def unregister_com_wrapper(self, unattended_mode=False):
        """Unregister both x86 and x64 C++ COM wrapper DLLs with sophisticated error handling"""
        dll_path_x64 = Path("NativeTTSWrapper/x64/Release/NativeTTSWrapper.dll")
        dll_path_x86 = Path("NativeTTSWrapper/Win32/Release/NativeTTSWrapper.dll")

        # Check if running in unattended mode (CI/automated)
        if not unattended_mode:
            unattended_mode = (
                os.getenv("CI") == "true" or
                os.getenv("GITHUB_ACTIONS") == "true" or
                "--unattended" in sys.argv
            )

        if unattended_mode:
            print("[CI MODE] Running COM unregistration in unattended mode")
        else:
            print("[INTERACTIVE] Running COM unregistration in interactive mode")

        # Check for administrator privileges (skip in CI)
        if not unattended_mode and not is_admin():
            print("❌ COM unregistration requires administrator privileges")
            print("   Please run as administrator")
            return False
        elif unattended_mode:
            print("[CI MODE] Skipping administrator check in unattended mode")
        else:
            print("[OK] Running as Administrator")

        print("[INFO] Starting comprehensive COM unregistration...")

        # Always run cleanup regardless of DLL existence
        cleanup_success = self._cleanup_com_registration()

        # Check if any DLLs exist
        if not dll_path_x64.exists() and not dll_path_x86.exists():
            print("[INFO] No COM wrapper DLLs found")
            if cleanup_success:
                print("✅ COM wrappers unregistered (cleanup completed)")
                return True
            else:
                print("⚠️ DLLs not found, but cleanup had issues")
                return True  # Still consider it successful if DLLs don't exist

        # Unregister both architectures
        success_count = 0
        total_count = 0

        for dll_path, arch_name in [(dll_path_x64, "x64"), (dll_path_x86, "x86")]:
            if not dll_path.exists():
                print(f"[INFO] Skipping {arch_name} unregistration - DLL not found: {dll_path}")
                continue

            total_count += 1
            print(f"[INFO] Unregistering {arch_name} COM wrapper: {dll_path}")

            try:
                # Unregister the specific DLL
                result = subprocess.run([
                    "regsvr32", "/s", "/u", str(dll_path.absolute())
                ], capture_output=True)

                if result.returncode == 0:
                    print(f"✅ {arch_name} COM wrapper unregistered successfully")
                    success_count += 1
                else:
                    print(f"[WARNING] {arch_name} unregistration returned exit code: {result.returncode}")

            except Exception as e:
                print(f"[ERROR] Failed to unregister {arch_name} COM wrapper: {e}")

        # Summary
        if cleanup_success:
            print("✅ Registry cleanup completed")
            if success_count == total_count:
                print(f"✅ Successfully unregistered all {success_count} COM wrappers!")
            elif success_count > 0:
                print(f"⚠️ Partially successful: {success_count}/{total_count} COM wrappers unregistered")
            else:
                print("⚠️ DLL unregistration had issues, but registry cleanup completed")
            return True
        else:
            if success_count == total_count and total_count > 0:
                print(f"✅ Successfully unregistered all {success_count} COM wrappers!")
                return True
            else:
                print("❌ Both DLL unregistration and cleanup had issues")
                return False

    def create_aacspeakhelper_config(self, engine_key="sherpa", azure_key=None, azure_region="uksouth"):
        """Create or update AACSpeakHelper settings.cfg file"""
        import configparser

        config = configparser.ConfigParser()

        # App section
        config['App'] = {
            'collectstats': 'True'
        }

        # Translate section
        config['translate'] = {
            'no_translate': 'True',  # We want direct TTS, no translation
            'start_lang': 'en',
            'end_lang': 'en',
            'replace_pb': 'True',
            'provider': 'GoogleTranslator',
            'microsoft_translator_secret_key': '',
            'papago_translator_client_id': '',
            'papago_translator_secret_key': '',
            'my_memory_translator_secret_key': '',
            'email': '',
            'libre_translator_secret_key': '',
            'url': '',
            'deep_l_translator_secret_key': '',
            'deepl_pro': 'false',
            'region': '',
            'yandex_translator_secret_key': '',
            'qcri_translator_secret_key': '',
            'baidu_translator_appid': '',
            'baidu_translator_secret_key': ''
        }

        # TTS section - set based on engine preference
        if engine_key == "azure" and azure_key:
            config['TTS'] = {
                'engine': 'azureTTS',
                'bypass_tts': 'False',
                'save_audio_file': 'True',
                'rate': '0',
                'volume': '100',
                'voice_id': 'en-GB-LibbyNeural'
            }
            config['azureTTS'] = {
                'key': azure_key,
                'location': azure_region,
                'voice_id': 'en-GB-LibbyNeural'
            }
        else:
            # Default to SherpaOnnx
            config['TTS'] = {
                'engine': 'SherpaOnnxTTS',
                'bypass_tts': 'False',
                'save_audio_file': 'True',
                'rate': '0',
                'volume': '100',
                'voice_id': 'en_GB-jenny_dioco-medium'
            }

        # Engine-specific sections
        config['googleTTS'] = {
            'creds': '',
            'voice_id': 'en-US-Wavenet-C'
        }

        config['googleTransTTS'] = {
            'voice_id': ''
        }

        config['sapi5TTS'] = {
            'voice_id': 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\TTS_MS_EN-US_DAVID_11.0'
        }

        config['SherpaOnnxTTS'] = {
            'voice_id': 'en_GB-jenny_dioco-medium'
        }

        if azure_key:
            config['azureTTS'] = {
                'key': azure_key,
                'location': azure_region,
                'voice_id': 'en-GB-LibbyNeural'
            }

        config['appCache'] = {
            'threshold': '7'
        }

        # Save the config file
        try:
            AACSPEAKHELPER_CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(AACSPEAKHELPER_CONFIG_PATH, 'w', encoding='utf-8') as f:
                config.write(f)
            print(f"✅ AACSpeakHelper config created: {AACSPEAKHELPER_CONFIG_PATH}")
            return True
        except Exception as e:
            print(f"❌ Failed to create AACSpeakHelper config: {e}")
            return False
    
    def create_voice_config(self, voice_name, engine_key, voice_id, **kwargs):
        """Create a voice configuration file"""
        engine_config = TTS_ENGINES.get(engine_key)
        if not engine_config:
            print(f"❌ Unknown engine: {engine_key}")
            return None
        
        # Create voice configuration in AACSpeakHelper format
        config = {
            "name": voice_name,
            "displayName": kwargs.get("display_name", voice_name),
            "description": kwargs.get("description", f"{engine_config['name']} voice"),
            "language": kwargs.get("language", "English"),
            "locale": kwargs.get("locale", "en-GB"),
            "gender": kwargs.get("gender", "Female"),
            "age": kwargs.get("age", "Adult"),
            "vendor": engine_config["name"],
            "ttsConfig": {
                "text": "",
                "args": {
                    "engine": engine_key,
                    "voice": voice_id,
                    "rate": kwargs.get("rate", 0),
                    "volume": kwargs.get("volume", 100)
                }
            }
        }
        
        # Save configuration file
        config_path = VOICE_CONFIGS_DIR / f"{voice_name}.json"
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2)
            print(f"✅ Voice config created: {config_path}")
            return config_path
        except Exception as e:
            print(f"❌ Failed to create voice config: {e}")
            return None
    
    def get_lcid_from_locale(self, locale):
        """Convert locale string to Windows LCID"""
        locale_map = {
            "en-US": "409", "en-GB": "809", "en-AU": "c09", "en-CA": "1009",
            "fr-FR": "40c", "de-DE": "407", "es-ES": "c0a", "it-IT": "410",
            "pt-BR": "416", "ja-JP": "411", "ko-KR": "412", "zh-CN": "804",
            "zh-TW": "404", "ru-RU": "419", "ar-SA": "401", "hi-IN": "439"
        }
        return locale_map.get(locale, "409")  # Default to en-US
    
    def register_sapi_voice(self, voice_name, config_path, config):
        """Register a voice with Windows SAPI"""
        try:
            registry_path = f"{SAPI_REGISTRY_PATH}\\{voice_name}"
            lcid = self.get_lcid_from_locale(config["locale"])
            
            # Create voice registry key
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, registry_path) as voice_key:
                # Basic SAPI registration
                winreg.SetValueEx(voice_key, "", 0, winreg.REG_SZ, config["displayName"])
                winreg.SetValueEx(voice_key, lcid, 0, winreg.REG_SZ, config["displayName"])
                winreg.SetValueEx(voice_key, "CLSID", 0, winreg.REG_SZ, NATIVE_TTS_WRAPPER_CLSID)
                winreg.SetValueEx(voice_key, "ConfigPath", 0, winreg.REG_SZ, str(config_path))
                
                # Create Attributes subkey
                with winreg.CreateKey(voice_key, "Attributes") as attr_key:
                    winreg.SetValueEx(attr_key, "Language", 0, winreg.REG_SZ, lcid)
                    winreg.SetValueEx(attr_key, "Gender", 0, winreg.REG_SZ, config["gender"])
                    winreg.SetValueEx(attr_key, "Age", 0, winreg.REG_SZ, config["age"])
                    winreg.SetValueEx(attr_key, "Vendor", 0, winreg.REG_SZ, config["vendor"])
                    winreg.SetValueEx(attr_key, "Version", 0, winreg.REG_SZ, "1.0")
                    winreg.SetValueEx(attr_key, "Name", 0, winreg.REG_SZ, config["displayName"])
                    winreg.SetValueEx(attr_key, "VoiceType", 0, winreg.REG_SZ, "AACSpeakHelper")
                    winreg.SetValueEx(attr_key, "Description", 0, winreg.REG_SZ, config["description"])
            
            print(f"✅ SAPI voice registered: {config['displayName']}")
            return True

        except Exception as e:
            print(f"❌ Failed to register SAPI voice: {e}")
            return False

    def unregister_sapi_voice(self, voice_name):
        """Unregister a voice from Windows SAPI"""
        try:
            registry_path = f"{SAPI_REGISTRY_PATH}\\{voice_name}"
            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, registry_path + "\\Attributes")
            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, registry_path)
            print(f"✅ SAPI voice unregistered: {voice_name}")
            return True
        except FileNotFoundError:
            print(f"⚠️ Voice not found in registry: {voice_name}")
            return True  # Consider it unregistered if not found
        except Exception as e:
            print(f"❌ Failed to unregister SAPI voice: {e}")
            return False

    def list_installed_voices(self):
        """List all installed SAPI voices"""
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, SAPI_REGISTRY_PATH) as voices_key:
                voices = []
                i = 0
                while True:
                    try:
                        voice_name = winreg.EnumKey(voices_key, i)

                        # Check if it's one of our voices (has ConfigPath)
                        try:
                            with winreg.OpenKey(voices_key, voice_name) as voice_key:
                                config_path = winreg.QueryValueEx(voice_key, "ConfigPath")[0]
                                display_name = winreg.QueryValueEx(voice_key, "")[0]
                                voices.append({
                                    "name": voice_name,
                                    "display_name": display_name,
                                    "config_path": config_path,
                                    "is_ours": True
                                })
                        except FileNotFoundError:
                            # Not one of our voices
                            try:
                                with winreg.OpenKey(voices_key, voice_name) as voice_key:
                                    display_name = winreg.QueryValueEx(voice_key, "")[0]
                                    voices.append({
                                        "name": voice_name,
                                        "display_name": display_name,
                                        "config_path": None,
                                        "is_ours": False
                                    })
                            except:
                                pass

                        i += 1
                    except OSError:
                        break

                return voices

        except Exception as e:
            print(f"❌ Failed to list voices: {e}")
            return []

    def install_voice(self, voice_name, engine_key, voice_id, **kwargs):
        """Install a complete SAPI voice"""
        print(f"\n🔧 Installing voice: {voice_name}")

        # Step 0: Check for administrator privileges
        if not warn_admin_required():
            return False

        # Step 1: Register COM wrapper if needed
        if not self.register_com_wrapper():
            return False

        # Step 2: Create/update AACSpeakHelper config
        azure_key = kwargs.get('azure_key')
        azure_region = kwargs.get('azure_region', 'uksouth')
        if not self.create_aacspeakhelper_config(engine_key, azure_key, azure_region):
            print("⚠️  Warning: Failed to create AACSpeakHelper config, but continuing...")

        # Step 3: Create voice configuration
        config_path = self.create_voice_config(voice_name, engine_key, voice_id, **kwargs)
        if not config_path:
            return False

        # Step 4: Load the configuration
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except Exception as e:
            print(f"❌ Failed to load voice config: {e}")
            return False

        # Step 5: Register with SAPI
        if not self.register_sapi_voice(voice_name, config_path, config):
            return False

        print(f"✅ Voice installation complete: {voice_name}")
        return True

    def uninstall_voice(self, voice_name, config_path=None):
        """Uninstall a SAPI voice"""
        print(f"\n🗑️ Uninstalling voice: {voice_name}")

        # Step 1: Unregister from SAPI
        if not self.unregister_sapi_voice(voice_name):
            return False

        # Step 2: Remove configuration file
        if config_path:
            # Use the provided config path
            config_file_path = Path(config_path)
        else:
            # Fallback to constructing the path
            config_file_path = VOICE_CONFIGS_DIR / f"{voice_name}.json"

        try:
            if config_file_path.exists():
                config_file_path.unlink()
                print(f"✅ Voice config removed: {config_file_path}")
            else:
                print(f"⚠️ Config file not found: {config_file_path}")
        except Exception as e:
            print(f"⚠️ Failed to remove voice config: {e}")

        print(f"✅ Voice uninstallation complete: {voice_name}")
        return True

    def interactive_main_menu(self):
        """Main interactive menu"""
        print("\n🎤 Welcome to the SAPI Voice Installer!")
        print("This tool helps you manage TTS voices for Windows SAPI.")
        print("💡 Tip: Press Ctrl+C at any time to exit")

        # Check for administrator privileges upfront
        if not is_admin():
            print("\n⚠️  NOTICE: Running without administrator privileges")
            print("   Voice installation and COM registration require admin rights.")

        while True:
            print("\n" + "="*60)
            print("Main Menu:")
            print("="*60)
            print("1. Install New Voice")
            print("2. Configure TTS Engine Credentials")
            print("3. List Installed Voices")
            print("4. Uninstall Voice")
            print("5. Register COM Server")
            print("6. Unregister COM Server")
            print("7. Exit")

            # Get user selection
            try:
                user_input = input(f"\nSelect option (1-7) or 'q' to quit: ").strip().lower()

                # Handle quit commands
                if user_input in ['q', 'quit', 'exit']:
                    print("Goodbye! 👋")
                    return

                # Handle numeric selection
                choice = int(user_input)
                if choice == 1:
                    self.interactive_voice_selection()
                elif choice == 2:
                    self.interactive_credential_configuration()
                elif choice == 3:
                    print("\n📋 Installed SAPI Voices")
                    print("="*60)
                    voices = self.list_installed_voices()
                    if not voices:
                        print("No SAPI voices are currently installed.")
                    else:
                        print(f"\nFound {len(voices)} installed SAPI voices:")
                        print("-" * 80)
                        for i, voice in enumerate(voices):
                            voice_name = voice.get('name', 'Unknown')
                            display_name = voice.get('display_name', 'Unknown')
                            is_ours = voice.get('is_ours', False)
                            config_path = voice.get('config_path', 'N/A')

                            status = "🎤 AACSpeakHelper" if is_ours else "🔊 System"
                            print(f"{i+1:2d}. {display_name:<40} | {status:<15} | {voice_name}")
                            if is_ours and config_path:
                                print(f"    Config: {config_path}")
                    input("\nPress Enter to continue...")
                elif choice == 4:
                    self.interactive_uninstall()
                elif choice == 5:
                    if self.register_com_wrapper():
                        print("✅ COM server registered successfully!")
                    input("\nPress Enter to continue...")
                elif choice == 6:
                    if self.unregister_com_wrapper():
                        print("✅ COM server unregistered successfully!")
                    input("\nPress Enter to continue...")
                elif choice == 7:
                    print("Goodbye! 👋")
                    return
                else:
                    print("Invalid selection. Please try again.")
            except ValueError:
                print("Please enter a number or 'q' to quit.")
            except KeyboardInterrupt:
                print("\n\nExiting... Goodbye! 👋")
                return

    def interactive_credential_configuration(self):
        """Interactive credential configuration menu"""
        print("\n🔧 TTS Engine Credential Configuration")
        print("Configure API keys and credentials for TTS engines.")

        # Load current configuration
        config, config_path = load_config()

        while True:
            print("\n" + "="*60)
            print("Available TTS Engines:")
            print("="*60)

            # Show engines with credential status
            engines = list(TTS_ENGINES.keys())

            for i, engine_key in enumerate(engines):
                engine_config = TTS_ENGINES[engine_key]
                engine_name = engine_config["name"]
                credential_status = get_credential_status(config, engine_key)

                print(f"{i+1}. {engine_name:<20} {credential_status}")

            print(f"{len(engines)+1}. Save & Exit")

            # Get user selection
            try:
                user_input = input(f"\nSelect engine to configure (1-{len(engines)+1}) or 'q' to quit: ").strip().lower()

                # Handle quit commands
                if user_input in ['q', 'quit', 'exit']:
                    return

                # Handle numeric selection
                choice = int(user_input) - 1
                if choice == len(engines):  # Save & Exit
                    # Save configuration
                    try:
                        with open(config_path, 'w', encoding='utf-8') as f:
                            config.write(f)
                        print(f"✅ Configuration saved to {config_path}")
                    except Exception as e:
                        print(f"❌ Failed to save configuration: {e}")
                    return
                elif 0 <= choice < len(engines):
                    selected_engine = engines[choice]

                    # Configure credentials for selected engine
                    config = configure_engine_credentials(config, selected_engine)

                    # Offer to test credentials
                    engine_config = TTS_ENGINES[selected_engine]
                    if engine_config["credential_fields"]:
                        test_choice = input("\nTest credentials now? (y/N): ").strip().lower()
                        if test_choice == 'y':
                            test_engine_credentials(config, selected_engine)

                    input("\nPress Enter to continue...")
                else:
                    print("Invalid selection. Please try again.")
            except ValueError:
                print("Please enter a number or 'q' to quit.")
            except KeyboardInterrupt:
                print("\n\nReturning to main menu...")
                return

    def interactive_voice_selection(self):
        """Interactive voice selection and installation"""
        print("\n🔍 Voice Installation")
        print("Discover and install TTS voices from available engines.")

        # Load configuration to check for credentials
        config, config_path = load_config()

        while True:
            print("\n" + "="*60)
            print("Available TTS Engines:")
            print("="*60)

            # Show available engines with credential status
            engines = list(TTS_ENGINES.keys())
            available_engines = []

            for i, engine_key in enumerate(engines):
                engine_config = TTS_ENGINES[engine_key]
                engine_name = engine_config["name"]
                credential_status = get_credential_status(config, engine_key)

                print(f"{i+1}. {engine_name:<20} {credential_status}")
                available_engines.append(engine_key)

            print(f"{len(engines)+1}. Configure Credentials")
            print(f"{len(engines)+2}. Exit")

            # Get user selection
            try:
                user_input = input(f"\nSelect TTS engine (1-{len(engines)+2}) or 'q' to quit: ").strip().lower()

                # Handle quit commands
                if user_input in ['q', 'quit', 'exit']:
                    return

                # Handle numeric selection
                choice = int(user_input) - 1
                if choice == len(engines):  # Configure Credentials option
                    self.interactive_credential_configuration()
                elif choice == len(engines) + 1:  # Exit option
                    return
                elif 0 <= choice < len(engines):
                    selected_engine = available_engines[choice]

                    # Check if credentials are configured before proceeding
                    engine_config = TTS_ENGINES[selected_engine]
                    if engine_config["credential_fields"]:
                        credential_status = get_credential_status(config, selected_engine)
                        if "❌ Missing" in credential_status:
                            print(f"\n⚠️  {engine_config['name']} requires credentials to be configured first.")
                            configure_choice = input("Configure credentials now? (y/N): ").strip().lower()
                            if configure_choice == 'y':
                                config = configure_engine_credentials(config, selected_engine)
                                # Save the updated config
                                try:
                                    config_path, _ = load_config()
                                    with open(config_path, 'w', encoding='utf-8') as f:
                                        config.write(f)
                                    print("✅ Configuration saved")
                                except Exception as e:
                                    print(f"⚠️  Warning: Failed to save configuration: {e}")
                            else:
                                print("Cannot proceed without credentials.")
                                input("Press Enter to continue...")
                                continue

                    self.browse_engine_voices(selected_engine, config)
                else:
                    print("Invalid selection. Please try again.")
            except ValueError:
                print("Please enter a number or 'q' to quit.")
            except KeyboardInterrupt:
                print("\n\nExiting... Goodbye! 👋")
                return

    def browse_engine_voices(self, engine_key, config):
        """Browse and select voices from a specific engine"""
        engine_config = TTS_ENGINES[engine_key]
        engine_name = engine_config["name"]

        print(f"\n🔍 Discovering voices from {engine_name}...")

        # Try to get voices from the actual engine
        voices = get_voices_from_engine(engine_key, config)

        if not voices:
            # Fallback to hardcoded voice list
            print("Using fallback voice list...")
            voice_list = engine_config["voice_list"]
            if not voice_list:
                print(f"No voices available for {engine_name}")
                input("Press Enter to continue...")
                return

            # Convert dict to list format
            voices = []
            for voice_name, voice_id in voice_list.items():
                voices.append({
                    'id': voice_id,
                    'name': voice_name,
                    'language': voice_name,
                    'gender': 'Unknown'
                })

        if not voices:
            print(f"No voices found for {engine_name}")
            input("Press Enter to continue...")
            return

        # Allow searching by language
        try:
            search_input = input(
                "Search for a voice by language/name (Enter=all, 'q'=back): "
            ).strip()

            # Handle quit/back commands
            if search_input.lower() in ['q', 'quit', 'back']:
                return

            search_term = search_input.lower()
        except KeyboardInterrupt:
            print("\n\nReturning to engine selection...")
            return

        matching_voices = []
        for voice in voices:
            voice_name = voice.get('name', voice.get('id', ''))
            voice_id = voice.get('id', voice.get('voice_id', ''))
            voice_language = voice.get('language', '')

            # Create searchable text that includes extracted language info from voice ID
            searchable_text = f"{voice_name} {voice_id} {voice_language}".lower()

            # Extract language info from voice ID (e.g., "en-GB-LibbyNeural" -> "english british")
            if voice_id:
                # Comprehensive language mappings for better search
                lang_mappings = {
                    # Major languages
                    'en-': 'english', 'fr-': 'french', 'de-': 'german', 'es-': 'spanish',
                    'it-': 'italian', 'pt-': 'portuguese', 'ru-': 'russian', 'ja-': 'japanese',
                    'ko-': 'korean', 'zh-': 'chinese', 'ar-': 'arabic', 'hi-': 'hindi',
                    'nl-': 'dutch', 'sv-': 'swedish', 'da-': 'danish', 'no-': 'norwegian',
                    'fi-': 'finnish', 'pl-': 'polish', 'cs-': 'czech', 'sk-': 'slovak',
                    'hu-': 'hungarian', 'ro-': 'romanian', 'bg-': 'bulgarian', 'hr-': 'croatian',
                    'sr-': 'serbian', 'sl-': 'slovenian', 'et-': 'estonian', 'lv-': 'latvian',
                    'lt-': 'lithuanian', 'el-': 'greek', 'tr-': 'turkish', 'he-': 'hebrew',
                    'th-': 'thai', 'vi-': 'vietnamese', 'id-': 'indonesian', 'ms-': 'malay',
                    'tl-': 'filipino', 'ur-': 'urdu', 'bn-': 'bengali', 'ta-': 'tamil',
                    'te-': 'telugu', 'ml-': 'malayalam', 'kn-': 'kannada', 'gu-': 'gujarati',
                    'pa-': 'punjabi', 'mr-': 'marathi', 'ne-': 'nepali', 'si-': 'sinhala',
                    'my-': 'burmese', 'km-': 'khmer', 'lo-': 'lao', 'ka-': 'georgian',
                    'hy-': 'armenian', 'az-': 'azerbaijani', 'kk-': 'kazakh', 'ky-': 'kyrgyz',
                    'uz-': 'uzbek', 'tg-': 'tajik', 'mn-': 'mongolian', 'ps-': 'pashto',
                    'fa-': 'persian', 'sw-': 'swahili', 'am-': 'amharic', 'zu-': 'zulu',
                    'af-': 'afrikaans', 'is-': 'icelandic', 'mt-': 'maltese', 'cy-': 'welsh',
                    'ga-': 'irish', 'eu-': 'basque', 'ca-': 'catalan', 'gl-': 'galician',

                    # Country/region mappings
                    '-us': 'american', '-gb': 'british', '-au': 'australian', '-ca': 'canadian',
                    '-fr': 'france', '-de': 'germany', '-es': 'spain', '-it': 'italy',
                    '-in': 'indian', '-pk': 'pakistani', '-bd': 'bangladeshi', '-af': 'afghanistan',
                    '-ir': 'iranian', '-iq': 'iraqi', '-sa': 'saudi', '-ae': 'emirates',
                    '-eg': 'egyptian', '-ma': 'moroccan', '-dz': 'algerian', '-tn': 'tunisian',
                    '-ly': 'libyan', '-sd': 'sudanese', '-so': 'somali', '-et': 'ethiopian',
                    '-ke': 'kenyan', '-tz': 'tanzanian', '-za': 'south african', '-ng': 'nigerian',
                    '-mx': 'mexican', '-ar': 'argentinian', '-br': 'brazilian', '-cl': 'chilean',
                    '-co': 'colombian', '-pe': 'peruvian', '-ve': 'venezuelan', '-ec': 'ecuadorian'
                }

                for code, lang in lang_mappings.items():
                    if code in voice_id.lower():
                        searchable_text += f" {lang}"

            if not search_term or search_term in searchable_text:
                matching_voices.append(voice)

        if not matching_voices:
            print("No matching voices found.")
            input("Press Enter to continue...")
            return

        # Display matching voices
        print(f"\n🎵 Found {len(matching_voices)} matching voices:")
        print("="*80)
        for i, voice in enumerate(matching_voices):
            voice_name = voice.get('name', voice.get('id', ''))
            voice_id = voice.get('id', voice.get('voice_id', ''))
            voice_language = voice.get('language', 'Unknown')
            voice_gender = voice.get('gender', 'Unknown')
            print(f"{i+1:2d}. {voice_name:<30} | {voice_language:<15} | {voice_gender:<8} | {voice_id}")

        print(f"{len(matching_voices)+1:2d}. Back to engine selection")

        # Get user selection
        while True:
            try:
                user_input = input(f"\nSelect voice to install (1-{len(matching_voices)+1}) or 'q' to go back: ").strip().lower()

                # Handle quit/back commands
                if user_input in ['q', 'quit', 'back']:
                    return

                # Handle numeric selection
                choice = int(user_input) - 1
                if choice == len(matching_voices):  # Back option
                    return
                elif 0 <= choice < len(matching_voices):
                    selected_voice = matching_voices[choice]
                    self.install_selected_voice(engine_key, selected_voice)
                    return
                else:
                    print("Invalid selection. Please try again.")
            except ValueError:
                print("Please enter a number or 'q' to go back.")
            except KeyboardInterrupt:
                print("\n\nReturning to engine selection...")
                return

    def install_selected_voice(self, engine_key, voice_data):
        """Install a selected voice"""
        voice_name = voice_data.get('name', voice_data.get('id', ''))
        voice_id = voice_data.get('id', voice_data.get('voice_id', ''))

        # Create a safe registry name
        safe_name = "".join(c for c in voice_name if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_name = safe_name.replace(' ', '-')

        print(f"\n🔧 Installing voice: {voice_name}")
        print(f"   Engine: {TTS_ENGINES[engine_key]['name']}")
        print(f"   Voice ID: {voice_id}")
        print(f"   Registry Name: {safe_name}")

        # Confirm installation
        try:
            confirm = input("\nProceed with installation? (y/N/q=back): ").strip().lower()
            if confirm in ['q', 'quit', 'back']:
                return
            elif confirm != 'y':
                print("Installation cancelled.")
                return
        except KeyboardInterrupt:
            print("\n\nInstallation cancelled.")
            return

        # Prepare installation parameters
        kwargs = {
            'display_name': voice_name,
            'language': voice_data.get('language', 'English'),
            'locale': voice_data.get('locale', 'en-GB'),
            'gender': voice_data.get('gender', 'Unknown'),
            'description': f"{TTS_ENGINES[engine_key]['name']} voice - {voice_name}"
        }

        # Install the voice
        success = self.install_voice(safe_name, engine_key, voice_id, **kwargs)

        if success:
            print(f"\n✅ Voice '{voice_name}' installed successfully!")
            print("   You can now use this voice in any SAPI-compatible application.")
        else:
            print(f"\n❌ Failed to install voice '{voice_name}'")

        input("\nPress Enter to continue...")

    def interactive_uninstall(self):
        """Interactive voice uninstallation"""
        print("\n🗑️  Voice Uninstallation")

        # Get list of installed voices
        voices = self.list_installed_voices()
        if not voices:
            print("No SAPI voices are currently installed.")
            return

        print(f"\nFound {len(voices)} installed SAPI voices:")
        print("="*60)

        for i, voice in enumerate(voices):
            print(f"{i+1:2d}. {voice}")

        print(f"{len(voices)+1:2d}. Back to main menu")

        # Get user selection
        while True:
            try:
                user_input = input(f"\nSelect voice to uninstall (1-{len(voices)+1}) or 'q' to go back: ").strip().lower()

                # Handle quit/back commands
                if user_input in ['q', 'quit', 'back']:
                    return

                # Handle numeric selection
                choice = int(user_input) - 1
                if choice == len(voices):  # Back option
                    return
                elif 0 <= choice < len(voices):
                    selected_voice = voices[choice]
                    voice_name = selected_voice.get('name', 'Unknown')
                    display_name = selected_voice.get('display_name', voice_name)
                    config_path = selected_voice.get('config_path')

                    # Confirm uninstallation
                    confirm = input(f"\nAre you sure you want to uninstall '{display_name}'? (y/N): ").strip().lower()
                    if confirm == 'y':
                        if self.uninstall_voice(voice_name, config_path):
                            print(f"✅ Voice '{display_name}' uninstalled successfully!")
                        else:
                            print(f"❌ Failed to uninstall voice '{display_name}'")
                    else:
                        print("Uninstallation cancelled.")
                    return
                else:
                    print("Invalid selection. Please try again.")
            except ValueError:
                print("Please enter a number or 'q' to go back.")
            except KeyboardInterrupt:
                print("\n\nReturning to main menu...")
                return


# Predefined voice configurations
PREDEFINED_VOICES = {
    "English-SherpaOnnx-Jenny": {
        "engine": "sherpa",
        "voice_id": "en_GB-jenny_dioco-medium",
        "display_name": "English (SherpaOnnx Jenny)",
        "description": "SherpaOnnx neural voice - Jenny (British English)",
        "language": "English",
        "locale": "en-GB",
        "gender": "Female",
        "age": "Adult"
    },
    "British-English-Azure-Libby": {
        "engine": "azure",
        "voice_id": "en-GB-LibbyNeural",
        "display_name": "British English (Azure Libby)",
        "description": "Azure TTS neural voice - Libby (British English)",
        "language": "English",
        "locale": "en-GB",
        "gender": "Female",
        "age": "Adult"
    },
    "American-English-Azure-Jenny": {
        "engine": "azure",
        "voice_id": "en-US-JennyNeural",
        "display_name": "American English (Azure Jenny)",
        "description": "Azure TTS neural voice - Jenny (American English)",
        "language": "English",
        "locale": "en-US",
        "gender": "Female",
        "age": "Adult"
    },
    "English-Google-Basic": {
        "engine": "google",
        "voice_id": "en-US-Wavenet-C",
        "display_name": "English (Google Wavenet)",
        "description": "Google TTS Wavenet voice (American English)",
        "language": "English",
        "locale": "en-US",
        "gender": "Female",
        "age": "Adult"
    }
}


def main():
    """Main CLI interface"""
    parser = argparse.ArgumentParser(
        description="OpenSpeechSAPI - Universal SAPI Voice Installer",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Interactive mode (discover and install voices)
  python sapi_voice_installer.py

  # Install a predefined voice
  python sapi_voice_installer.py install English-SherpaOnnx-Jenny

  # Install a custom voice
  python sapi_voice_installer.py install-custom MyVoice sherpa en_GB-jenny_dioco-medium

  # List all voices
  python sapi_voice_installer.py list

  # Uninstall a voice
  python sapi_voice_installer.py uninstall English-SherpaOnnx-Jenny

  # Register COM wrapper only
  python sapi_voice_installer.py register-com

Interactive Mode:
  When run without arguments, the installer enters interactive mode where you can:
  - Browse available TTS engines (Azure, Sherpa-ONNX, Google, etc.)
  - Search and discover voices from each engine
  - Install voices with a user-friendly interface
  - See credential status for each engine
        """
    )

    subparsers = parser.add_subparsers(dest='command', help='Available commands')

    # Install predefined voice
    install_parser = subparsers.add_parser('install', help='Install a predefined voice')
    install_parser.add_argument('voice_name', choices=list(PREDEFINED_VOICES.keys()),
                               help='Name of predefined voice to install')

    # Install custom voice
    custom_parser = subparsers.add_parser('install-custom', help='Install a custom voice')
    custom_parser.add_argument('voice_name', help='Name for the voice')
    custom_parser.add_argument('engine', choices=list(TTS_ENGINES.keys()), help='TTS engine')
    custom_parser.add_argument('voice_id', help='Voice ID for the engine')
    custom_parser.add_argument('--display-name', help='Display name for the voice')
    custom_parser.add_argument('--locale', default='en-GB', help='Voice locale (default: en-GB)')
    custom_parser.add_argument('--gender', default='Female', help='Voice gender (default: Female)')

    # List voices
    list_parser = subparsers.add_parser('list', help='List all installed SAPI voices')

    # Uninstall voice
    uninstall_parser = subparsers.add_parser('uninstall', help='Uninstall a voice')
    uninstall_parser.add_argument('voice_name', help='Name of voice to uninstall')

    # Register COM wrapper
    com_parser = subparsers.add_parser('register-com', help='Register the COM wrapper')

    # Unregister COM wrapper
    uncom_parser = subparsers.add_parser('unregister-com', help='Unregister the COM wrapper')

    args = parser.parse_args()

    if not args.command:
        # Interactive mode - no command provided
        installer = SAPIVoiceInstaller()
        installer.interactive_main_menu()
        return

    # Check for admin privileges
    try:
        import ctypes
        if not ctypes.windll.shell32.IsUserAnAdmin():
            print("❌ This tool requires administrator privileges")
            print("   Please run as administrator")
            sys.exit(1)
    except:
        print("⚠️ Could not check admin privileges - continuing anyway")

    installer = SAPIVoiceInstaller()

    if args.command == 'install':
        voice_config = PREDEFINED_VOICES[args.voice_name]
        success = installer.install_voice(
            args.voice_name,
            voice_config['engine'],
            voice_config['voice_id'],
            **{k: v for k, v in voice_config.items() if k not in ['engine', 'voice_id']}
        )
        sys.exit(0 if success else 1)

    elif args.command == 'install-custom':
        kwargs = {}
        if args.display_name:
            kwargs['display_name'] = args.display_name
        kwargs['locale'] = args.locale
        kwargs['gender'] = args.gender

        success = installer.install_voice(
            args.voice_name,
            args.engine,
            args.voice_id,
            **kwargs
        )
        sys.exit(0 if success else 1)

    elif args.command == 'list':
        voices = installer.list_installed_voices()
        if not voices:
            print("No SAPI voices found")
            return

        print("\n📋 Installed SAPI Voices:")
        print("=" * 60)
        for voice in voices:
            status = "🔧 AACSpeakHelper" if voice['is_ours'] else "🖥️ System"
            print(f"{status} {voice['display_name']}")
            print(f"   Registry: {voice['name']}")
            if voice['config_path']:
                print(f"   Config: {voice['config_path']}")
            print()

    elif args.command == 'uninstall':
        success = installer.uninstall_voice(args.voice_name)
        sys.exit(0 if success else 1)

    elif args.command == 'register-com':
        success = installer.register_com_wrapper()
        sys.exit(0 if success else 1)

    elif args.command == 'unregister-com':
        success = installer.unregister_com_wrapper()
        sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
