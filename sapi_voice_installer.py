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
from pathlib import Path

# Import the TTS engine definitions from cli_config_creator
try:
    from AACSpeakHelper.cli_config_creator import TTS_ENGINES
except ImportError:
    # Fallback TTS engines if cli_config_creator is not available
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
        }
    }

# Constants
NATIVE_TTS_WRAPPER_CLSID = "{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
VOICE_CONFIGS_DIR = Path("C:/Program Files/OpenAssistive/OpenSpeech/voice_configs")
AACSPEAKHELPER_CONFIG_PATH = Path("AACSpeakHelper/settings.cfg")
SAPI_REGISTRY_PATH = r"SOFTWARE\Microsoft\SPEECH\Voices\Tokens"

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
    
    def register_com_wrapper(self):
        """Register the C++ COM wrapper DLL"""
        dll_path = Path("NativeTTSWrapper/x64/Release/NativeTTSWrapper.dll")
        
        if not dll_path.exists():
            print(f"❌ COM wrapper DLL not found: {dll_path}")
            print("   Please build the C++ project first")
            return False
        
        try:
            # Register the DLL
            result = subprocess.run([
                "regsvr32", "/s", str(dll_path.absolute())
            ], check=True, capture_output=True)
            
            print(f"✅ COM wrapper registered: {dll_path}")
            return True
            
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed to register COM wrapper: {e}")
            return False
    
    def unregister_com_wrapper(self):
        """Unregister the C++ COM wrapper DLL"""
        dll_path = Path("NativeTTSWrapper/x64/Release/NativeTTSWrapper.dll")

        if not dll_path.exists():
            print(f"⚠️ COM wrapper DLL not found: {dll_path}")
            return True  # Consider it unregistered if it doesn't exist

        try:
            # Unregister the DLL
            result = subprocess.run([
                "regsvr32", "/s", "/u", str(dll_path.absolute())
            ], check=True, capture_output=True)

            print(f"✅ COM wrapper unregistered: {dll_path}")
            return True

        except subprocess.CalledProcessError as e:
            print(f"❌ Failed to unregister COM wrapper: {e}")
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

    def uninstall_voice(self, voice_name):
        """Uninstall a SAPI voice"""
        print(f"\n🗑️ Uninstalling voice: {voice_name}")

        # Step 1: Unregister from SAPI
        if not self.unregister_sapi_voice(voice_name):
            return False

        # Step 2: Remove configuration file
        config_path = VOICE_CONFIGS_DIR / f"{voice_name}.json"
        try:
            if config_path.exists():
                config_path.unlink()
                print(f"✅ Voice config removed: {config_path}")
        except Exception as e:
            print(f"⚠️ Failed to remove voice config: {e}")

        print(f"✅ Voice uninstallation complete: {voice_name}")
        return True


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
        description="SAPI Voice Installer for AACSpeakHelper",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
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
        parser.print_help()
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
