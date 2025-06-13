#!/usr/bin/env python3
"""
SAPI Voice Manager - CLI tool for managing configuration-based SAPI voices
Integrates with AACSpeakHelper pipe service and Windows SAPI registry

This tool creates SAPI voices that are configuration files, where each voice
corresponds to specific TTS engine settings sent to AACSpeakHelper pipe service.

Author: OpenAssistive
License: MIT
"""

import os
import sys
import json
import argparse
import configparser
import subprocess
import winreg
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# TTS Engine configurations matching AACSpeakHelper
TTS_ENGINES = {
    "azure": {
        "name": "Azure TTS",
        "config_section": "azureTTS", 
        "credential_fields": ["key", "location"],
        "required_fields": ["voice"],
        "description": "Microsoft Azure Cognitive Services TTS"
    },
    "google": {
        "name": "Google TTS",
        "config_section": "googleTTS",
        "credential_fields": ["creds"],
        "required_fields": ["voice", "lang"],
        "description": "Google Cloud Text-to-Speech"
    },
    "sherpa": {
        "name": "Sherpa-ONNX",
        "config_section": "SherpaOnnxTTS", 
        "credential_fields": [],
        "required_fields": ["voice_id"],
        "description": "Local Sherpa-ONNX TTS engine"
    },
    "google_trans": {
        "name": "Google Translate TTS",
        "config_section": "googleTransTTS",
        "credential_fields": [],
        "required_fields": ["voice_id"],
        "description": "Google Translate TTS (free)"
    }
}

# Common voice locales for SAPI registration
VOICE_LOCALES = {
    "en-US": {"lcid": "409", "name": "English (United States)"},
    "en-GB": {"lcid": "809", "name": "English (United Kingdom)"},
    "en-AU": {"lcid": "c09", "name": "English (Australia)"},
    "en-CA": {"lcid": "1009", "name": "English (Canada)"},
    "fr-FR": {"lcid": "40c", "name": "French (France)"},
    "de-DE": {"lcid": "407", "name": "German (Germany)"},
    "es-ES": {"lcid": "c0a", "name": "Spanish (Spain)"},
    "it-IT": {"lcid": "410", "name": "Italian (Italy)"},
    "pt-BR": {"lcid": "416", "name": "Portuguese (Brazil)"},
    "ja-JP": {"lcid": "411", "name": "Japanese (Japan)"},
    "ko-KR": {"lcid": "412", "name": "Korean (Korea)"},
    "zh-CN": {"lcid": "804", "name": "Chinese (Simplified)"},
    "zh-TW": {"lcid": "404", "name": "Chinese (Traditional)"}
}

class SapiVoiceManager:
    """Main class for managing SAPI voices with pipe service integration"""
    
    def __init__(self):
        self.voice_configs_dir = Path("voice_configs")
        self.voice_configs_dir.mkdir(exist_ok=True)
        self.installer_path = self._find_installer()
        
    def _find_installer(self) -> Optional[Path]:
        """Find the SherpaOnnx SAPI installer executable"""
        possible_paths = [
            Path("Installer/bin/Release/net6.0/SherpaOnnxSAPIInstaller.exe"),
            Path("Installer/bin/Debug/net6.0/SherpaOnnxSAPIInstaller.exe"),
            Path("dist/SherpaOnnxSAPIInstaller.exe"),
            Path("SherpaOnnxSAPIInstaller.exe")
        ]
        
        for path in possible_paths:
            if path.exists():
                return path
        return None
    
    def print_banner(self):
        """Print the application banner"""
        print("\n" + "="*60)
        print("    SAPI Voice Manager - Configuration-Based Voices")
        print("    Integrates with AACSpeakHelper Pipe Service")
        print("="*60)
    
    def print_main_menu(self):
        """Print the main menu"""
        print("\nMain Menu:")
        print("1. Create new voice configuration")
        print("2. List voice configurations") 
        print("3. Install voice to SAPI")
        print("4. Remove voice from SAPI")
        print("5. Test pipe service connection")
        print("6. View voice configuration")
        print("7. Edit voice configuration")
        print("8. Export voice configuration")
        print("9. Import voice configuration")
        print("10. Exit")
        print("-" * 40)
    
    def create_voice_config(self):
        """Interactive voice configuration creation"""
        print("\n=== Create New Voice Configuration ===")
        
        # Get basic voice information
        voice_name = input("Voice name (e.g., British-English-Azure-Libby): ").strip()
        if not voice_name:
            print("Voice name is required!")
            return
            
        display_name = input("Display name (e.g., British English (Azure Libby)): ").strip()
        if not display_name:
            display_name = voice_name
            
        description = input("Description: ").strip()
        
        # Select locale
        print("\nAvailable locales:")
        locales = list(VOICE_LOCALES.keys())
        for i, locale in enumerate(locales):
            print(f"{i+1}. {locale} - {VOICE_LOCALES[locale]['name']}")
        
        while True:
            try:
                choice = int(input(f"Select locale (1-{len(locales)}): ")) - 1
                if 0 <= choice < len(locales):
                    selected_locale = locales[choice]
                    break
                else:
                    print("Invalid selection!")
            except ValueError:
                print("Please enter a number!")
        
        # Select gender
        print("\nGender:")
        print("1. Female")
        print("2. Male") 
        print("3. Neutral")
        
        gender_map = {"1": "Female", "2": "Male", "3": "Neutral"}
        while True:
            gender_choice = input("Select gender (1-3): ")
            if gender_choice in gender_map:
                selected_gender = gender_map[gender_choice]
                break
            else:
                print("Invalid selection!")
        
        # Select TTS engine
        print("\nAvailable TTS engines:")
        engines = list(TTS_ENGINES.keys())
        for i, engine in enumerate(engines):
            engine_info = TTS_ENGINES[engine]
            print(f"{i+1}. {engine_info['name']} - {engine_info['description']}")
        
        while True:
            try:
                choice = int(input(f"Select TTS engine (1-{len(engines)}): ")) - 1
                if 0 <= choice < len(engines):
                    selected_engine = engines[choice]
                    break
                else:
                    print("Invalid selection!")
            except ValueError:
                print("Please enter a number!")
        
        # Configure engine-specific settings
        tts_config = self._configure_engine_settings(selected_engine)
        if not tts_config:
            print("Engine configuration cancelled.")
            return
        
        # Create voice configuration
        voice_config = {
            "name": voice_name,
            "displayName": display_name,
            "description": description,
            "language": VOICE_LOCALES[selected_locale]["name"].split(" (")[0],
            "locale": selected_locale,
            "gender": selected_gender,
            "age": "Adult",
            "vendor": TTS_ENGINES[selected_engine]["name"],
            "ttsConfig": tts_config
        }
        
        # Save configuration
        config_file = self.voice_configs_dir / f"{voice_name}.json"
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(voice_config, f, indent=2, ensure_ascii=False)
            print(f"\n✅ Voice configuration saved: {config_file}")
            print(f"   Display Name: {display_name}")
            print(f"   Engine: {TTS_ENGINES[selected_engine]['name']}")
            print(f"   Locale: {selected_locale}")
        except Exception as e:
            print(f"❌ Error saving configuration: {e}")
    
    def _configure_engine_settings(self, engine_key: str) -> Optional[Dict]:
        """Configure engine-specific settings"""
        engine_info = TTS_ENGINES[engine_key]
        print(f"\n=== Configure {engine_info['name']} ===")
        
        # Base TTS configuration
        tts_config = {
            "engine": engine_key,
            "TTS": {
                "engine": engine_key,
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
        
        # Engine-specific configuration
        engine_section = {}
        
        # Get credentials
        for field in engine_info["credential_fields"]:
            value = input(f"Enter {field}: ").strip()
            if not value and field in ["key", "creds"]:
                print(f"Warning: {field} is typically required for {engine_info['name']}")
            engine_section[field] = value
        
        # Get required fields
        for field in engine_info["required_fields"]:
            if field == "voice" and engine_key == "azure":
                value = input("Azure voice name (e.g., en-US-JennyNeural): ").strip()
            elif field == "voice_id":
                if engine_key == "sherpa":
                    value = input("Sherpa voice ID (e.g., sherpa-piper-amy-normal): ").strip()
                else:
                    value = input(f"{field}: ").strip()
            else:
                value = input(f"{field}: ").strip()
            
            if not value:
                print(f"❌ {field} is required!")
                return None
            engine_section[field] = value
        
        # Add engine-specific configuration
        section_name = engine_info["config_section"]
        tts_config[section_name] = engine_section
        
        # Set voice_id in TTS section
        if "voice" in engine_section:
            tts_config["TTS"]["voice_id"] = engine_section["voice"]
        elif "voice_id" in engine_section:
            tts_config["TTS"]["voice_id"] = engine_section["voice_id"]
        
        return tts_config

    def list_voice_configs(self):
        """List all voice configurations"""
        print("\n=== Voice Configurations ===")

        config_files = list(self.voice_configs_dir.glob("*.json"))
        if not config_files:
            print("No voice configurations found.")
            return

        for i, config_file in enumerate(config_files, 1):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                print(f"\n{i}. {config.get('displayName', 'Unknown')}")
                print(f"   File: {config_file.name}")
                print(f"   Engine: {config.get('vendor', 'Unknown')}")
                print(f"   Locale: {config.get('locale', 'Unknown')}")
                print(f"   Gender: {config.get('gender', 'Unknown')}")

                # Show if installed in SAPI
                if self._is_voice_installed(config.get('name', '')):
                    print("   Status: ✅ Installed in SAPI")
                else:
                    print("   Status: ⚪ Not installed")

            except Exception as e:
                print(f"❌ Error reading {config_file.name}: {e}")

    def install_voice_to_sapi(self):
        """Install a voice configuration to SAPI"""
        print("\n=== Install Voice to SAPI ===")

        # List available configurations
        config_files = list(self.voice_configs_dir.glob("*.json"))
        if not config_files:
            print("No voice configurations found. Create one first.")
            return

        print("Available configurations:")
        for i, config_file in enumerate(config_files, 1):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                print(f"{i}. {config.get('displayName', config_file.stem)}")
            except:
                print(f"{i}. {config_file.stem} (error reading)")

        # Get user selection
        while True:
            try:
                choice = int(input(f"Select configuration to install (1-{len(config_files)}): ")) - 1
                if 0 <= choice < len(config_files):
                    selected_file = config_files[choice]
                    break
                else:
                    print("Invalid selection!")
            except ValueError:
                print("Please enter a number!")

        # Install using the installer
        if not self.installer_path:
            print("❌ Installer not found. Please build the project first.")
            print("   Run: dotnet build Installer/Installer.csproj -c Release")
            return

        config_name = selected_file.stem
        try:
            result = subprocess.run([
                str(self.installer_path),
                "install-pipe-voice",
                config_name
            ], capture_output=True, text=True, check=True)

            print(f"✅ Successfully installed voice: {config_name}")
            print("   The voice should now appear in Windows SAPI applications.")

        except subprocess.CalledProcessError as e:
            print(f"❌ Installation failed: {e}")
            if e.stdout:
                print(f"Output: {e.stdout}")
            if e.stderr:
                print(f"Error: {e.stderr}")
        except Exception as e:
            print(f"❌ Error running installer: {e}")

    def remove_voice_from_sapi(self):
        """Remove a voice from SAPI"""
        print("\n=== Remove Voice from SAPI ===")

        # List installed voices
        installed_voices = self._get_installed_voices()
        if not installed_voices:
            print("No pipe-based voices are currently installed.")
            return

        print("Installed pipe-based voices:")
        for i, voice_name in enumerate(installed_voices, 1):
            print(f"{i}. {voice_name}")

        # Get user selection
        while True:
            try:
                choice = int(input(f"Select voice to remove (1-{len(installed_voices)}): ")) - 1
                if 0 <= choice < len(installed_voices):
                    selected_voice = installed_voices[choice]
                    break
                else:
                    print("Invalid selection!")
            except ValueError:
                print("Please enter a number!")

        # Remove using the installer
        if not self.installer_path:
            print("❌ Installer not found. Please build the project first.")
            return

        try:
            result = subprocess.run([
                str(self.installer_path),
                "remove-pipe-voice",
                selected_voice
            ], capture_output=True, text=True, check=True)

            print(f"✅ Successfully removed voice: {selected_voice}")

        except subprocess.CalledProcessError as e:
            print(f"❌ Removal failed: {e}")
            if e.stdout:
                print(f"Output: {e.stdout}")
            if e.stderr:
                print(f"Error: {e.stderr}")
        except Exception as e:
            print(f"❌ Error running installer: {e}")

    def install_voice_by_name(self, voice_name: str) -> bool:
        """Install a specific voice by name (non-interactive)"""
        print(f"Installing voice: {voice_name}")

        # Find the configuration file
        config_file = self.voice_configs_dir / f"{voice_name}.json"
        if not config_file.exists():
            print(f"❌ Voice configuration not found: {config_file}")
            print("Available configurations:")
            self.list_voice_configs()
            return False

        # Check if installer exists
        if not self.installer_path:
            print("❌ Installer not found. Please build the project first.")
            print("   Run: dotnet build Installer/Installer.csproj -c Release")
            return False

        # Install using the installer
        try:
            result = subprocess.run([
                str(self.installer_path),
                "install-pipe-voice",
                voice_name
            ], capture_output=True, text=True, check=True)

            print(f"✅ Successfully installed voice: {voice_name}")
            print("   The voice should now appear in Windows SAPI applications.")
            return True

        except subprocess.CalledProcessError as e:
            print(f"❌ Installation failed: {e}")
            if e.stdout:
                print(f"Output: {e.stdout}")
            if e.stderr:
                print(f"Error: {e.stderr}")
            return False
        except Exception as e:
            print(f"❌ Error running installer: {e}")
            return False

    def remove_voice_by_name(self, voice_name: str) -> bool:
        """Remove a specific voice by name (non-interactive)"""
        print(f"Removing voice: {voice_name}")

        # Check if voice is installed
        if not self._is_voice_installed(voice_name):
            print(f"❌ Voice not found in SAPI registry: {voice_name}")
            print("Installed voices:")
            installed = self._get_installed_voices()
            if installed:
                for voice in installed:
                    print(f"  - {voice}")
            else:
                print("  No pipe-based voices installed.")
            return False

        # Check if installer exists
        if not self.installer_path:
            print("❌ Installer not found. Please build the project first.")
            return False

        # Remove using the installer
        try:
            result = subprocess.run([
                str(self.installer_path),
                "remove-pipe-voice",
                voice_name
            ], capture_output=True, text=True, check=True)

            print(f"✅ Successfully removed voice: {voice_name}")
            return True

        except subprocess.CalledProcessError as e:
            print(f"❌ Removal failed: {e}")
            if e.stdout:
                print(f"Output: {e.stdout}")
            if e.stderr:
                print(f"Error: {e.stderr}")
            return False
        except Exception as e:
            print(f"❌ Error running installer: {e}")
            return False

    def list_installed_voices(self):
        """List installed SAPI voices (non-interactive)"""
        print("Installed pipe-based voices:")
        installed_voices = self._get_installed_voices()
        if not installed_voices:
            print("  No pipe-based voices are currently installed.")
        else:
            for voice in installed_voices:
                print(f"  - {voice}")

    def remove_all_voices(self) -> bool:
        """Remove all installed pipe-based voices (non-interactive)"""
        print("Removing all installed pipe-based voices...")

        installed_voices = self._get_installed_voices()
        if not installed_voices:
            print("No pipe-based voices are currently installed.")
            return True

        success_count = 0
        for voice_name in installed_voices:
            print(f"Removing: {voice_name}")
            if self.remove_voice_by_name(voice_name):
                success_count += 1

        print(f"Removed {success_count}/{len(installed_voices)} voices.")
        return success_count == len(installed_voices)

    def view_voice_config_by_name(self, voice_name: str):
        """View a specific voice configuration by name (non-interactive)"""
        config_file = self.voice_configs_dir / f"{voice_name}.json"
        if not config_file.exists():
            print(f"❌ Voice configuration not found: {voice_name}")
            print("Available configurations:")
            config_files = list(self.voice_configs_dir.glob("*.json"))
            if config_files:
                for config_file in config_files:
                    print(f"  - {config_file.stem}")
            else:
                print("  No configurations found.")
            return

        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)

            print(f"=== {config.get('displayName', 'Unknown')} ===")
            print(json.dumps(config, indent=2, ensure_ascii=False))

        except Exception as e:
            print(f"❌ Error reading configuration: {e}")

    def test_pipe_service(self):
        """Test connection to AACSpeakHelper pipe service"""
        print("\n=== Test Pipe Service Connection ===")

        if not self.installer_path:
            print("❌ Installer not found. Please build the project first.")
            return

        try:
            result = subprocess.run([
                str(self.installer_path),
                "test-pipe-service"
            ], capture_output=True, text=True, check=True)

            print("Pipe service test results:")
            print(result.stdout)

        except subprocess.CalledProcessError as e:
            print("Pipe service test failed:")
            if e.stdout:
                print(e.stdout)
            if e.stderr:
                print(e.stderr)
        except Exception as e:
            print(f"❌ Error running test: {e}")

    def _is_voice_installed(self, voice_name: str) -> bool:
        """Check if a voice is installed in SAPI registry"""
        try:
            registry_path = f"SOFTWARE\\Microsoft\\SPEECH\\Voices\\Tokens\\{voice_name}"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_path):
                return True
        except:
            return False

    def _get_installed_voices(self) -> List[str]:
        """Get list of installed pipe-based voices"""
        installed_voices = []
        try:
            registry_path = "SOFTWARE\\Microsoft\\SPEECH\\Voices\\Tokens"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_path) as key:
                i = 0
                while True:
                    try:
                        voice_name = winreg.EnumKey(key, i)
                        # Check if it's a pipe-based voice (our naming convention)
                        if any(pattern in voice_name for pattern in ["British-English", "American-English", "Azure", "Sherpa"]):
                            installed_voices.append(voice_name)
                        i += 1
                    except WindowsError:
                        break
        except:
            pass
        return installed_voices

    def view_voice_config(self):
        """View a voice configuration"""
        print("\n=== View Voice Configuration ===")

        config_files = list(self.voice_configs_dir.glob("*.json"))
        if not config_files:
            print("No voice configurations found.")
            return

        print("Available configurations:")
        for i, config_file in enumerate(config_files, 1):
            print(f"{i}. {config_file.stem}")

        while True:
            try:
                choice = int(input(f"Select configuration to view (1-{len(config_files)}): ")) - 1
                if 0 <= choice < len(config_files):
                    selected_file = config_files[choice]
                    break
                else:
                    print("Invalid selection!")
            except ValueError:
                print("Please enter a number!")

        try:
            with open(selected_file, 'r', encoding='utf-8') as f:
                config = json.load(f)

            print(f"\n=== {config.get('displayName', 'Unknown')} ===")
            print(json.dumps(config, indent=2, ensure_ascii=False))

        except Exception as e:
            print(f"❌ Error reading configuration: {e}")

    def edit_voice_config(self):
        """Edit a voice configuration"""
        print("\n=== Edit Voice Configuration ===")
        print("Note: This opens the JSON file in the default text editor.")

        config_files = list(self.voice_configs_dir.glob("*.json"))
        if not config_files:
            print("No voice configurations found.")
            return

        print("Available configurations:")
        for i, config_file in enumerate(config_files, 1):
            print(f"{i}. {config_file.stem}")

        while True:
            try:
                choice = int(input(f"Select configuration to edit (1-{len(config_files)}): ")) - 1
                if 0 <= choice < len(config_files):
                    selected_file = config_files[choice]
                    break
                else:
                    print("Invalid selection!")
            except ValueError:
                print("Please enter a number!")

        try:
            # Open in default editor
            os.startfile(str(selected_file))
            print(f"✅ Opened {selected_file.name} in default editor.")
            print("   Save the file when you're done editing.")

        except Exception as e:
            print(f"❌ Error opening file: {e}")
            print(f"   You can manually edit: {selected_file}")

    def export_voice_config(self):
        """Export voice configuration to a different location"""
        print("\n=== Export Voice Configuration ===")

        config_files = list(self.voice_configs_dir.glob("*.json"))
        if not config_files:
            print("No voice configurations found.")
            return

        print("Available configurations:")
        for i, config_file in enumerate(config_files, 1):
            print(f"{i}. {config_file.stem}")

        while True:
            try:
                choice = int(input(f"Select configuration to export (1-{len(config_files)}): ")) - 1
                if 0 <= choice < len(config_files):
                    selected_file = config_files[choice]
                    break
                else:
                    print("Invalid selection!")
            except ValueError:
                print("Please enter a number!")

        export_path = input("Export to (path/filename.json): ").strip()
        if not export_path:
            print("Export cancelled.")
            return

        try:
            import shutil
            shutil.copy2(selected_file, export_path)
            print(f"✅ Configuration exported to: {export_path}")

        except Exception as e:
            print(f"❌ Export failed: {e}")

    def import_voice_config(self):
        """Import voice configuration from a file"""
        print("\n=== Import Voice Configuration ===")

        import_path = input("Import from (path/filename.json): ").strip()
        if not import_path or not os.path.exists(import_path):
            print("File not found or path not specified.")
            return

        try:
            # Validate the configuration
            with open(import_path, 'r', encoding='utf-8') as f:
                config = json.load(f)

            required_fields = ["name", "displayName", "locale", "gender", "ttsConfig"]
            missing_fields = [field for field in required_fields if field not in config]

            if missing_fields:
                print(f"❌ Invalid configuration. Missing fields: {', '.join(missing_fields)}")
                return

            # Copy to voice_configs directory
            import shutil
            dest_path = self.voice_configs_dir / f"{config['name']}.json"
            shutil.copy2(import_path, dest_path)

            print(f"✅ Configuration imported: {config['displayName']}")
            print(f"   Saved as: {dest_path}")

        except json.JSONDecodeError:
            print("❌ Invalid JSON file.")
        except Exception as e:
            print(f"❌ Import failed: {e}")

    def run(self):
        """Main application loop"""
        self.print_banner()

        while True:
            self.print_main_menu()
            choice = input("Enter your choice (1-10): ").strip()

            if choice == "1":
                self.create_voice_config()
            elif choice == "2":
                self.list_voice_configs()
            elif choice == "3":
                self.install_voice_to_sapi()
            elif choice == "4":
                self.remove_voice_from_sapi()
            elif choice == "5":
                self.test_pipe_service()
            elif choice == "6":
                self.view_voice_config()
            elif choice == "7":
                self.edit_voice_config()
            elif choice == "8":
                self.export_voice_config()
            elif choice == "9":
                self.import_voice_config()
            elif choice == "10":
                print("\nGoodbye!")
                break
            else:
                print("❌ Invalid choice. Please try again.")

            input("\nPress Enter to continue...")


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description="SAPI Voice Manager - Configuration-based voice management",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python SapiVoiceManager.py                    # Interactive mode
  python SapiVoiceManager.py --list             # List configurations
  python SapiVoiceManager.py --test-pipe        # Test pipe service
  python SapiVoiceManager.py --install voice1   # Install specific voice
        """
    )

    parser.add_argument("--list", action="store_true",
                       help="List all voice configurations")
    parser.add_argument("--list-installed", action="store_true",
                       help="List installed SAPI voices")
    parser.add_argument("--test-pipe", action="store_true",
                       help="Test AACSpeakHelper pipe service connection")
    parser.add_argument("--install", metavar="VOICE_NAME",
                       help="Install specific voice to SAPI")
    parser.add_argument("--remove", metavar="VOICE_NAME",
                       help="Remove specific voice from SAPI")
    parser.add_argument("--remove-all", action="store_true",
                       help="Remove all installed pipe-based voices")
    parser.add_argument("--create", action="store_true",
                       help="Create new voice configuration")
    parser.add_argument("--view", metavar="VOICE_NAME",
                       help="View specific voice configuration")

    args = parser.parse_args()

    manager = SapiVoiceManager()

    # Handle command line arguments
    if args.list:
        manager.list_voice_configs()
    elif args.list_installed:
        manager.list_installed_voices()
    elif args.test_pipe:
        manager.test_pipe_service()
    elif args.install:
        success = manager.install_voice_by_name(args.install)
        sys.exit(0 if success else 1)
    elif args.remove:
        success = manager.remove_voice_by_name(args.remove)
        sys.exit(0 if success else 1)
    elif args.remove_all:
        success = manager.remove_all_voices()
        sys.exit(0 if success else 1)
    elif args.view:
        manager.view_voice_config_by_name(args.view)
    elif args.create:
        manager.create_voice_config()
    else:
        # Interactive mode
        manager.run()


if __name__ == "__main__":
    main()
