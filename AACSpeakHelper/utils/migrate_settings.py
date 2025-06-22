#!/usr/bin/env python3
"""
AACSpeakHelper Settings Migration Script

This standalone script migrates a settings.cfg file to the correct location
for AACSpeakHelper. It will:
1. Look for settings.cfg next to this script
2. Determine the correct target location based on the system
3. Backup any existing settings.cfg
4. Copy the new settings.cfg to the target location

Usage: python migrate_settings.py
       or run the compiled executable
"""

import os
import sys
import shutil
import datetime
import argparse


def get_target_config_directory():
    """
    Determine the target configuration directory for AACSpeakHelper.
    
    Returns:
        str: Path to the configuration directory
    """
    if getattr(sys, "frozen", False):
        # Running as frozen executable - use AppData location
        config_dir = os.path.join(
            os.path.expanduser("~"),
            "AppData",
            "Roaming", 
            "Ace Centre",
            "AACSpeakHelper"
        )
    else:
        # Running as Python script - could be development or user running script
        # Default to AppData location for consistency
        config_dir = os.path.join(
            os.path.expanduser("~"),
            "AppData",
            "Roaming",
            "Ace Centre", 
            "AACSpeakHelper"
        )
    
    return config_dir


def get_script_directory():
    """
    Get the directory where this script is located.
    
    Returns:
        str: Path to the script directory
    """
    if getattr(sys, "frozen", False):
        # Running as frozen executable
        return os.path.dirname(sys.executable)
    else:
        # Running as Python script
        return os.path.dirname(os.path.abspath(__file__))


def create_backup_directory(script_dir):
    """
    Create a backup directory parallel to the migration script.
    
    Args:
        script_dir (str): Directory where the migration script is located
        
    Returns:
        str: Path to the backup directory
    """
    backup_dir = os.path.join(script_dir, "settings_backup")
    os.makedirs(backup_dir, exist_ok=True)
    return backup_dir


def backup_existing_settings(target_config_path, backup_dir):
    """
    Backup existing settings.cfg if it exists.
    
    Args:
        target_config_path (str): Path to the existing settings.cfg
        backup_dir (str): Directory to store the backup
        
    Returns:
        str or None: Path to the backup file if created, None otherwise
    """
    if not os.path.exists(target_config_path):
        return None
    
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f"settings_backup_{timestamp}.cfg"
    backup_path = os.path.join(backup_dir, backup_filename)
    
    try:
        shutil.copy2(target_config_path, backup_path)
        print(f"✅ Backed up existing settings to: {backup_path}")
        return backup_path
    except Exception as e:
        print(f"❌ Error backing up existing settings: {e}")
        return None


def copy_settings_file(source_path, target_path):
    """
    Copy the settings file from source to target location.
    
    Args:
        source_path (str): Path to the source settings.cfg
        target_path (str): Path to the target location
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Ensure target directory exists
        target_dir = os.path.dirname(target_path)
        os.makedirs(target_dir, exist_ok=True)
        
        # Copy the file
        shutil.copy2(source_path, target_path)
        print(f"✅ Successfully copied settings.cfg to: {target_path}")
        return True
    except Exception as e:
        print(f"❌ Error copying settings file: {e}")
        return False


def validate_settings_file(settings_path):
    """
    Basic validation of the settings file.
    
    Args:
        settings_path (str): Path to the settings file
        
    Returns:
        bool: True if file appears valid, False otherwise
    """
    try:
        import configparser
        config = configparser.ConfigParser()
        config.read(settings_path)
        
        # Check for essential sections
        required_sections = ['App', 'TTS']
        missing_sections = [section for section in required_sections if not config.has_section(section)]
        
        if missing_sections:
            print(f"⚠️  Warning: Settings file is missing sections: {missing_sections}")
            return False
        
        print("✅ Settings file appears to be valid")
        return True
    except Exception as e:
        print(f"⚠️  Warning: Could not validate settings file: {e}")
        return False


def main():
    """Main migration function."""
    parser = argparse.ArgumentParser(description="Migrate AACSpeakHelper settings.cfg file")
    parser.add_argument("--source", help="Path to source settings.cfg file (default: look next to script)")
    parser.add_argument("--target", help="Target directory (default: auto-detect)")
    parser.add_argument("--no-backup", action="store_true", help="Skip backing up existing settings")
    parser.add_argument("--dry-run", action="store_true", help="Show what would be done without actually doing it")
    
    args = parser.parse_args()
    
    print("🔧 AACSpeakHelper Settings Migration Script")
    print("=" * 50)
    
    # Determine script directory
    script_dir = get_script_directory()
    print(f"📁 Script directory: {script_dir}")
    
    # Determine source settings file
    if args.source:
        source_settings = args.source
    else:
        source_settings = os.path.join(script_dir, "settings.cfg")
    
    print(f"📄 Looking for source settings: {source_settings}")
    
    # Check if source file exists
    if not os.path.exists(source_settings):
        print(f"❌ Error: Source settings file not found: {source_settings}")
        print("💡 Make sure settings.cfg is in the same directory as this script")
        return 1
    
    # Validate source settings file
    if not validate_settings_file(source_settings):
        response = input("⚠️  Settings file validation failed. Continue anyway? (y/N): ")
        if response.lower() != 'y':
            print("❌ Migration cancelled")
            return 1
    
    # Determine target directory
    if args.target:
        target_config_dir = args.target
    else:
        target_config_dir = get_target_config_directory()
    
    target_settings_path = os.path.join(target_config_dir, "settings.cfg")
    print(f"🎯 Target location: {target_settings_path}")
    
    if args.dry_run:
        print("\n🔍 DRY RUN - No files will be modified")
        print(f"Would copy: {source_settings}")
        print(f"To: {target_settings_path}")
        if os.path.exists(target_settings_path) and not args.no_backup:
            backup_dir = create_backup_directory(script_dir)
            print(f"Would backup existing file to: {backup_dir}")
        return 0
    
    # Create backup if existing file exists and backup is not disabled
    if not args.no_backup and os.path.exists(target_settings_path):
        backup_dir = create_backup_directory(script_dir)
        backup_path = backup_existing_settings(target_settings_path, backup_dir)
        if backup_path is None:
            response = input("⚠️  Could not create backup. Continue anyway? (y/N): ")
            if response.lower() != 'y':
                print("❌ Migration cancelled")
                return 1
    
    # Copy the settings file
    success = copy_settings_file(source_settings, target_settings_path)
    
    if success:
        print("\n🎉 Migration completed successfully!")
        print(f"📁 Settings file location: {target_settings_path}")
        print("\n💡 You can now run AACSpeakHelper and it will use the new settings.")
        return 0
    else:
        print("\n❌ Migration failed!")
        return 1


if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n❌ Migration cancelled by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        sys.exit(1)
