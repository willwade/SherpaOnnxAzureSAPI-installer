"""
AACSpeakHelperServer.py - Simplified Windows Named Pipe Server for TTS

This is a drastically simplified version that merges all functionality into a single file.
It provides text-to-speech services through a Windows named pipe for SAPI integration.

Key simplifications:
- All utility functions merged into this file
- Removed encryption/decryption complexity
- Removed GUI configuration tools
- Simplified translation features
- Direct configuration from settings.cfg
- Minimal dependencies

The server:
1. Creates a named pipe (\\\\.\\pipe\\AACSpeakHelper) to receive requests
2. Processes incoming JSON messages containing text to speak
3. Converts text to speech using the configured TTS engine
4. Returns audio bytes back through the pipe for SAPI integration

Author: Ace Centre (Simplified)
"""

import logging
import os
import sys
import json
import time
import threading
import win32file
import win32pipe
import win32event
import win32api
import configparser
from pathlib import Path
from PySide6.QtWidgets import QApplication, QWidget, QSystemTrayIcon, QMenu
from PySide6.QtGui import QIcon, QAction
from PySide6.QtCore import QThread, Signal, Slot, QTimer

# TTS imports - only import what we actually use
from tts_wrapper import (
    MicrosoftTTS,
    SherpaOnnxTTS,
    GoogleTTS,
)

# Global variables
config = None
audio_files_path = None
config_path = None
ready = True

def setup_logging():
    """Setup logging to file"""
    if getattr(sys, "frozen", False):
        log_dir = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Ace Centre", "AACSpeakHelper")
    else:
        log_dir = os.path.dirname(os.path.abspath(__file__))
    
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    log_file = os.path.join(log_dir, "app.log")
    
    logging.basicConfig(
        filename=log_file,
        filemode="a",
        format="%(asctime)s — %(levelname)s — %(message)s",
        level=logging.DEBUG,
    )
    return log_file

def get_paths(configuration_path=None):
    """Get configuration and audio file paths"""
    if configuration_path and os.path.exists(configuration_path):
        audio_files_path = os.path.join(os.path.dirname(configuration_path), "Audio Files")
    else:
        if getattr(sys, "frozen", False):
            home_directory = os.path.expanduser("~")
            application_path = os.path.join(home_directory, "AppData", "Roaming", "Ace Centre", "AACSpeakHelper")
        else:
            application_path = os.path.dirname(__file__)
        
        audio_files_path = os.path.join(application_path, "Audio Files")
        configuration_path = os.path.join(application_path, "settings.cfg")
    
    os.makedirs(audio_files_path, exist_ok=True)
    return configuration_path, audio_files_path

def load_config():
    """Load configuration from settings.cfg"""
    global config, config_path, audio_files_path

    # Look for settings.cfg in current directory first, then AppData
    current_dir_config = os.path.join(os.path.dirname(__file__), "settings.cfg")
    if os.path.exists(current_dir_config):
        config_path = current_dir_config
        audio_files_path = os.path.join(os.path.dirname(__file__), "Audio Files")
    else:
        config_path, audio_files_path = get_paths()

    config = configparser.ConfigParser()

    if os.path.exists(config_path):
        config.read(config_path)
        logging.info(f"Loaded config from: {config_path}")
    else:
        logging.warning(f"Config file not found: {config_path}")
        # Create minimal default config
        config.add_section('TTS')
        config.set('TTS', 'engine', 'SherpaOnnxTTS')
        config.set('TTS', 'rate', '0')
        config.set('TTS', 'volume', '100')

    os.makedirs(audio_files_path, exist_ok=True)
    return config

def create_tts_engine(engine_name, voice_id=None):
    """Create TTS engine instance"""
    try:
        if engine_name == "azureTTS":
            key = config.get("azureTTS", "key", fallback="")
            location = config.get("azureTTS", "location", fallback="uksouth")
            voice = voice_id or config.get("azureTTS", "voice_id", fallback="en-GB-LibbyNeural")

            logging.info(f"Azure TTS config - key: {key[:10]}..., location: {location}, voice: {voice}")

            if not key:
                logging.error("Azure TTS key not configured")
                return None

            # New API: MicrosoftTTS expects credentials tuple
            credentials = (key, location)
            tts = MicrosoftTTS(credentials=credentials)
            # Set voice after creation
            tts.voice = voice
            return tts
            
        elif engine_name == "SherpaOnnxTTS":
            # SherpaOnnx doesn't use 'voice' parameter in constructor
            # The voice is set through the model path
            return SherpaOnnxTTS()
            
        elif engine_name == "googleTTS":
            creds = config.get("googleTTS", "creds", fallback="")
            voice = voice_id or config.get("googleTTS", "voice_id", fallback="en-US-Wavenet-C")
            
            if creds and os.path.exists(creds):
                return GoogleTTS(credentials=creds, voice=voice)
            else:
                logging.warning("Google TTS credentials not found")
                return None
                
        else:
            logging.error(f"Unknown TTS engine: {engine_name}")
            return None
            
    except Exception as e:
        logging.error(f"Error creating TTS engine {engine_name}: {e}")
        return None

def add_wav_header(pcm_data, sample_rate=22050, channels=1, bits_per_sample=16):
    """Add WAV header to raw PCM data"""
    import struct

    # Calculate sizes
    byte_rate = sample_rate * channels * bits_per_sample // 8
    block_align = channels * bits_per_sample // 8
    data_size = len(pcm_data)
    file_size = 36 + data_size

    # Create WAV header
    header = struct.pack('<4sI4s4sIHHIIHH4sI',
        b'RIFF',           # Chunk ID
        file_size,         # Chunk size
        b'WAVE',           # Format
        b'fmt ',           # Subchunk1 ID
        16,                # Subchunk1 size (PCM)
        1,                 # Audio format (PCM)
        channels,          # Number of channels
        sample_rate,       # Sample rate
        byte_rate,         # Byte rate
        block_align,       # Block align
        bits_per_sample,   # Bits per sample
        b'data',           # Subchunk2 ID
        data_size          # Subchunk2 size
    )

    return header + pcm_data

def speak_to_bytes(text, engine_name=None, voice_id=None):
    """Convert text to audio bytes with proper WAV headers using new synthesize() method"""
    try:
        engine_name = engine_name or config.get("TTS", "engine", fallback="SherpaOnnxTTS")
        tts = create_tts_engine(engine_name, voice_id)

        if not tts:
            logging.error(f"Failed to create TTS engine: {engine_name}")
            return None

        # Use new synthesize() method for silent audio generation (no playback)
        logging.info(f"🎤 Generating audio bytes silently for {engine_name}")

        # Get audio bytes without any audio playback using the new synthesize() method
        try:
            logging.debug("Using new synthesize() method")
            # Use the new synthesize method - it provides silent audio synthesis without playback
            audio_bytes = tts.synthesize(text, voice_id=voice_id)
        except Exception as e:
            logging.error(f"Audio generation failed: {e}")
            audio_bytes = None

        if audio_bytes:
            logging.info(f"✅ Generated {len(audio_bytes)} bytes of audio data")

            # Check if the audio data already has WAV headers
            if len(audio_bytes) >= 12 and audio_bytes[:4] == b'RIFF' and audio_bytes[8:12] == b'WAVE':
                logging.info("✅ Audio data already has WAV headers")
                return audio_bytes
            else:
                # Add WAV header to raw PCM data
                logging.info("Adding WAV header to raw PCM data")
                wav_data = add_wav_header(audio_bytes, sample_rate=22050, channels=1, bits_per_sample=16)
                logging.info(f"✅ Added WAV header, total size: {len(wav_data)} bytes")
                return wav_data
        else:
            logging.error("TTS engine returned no audio data")
            return None

    except Exception as e:
        logging.error(f"Error in speak_to_bytes: {e}")
        return None

def check_single_instance():
    """Check if another instance is already running"""
    mutex_name = "Global\\AACSpeakHelperServerMutex"
    try:
        win32event.CreateMutex(None, True, mutex_name)
        if win32api.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            logging.warning("Another instance is already running!")
            return False
        return True
    except Exception as e:
        logging.error(f"Error checking single instance: {e}")
        return True

class SystemTrayIcon(QSystemTrayIcon):
    def __init__(self, parent=None):
        # Use the configure icon from assets directory
        icon_path = os.path.join(os.path.dirname(__file__), "assets", "configure.ico")
        if os.path.exists(icon_path):
            icon = QIcon(icon_path)
            super().__init__(icon, parent)
        else:
            # Fallback to no icon if file not found
            logging.warning(f"Icon not found at: {icon_path}")
            super().__init__(parent)

        self.parent = parent
        menu = QMenu(parent)

        exitAction = menu.addAction("Exit")
        exitAction.triggered.connect(self.exit)
        self.setContextMenu(menu)
    
    def exit(self):
        if self.parent and hasattr(self.parent, 'pipe_thread'):
            self.parent.pipe_thread.quit()
        os._exit(0)

class PipeServerThread(QThread):
    """Windows named pipe server thread"""

    def run(self):
        pipe_name = r"\\.\pipe\AACSpeakHelper"
        logging.info("Starting AACSpeakHelper pipe server...")

        while True:
            pipe = None
            try:
                # Create named pipe
                pipe = win32pipe.CreateNamedPipe(
                    pipe_name,
                    win32pipe.PIPE_ACCESS_DUPLEX,
                    win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_READMODE_MESSAGE | win32pipe.PIPE_WAIT,
                    win32pipe.PIPE_UNLIMITED_INSTANCES,
                    65536, 65536, 0, None,
                )

                logging.info("Waiting for client connection...")
                win32pipe.ConnectNamedPipe(pipe, None)
                logging.info("Client connected.")

                # Read request
                result, data = win32file.ReadFile(pipe, 64 * 1024)
                if result == 0:
                    message = data.decode('utf-8')
                    logging.info(f"Received: {message[:50]}...")

                    # Process request directly (no Qt signals)
                    audio_bytes = self.process_request(message)

                    # Send response
                    import struct
                    length_prefix = struct.pack('<I', len(audio_bytes))
                    win32file.WriteFile(pipe, length_prefix + audio_bytes)
                    logging.info(f"Sent {len(audio_bytes)} bytes")

            except Exception as e:
                logging.error(f"Pipe error: {e}")
            finally:
                if pipe:
                    win32file.CloseHandle(pipe)

    def process_request(self, message):
        """Process a TTS request and return audio bytes"""
        try:
            data = json.loads(message)

            # Extract text and engine info - support both formats
            if "text" in data and "args" in data:
                # Simple test client format: {"text": "...", "args": {...}}
                text = data["text"]
                args = data["args"]
                engine = args.get("engine", "sherpaonnx")
                voice = args.get("voice", None)
            elif "clipboard_text" in data and "args" in data:
                # C++ wrapper format: {"clipboard_text": "...", "args": {...}, "config": {...}}
                text = data["clipboard_text"]
                args = data["args"]
                engine = args.get("engine", "sherpaonnx")
                voice = args.get("voice", None)
            else:
                logging.error("Invalid message format - missing text and args")
                return b""

            # Map engine names
            if engine == "azure":
                engine_name = "azureTTS"
            elif engine == "sherpaonnx":
                engine_name = "SherpaOnnxTTS"
            else:
                engine_name = "SherpaOnnxTTS"  # Default

            logging.info(f"Processing: '{text}' with {engine_name}")

            # Generate audio
            if engine_name == "SherpaOnnxTTS":
                audio_bytes = speak_to_bytes(text, engine_name, None)
            else:
                audio_bytes = speak_to_bytes(text, engine_name, voice)

            return audio_bytes or b""  # Return empty bytes if failed

        except Exception as e:
            logging.error(f"Error processing request: {e}")
            return b""

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.pipe_thread = None
        self.tray_icon = None
        self.init_ui()
        self.init_pipe_server()
    
    def init_ui(self):
        self.tray_icon = SystemTrayIcon(self)
        self.tray_icon.setVisible(True)
        self.tray_icon.setToolTip("AACSpeakHelper Server")
    
    def init_pipe_server(self):
        self.pipe_thread = PipeServerThread()
        self.pipe_thread.start()

def main():
    """Main entry point"""
    logfile = setup_logging()
    logging.info("=== AACSpeakHelper Server Starting ===")
    
    if not check_single_instance():
        print("Another instance is already running!")
        return
    
    # Load configuration
    load_config()
    
    # Create Qt application
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    
    # Create main window
    window = MainWindow()
    
    logging.info("Server started successfully")
    print("AACSpeakHelper Server is running...")
    print(f"Logs: {logfile}")
    
    # Run the application
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
