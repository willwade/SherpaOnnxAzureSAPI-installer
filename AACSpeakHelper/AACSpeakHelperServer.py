"""
AACSpeakHelperServer.py - Windows Named Pipe Server for AACSpeakHelper

This server provides text-to-speech (TTS) and translation services through a Windows named pipe.
It runs as a system tray application and processes requests from client applications.

The server:
1. Creates a named pipe (\\\\.\\pipe\\AACSpeakHelper) to receive requests
2. Processes incoming JSON messages containing text to speak/translate
3. Performs translation if requested
4. Converts text to speech using the configured TTS engine
5. Optionally replaces clipboard content with translated text

Credentials:
- All API keys and credentials are read from the settings.cfg file
- No environment variables are required for normal operation
- The configuration file is typically located in the user's AppData folder

Usage:
- Run this script to start the server
- Use client.py to send requests to the server
- Configure settings using the GUI or CLI configuration tools

Note for developers:
- During development, credentials can be stored in .envrc
- These are used by test scripts, not by the server itself

Author: Ace Centre
"""

import logging
import os
import sys
import warnings
import unicodedata

warnings.filterwarnings("ignore")


def setup_logging():
    if getattr(sys, "frozen", False):
        # If the application is run as a bundle, use the AppData directory
        log_dir = os.path.join(
            os.path.expanduser("~"),
            "AppData",
            "Roaming",
            "Ace Centre",
            "AACSpeakHelper",
        )
    else:
        # If run from a Python environment, use the current directory
        log_dir = os.path.dirname(os.path.abspath(__file__))

    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    log_file = os.path.join(log_dir, "app.log")

    logging.basicConfig(
        filename=log_file,
        filemode="a",
        format="%(asctime)s — %(name)s — %(levelname)s — %(funcName)s:%(lineno)d — %(message)s",
        level=logging.DEBUG,
    )

    return log_file


logfile = setup_logging()

def log_debug_info():
    """Log comprehensive debug information for frozen vs development differences"""
    logging.info("=== AACSpeakHelper Server Debug Information ===")
    logging.info(f"Python executable: {sys.executable}")
    logging.info(f"Python version: {sys.version}")
    logging.info(f"Current working directory: {os.getcwd()}")
    logging.info(f"Script location: {__file__}")
    logging.info(f"Frozen executable: {getattr(sys, 'frozen', False)}")
    if hasattr(sys, '_MEIPASS'):
        logging.info(f"PyInstaller temp dir: {sys._MEIPASS}")
    logging.info(f"Python path: {sys.path}")

    # Log file system information
    current_dir = os.getcwd()
    try:
        files_in_dir = os.listdir(current_dir)
        logging.info(f"Files in current directory: {files_in_dir}")
    except Exception as e:
        logging.error(f"Error listing current directory: {e}")

    # Check for config files
    config_files = ['settings.cfg', 'config.cfg', '.envrc']
    for config_file in config_files:
        config_path = os.path.join(current_dir, config_file)
        if os.path.exists(config_path):
            try:
                size = os.path.getsize(config_path)
                logging.info(f"Found config file: {config_path} (size: {size} bytes)")
            except Exception as e:
                logging.error(f"Error checking config file {config_path}: {e}")
        else:
            logging.info(f"Config file not found: {config_path}")

    # Check environment variables
    important_env_vars = ['PATH', 'PYTHONPATH', 'APPDATA', 'USERPROFILE']
    for var in important_env_vars:
        value = os.environ.get(var, 'Not set')
        logging.info(f"Environment {var}: {value}")

    logging.info("=== End Debug Information ===")

# Call debug logging immediately
log_debug_info()

import json
import time
import threading
import pyperclip
import win32file
import win32pipe
import win32event
import win32api
from PySide6.QtWidgets import QApplication, QWidget, QSystemTrayIcon, QMenu, QMessageBox
from PySide6.QtGui import QIcon, QAction
from PySide6.QtCore import QThread, Signal, Slot, QTimer
from deep_translator import *
import utils
import tts_utils
import subprocess
import configparser


def check_single_instance():
    """
    Check if another instance of AACSpeakHelper server is already running.
    Uses a named mutex to ensure only one instance can run at a time.

    Returns:
        bool: True if this is the only instance, False if another instance is already running
    """
    mutex_name = "Global\\AACSpeakHelperServerMutex"

    try:
        # Try to create a named mutex
        win32event.CreateMutex(None, True, mutex_name)

        if win32api.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            logging.warning("Another instance of AACSpeakHelper server is already running!")
            return False
        else:
            logging.info("Single instance check passed - this is the only server instance")
            return True

    except Exception as e:
        logging.error(f"Error checking single instance: {e}")
        # If we can't check, assume it's safe to continue
        return True


class SystemTrayIcon(QSystemTrayIcon):
    def __init__(self, icon, parent=None):
        super().__init__(icon, parent)
        self.parent = parent
        menu = QMenu(parent)
        self.config_path, self.audio_files_path = utils.get_paths(None)
        logging.info(f"Config path: {self.config_path}")
        logging.info(f"Audio files path: {self.audio_files_path}")

        openLogsAction = QAction("Open logs", self)
        menu.addAction(openLogsAction)
        openLogsAction.triggered.connect(self.open_logs)

        openCacheAction = QAction("Open Cache", self)
        menu.addAction(openCacheAction)
        openCacheAction.triggered.connect(self.open_cache)

        menu.addSeparator()

        self.lastRunAction = QAction("Last run info not available", self)
        self.lastRunAction.setEnabled(False)
        menu.addAction(self.lastRunAction)

        exitAction = menu.addAction("Exit")
        exitAction.triggered.connect(self.exit)

        self.setContextMenu(menu)

    def exit(self):
        self.parent.pipe_thread.quit()
        os._exit(0)
        # QApplication.quit()

    def update_last_run_info(self, last_run_time, duration):
        self.lastRunAction.setText(
            f"Last run at {last_run_time} - took {duration} secs"
        )

    def open_logs(self):
        logging.info("Opening logs...")
        subprocess.Popen(["notepad", logfile])

    def open_cache(self):
        logging.info("Opening cache...")
        subprocess.Popen(["explorer", self.audio_files_path])
        # Implement cache opening logic here


class PipeServerThread(QThread):
    """
    Thread that handles the Windows named pipe communication.

    This thread:
    1. Creates a named pipe (\\\\.\\pipe\\AACSpeakHelper)
    2. Waits for client connections
    3. Reads data from clients
    4. Emits a signal with the received message
    5. Optionally returns voice list data or audio bytes to the client

    The pipe accepts JSON-formatted messages containing:
    - args: Command-line arguments from the client
    - config: Configuration settings
    - clipboard_text: Text to process

    Required Credentials (from settings.cfg):
    - For Azure TTS: key, location in [azureTTS] section
    - For Google TTS: creds in [googleTTS] section
    - For Microsoft Translator: microsoft_translator_secret_key in [translate] section
    - For Microsoft Translator: region in [translate] section
    """

    message_received = Signal(str, object)  # message, pipe_handle
    voices = None
    audio_bytes = None

    def run(self):
        pipe_name = r"\\.\pipe\AACSpeakHelper"
        while True:
            pipe = None
            try:
                # Create the named pipe
                pipe = win32pipe.CreateNamedPipe(
                    pipe_name,
                    win32pipe.PIPE_ACCESS_DUPLEX,
                    win32pipe.PIPE_TYPE_MESSAGE
                    | win32pipe.PIPE_READMODE_MESSAGE
                    | win32pipe.PIPE_WAIT,
                    win32pipe.PIPE_UNLIMITED_INSTANCES,
                    65536,
                    65536,
                    0,
                    None,
                )

                logging.info("Waiting for client connection...")
                # Wait for a client to connect
                win32pipe.ConnectNamedPipe(pipe, None)
                logging.info("Client connected.")

                # Read data from the client
                result, data = win32file.ReadFile(pipe, 64 * 1024)
                get_voices = False
                if result == 0:
                    # Handle different message formats
                    try:
                        # First, try to decode as regular UTF-8 JSON
                        if data.startswith(b'{'):
                            # Looks like direct JSON
                            message = data.decode('utf-8')
                            logging.info("PipeServerThread: Detected direct JSON message")
                        elif len(data) >= 4:
                            # Check if data starts with what looks like a length prefix
                            import struct
                            potential_length = struct.unpack('<I', data[0:4])[0]  # Little-endian uint32

                            # If the length makes sense and matches remaining data, it's length-prefixed
                            if potential_length > 0 and potential_length <= len(data) - 4:
                                logging.info(f"PipeServerThread: Detected length-prefixed message, length: {potential_length}")
                                # Extract the actual message after the length prefix
                                message_data = data[4:4+potential_length]
                                message = message_data.decode('utf-8')
                            else:
                                # Try to decode the whole thing as UTF-8
                                message = data.decode('utf-8')
                        else:
                            # Too short for length prefix, decode as regular message
                            message = data.decode('utf-8')
                    except (struct.error, UnicodeDecodeError) as e:
                        logging.warning(f"Failed to parse message format: {e}")
                        # Last resort: try different encodings or skip problematic bytes
                        try:
                            # Try UTF-8 with error handling
                            message = data.decode('utf-8', errors='ignore')
                            logging.warning("Decoded message with ignored errors")
                        except Exception as e2:
                            logging.error(f"Failed to decode message. Data starts with: {data[:10].hex()}")
                            logging.error(f"Data length: {len(data)}")
                            continue

                    logging.info(f"PipeServerThread: Received data: {message[:50]}...")
                    logging.info(f"PipeServerThread: Processing message in thread {threading.current_thread().name}")

                    # Reset response data
                    self.voices = None
                    self.audio_bytes = None

                    # Emit signal to process the message in the main thread
                    self.message_received.emit(message, pipe)

                    # Wait for the main thread to process and set response data
                    # Give it up to 30 seconds to complete TTS processing
                    max_wait_time = 30.0
                    wait_interval = 0.1
                    elapsed_time = 0.0

                    while elapsed_time < max_wait_time:
                        time.sleep(wait_interval)
                        elapsed_time += wait_interval

                        # Check if we have a response ready
                        if self.voices is not None or self.audio_bytes is not None:
                            logging.info(f"Response ready after {elapsed_time:.1f} seconds")
                            break

                    if elapsed_time >= max_wait_time:
                        logging.warning("Timeout waiting for TTS processing to complete")
                    try:
                        # Check if client is requesting voice list
                        parsed_message = json.loads(message)
                        get_voices = parsed_message.get("args", {}).get(
                            "listvoices", False
                        )
                    except Exception as e:
                        logging.error(f"Error parsing message: {e}")
                        get_voices = False
                logging.info("Processing complete. Ready for next connection.")
            except Exception as e:
                logging.error(f"Pipe server error: {e}", exc_info=True)
            finally:
                if pipe:
                    # Send responses back to client
                    try:
                        # If client requested voices and we have them, send them back
                        if get_voices and self.voices:
                            win32file.WriteFile(pipe, json.dumps(self.voices).encode())
                            self.voices = None
                            logging.info("Sent voice list back through pipe")

                        # If we have audio bytes to send back, send them
                        elif self.audio_bytes:
                            # Send audio bytes with length prefix for reliable transmission
                            import struct
                            audio_length = len(self.audio_bytes)
                            length_prefix = struct.pack('<I', audio_length)  # Little-endian uint32
                            win32file.WriteFile(pipe, length_prefix + self.audio_bytes)
                            self.audio_bytes = None
                            logging.info(f"Sent {audio_length} bytes of audio data back through pipe")

                    except Exception as e:
                        logging.error(f"Error sending response through pipe: {e}")

                    # Close the pipe handle
                    win32file.CloseHandle(pipe)
                logging.info("Pipe closed. Reopening for next connection.")


class CacheCleanerThread(QThread):
    def run(self):
        remove_stale_temp_files(utils.audio_files_path)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.cache_timer = None
        self.cache_cleaner = None
        self.pipe_thread = None
        self.tray_icon = None
        self.icon = QIcon("assets/translate.ico")
        self.icon_loading = QIcon("assets/translate_loading.ico")
        self.init_ui()
        self.init_pipe_server()
        self.init_cache_cleaner()
        self.init_log_cleaner()
        self.tray_icon.setToolTip("Waiting for new client...")

    def init_ui(self):
        self.tray_icon = SystemTrayIcon(self.icon, self)
        self.tray_icon.setVisible(True)

    def init_pipe_server(self):
        self.pipe_thread = PipeServerThread()
        self.pipe_thread.message_received.connect(self.handle_message)
        self.pipe_thread.start()

    def init_cache_cleaner(self):
        self.cache_cleaner = CacheCleanerThread()
        self.cache_timer = QTimer(self)
        self.cache_timer.timeout.connect(lambda: self.cache_cleaner.start())
        self.cache_timer.start(24 * 60 * 60 * 1000)  # Run once a day

    def init_log_cleaner(self):
        self.log_timer = QTimer(self)
        self.log_timer.timeout.connect(self.check_log_size)
        self.log_timer.start(3600000)  # Check every hour, adjust as needed

    def check_log_size(self):
        log_file = logfile
        size_limit_mb = 1  # Size limit in MB
        size_limit_bytes = size_limit_mb * 1024 * 1024  # Convert MB to bytes

        try:
            if (
                os.path.isfile(log_file)
                and os.path.getsize(log_file) > size_limit_bytes
            ):
                with open(log_file, "w"):  # Truncate the log file
                    logging.info("Log file exceeded size limit; log file truncated.")
        except Exception as e:
            logging.error(f"Error cleaning log file: {e}", exc_info=True)

    @Slot(str, object)
    def handle_message(self, message, pipe_handle):
        """
        Process a message received from a client.

        This method:
        1. Parses the JSON message
        2. Extracts configuration and text data
        3. Initializes the TTS and translation engines
        4. Translates the text if requested
        5. Speaks the text using the configured TTS engine
        6. Optionally updates the clipboard with translated text
        7. Sends audio bytes back through pipe if requested

        Args:
            message (str): JSON-formatted message from the client
            pipe_handle: Windows pipe handle for sending responses back
        """
        logging.info(f"MainWindow.handle_message: Starting to process message in thread {threading.current_thread().name}")
        try:
            # Parse the JSON message
            data = json.loads(message)

            # Support both new format (args/config/clipboard_text) and old format (text/args)
            if "clipboard_text" in data and "config" in data:
                # New AACSpeakHelper format
                args = data["args"]
                config_dict = data["config"]
                clipboard_text = data["clipboard_text"]
            elif "text" in data and "args" in data:
                # Legacy COM wrapper format - convert to new format
                logging.info("Converting legacy COM wrapper message format")
                args = data["args"]
                clipboard_text = data["text"]

                # Create default config based on engine type
                engine = args.get("engine", "sherpaonnx")
                if engine == "sherpaonnx":
                    config_dict = {
                        "TTS": {
                            "engine": "SherpaOnnxTTS",
                            "bypass_tts": "False",
                            "save_audio_file": "True",
                            "rate": str(args.get("rate", 0)),
                            "volume": str(args.get("volume", 100))
                        },
                        "translate": {
                            "no_translate": "True"
                        },
                        "App": {
                            "collectstats": "True"
                        }
                    }
                elif engine == "azure":
                    config_dict = {
                        "TTS": {
                            "engine": "azureTTS",
                            "bypass_tts": "False",
                            "save_audio_file": "True",
                            "rate": str(args.get("rate", 0)),
                            "volume": str(args.get("volume", 100))
                        },
                        "translate": {
                            "no_translate": "True"
                        },
                        "App": {
                            "collectstats": "True"
                        }
                    }
                else:
                    # Default fallback
                    config_dict = {
                        "TTS": {
                            "engine": "SherpaOnnxTTS",
                            "bypass_tts": "False",
                            "save_audio_file": "True",
                            "rate": "0",
                            "volume": "100"
                        },
                        "translate": {
                            "no_translate": "True"
                        },
                        "App": {
                            "collectstats": "True"
                        }
                    }
            else:
                raise ValueError("Message must contain either (clipboard_text, config, args) or (text, args)")

            # Create a ConfigParser object and update it with received config
            config = configparser.ConfigParser()
            for section, options in config_dict.items():
                config[section] = options

            # Log the Google credentials path (if available)
            if "googleTTS" in config and "creds" in config["googleTTS"]:
                logging.info(
                    f"Using Google credentials: {config['googleTTS']['creds']}"
                )

            # Get the config path from the received config
            config_path = config.get("App", "config_path", fallback=None)
            # Use utils.get_paths to get the paths
            # config_path, audio_files_path = utils.get_paths(config_path)
            # TODO: Disable config_path for now
            config_path, audio_files_path = utils.get_paths()

            if "App" not in config:
                config["App"] = {}
            config["App"]["config_path"] = config_path
            config["App"]["audio_files_path"] = audio_files_path
            # Initialize utils with the new config and args
            utils.init(config, args)

            # Initialize TTS
            tts_utils.init(utils)
            # Process the clipboard text
            if not tts_utils.ready:
                logging.info(
                    "Application is not ready. Please wait until current session is finished."
                )
                return
            self.tray_icon.setToolTip("Handling new message ...")
            self.tray_icon.setIcon(self.icon_loading)
            logging.info(f"Handling new message: {message[:50]}...")
            if config.getboolean("translate", "no_translate"):
                text_to_process = clipboard_text
            else:
                text_to_process = translate_clipboard(clipboard_text, config)

            # Perform TTS if not bypassed
            if not config.getboolean("TTS", "bypass_tts", fallback=False):
                # Check if client wants audio bytes back (for SAPI integration)
                return_audio_bytes = args.get("return_audio_bytes", False)
                if return_audio_bytes:
                    # Get audio bytes and store them for pipe thread to send back
                    audio_bytes = tts_utils.speak_to_bytes(text_to_process, args.get("listvoices", False))
                    if audio_bytes:
                        self.pipe_thread.audio_bytes = audio_bytes
                        logging.info(f"Generated {len(audio_bytes)} bytes of audio data for pipe transmission")
                    else:
                        logging.error("Failed to generate audio bytes")
                        self.pipe_thread.audio_bytes = None
                else:
                    tts_utils.speak(text_to_process, args.get("listvoices", False))
            if tts_utils.voices:
                self.pipe_thread.voices = tts_utils.voices
            # Replace clipboard if specified
            if (
                config.getboolean("translate", "replace_pb")
                and text_to_process is not None
            ):
                pyperclip.copy(text_to_process)

            current_time = time.strftime("%Y-%m-%d %H:%M:%S")
            self.tray_icon.update_last_run_info(current_time, "N/A")
            logging.info(f"Processed message at {current_time}")
            self.tray_icon.setIcon(self.icon)
            self.tray_icon.setToolTip("Waiting for new client...")
        except Exception as e:
            logging.error(f"Error handling message: {e}", exc_info=True)
        finally:
            logging.info("Message handling complete.")


def translate_clipboard(text, config):
    try:
        translator = config.get("translate", "provider")
        key = (
            config.get("translate", f"{translator}_secret_key")
            if not translator == "GoogleTranslator"
            else None
        )
        email = (
            config.get("translate", "email")
            if translator == "MyMemoryTranslator"
            else None
        )
        region = (
            config.get("translate", "region")
            if translator == "MicrosoftTranslator"
            else None
        )
        pro = (
            config.getboolean("translate", "deepl_pro")
            if translator == "DeeplTranslator"
            else None
        )
        url = config.get("translate", "url") if translator == "LibreProvider" else None
        client_id = (
            config.get("translate", "papagotranslator_client_id")
            if translator == "PapagoTranslator"
            else None
        )
        appid = (
            config.get("translate", "baidutranslator_appid")
            if translator == "BaiduTranslator"
            else None
        )

        match translator:
            case "GoogleTranslator":
                translate_instance = GoogleTranslator(
                    source="auto", target=config.get("translate", "end_lang")
                )
            case "PonsTranslator":
                translate_instance = PonsTranslator(
                    source="auto", target=config.get("translate", "end_lang")
                )
            case "LingueeTranslator":
                translate_instance = LingueeTranslator(
                    source="auto", target=config.get("translate", "end_lang")
                )
            case "MyMemoryTranslator":
                translate_instance = MyMemoryTranslator(
                    source=config.get("translate", "start_lang"),
                    target=config.get("translate", "end_lang"),
                    email=email,
                )
            case "YandexTranslator":
                translate_instance = YandexTranslator(
                    source=config.get("translate", "start_lang"),
                    target=config.get("translate", "end_lang"),
                    api_key=key,
                )
            case "MicrosoftTranslator":
                translate_instance = MicrosoftTranslator(
                    api_key=key,
                    source=config.get("translate", "start_lang"),
                    target=config.get("translate", "end_lang"),
                    region=region,
                )
            case "QcriTranslator":
                translate_instance = QcriTranslator(
                    source="auto",
                    target=config.get("translate", "end_lang"),
                    api_key=key,
                )
            case "DeeplTranslator":
                translate_instance = DeeplTranslator(
                    source=config.get("translate", "start_lang"),
                    target=config.get("translate", "end_lang"),
                    api_key=key,
                    use_free_api=not pro,
                )
            case "LibreTranslator":
                translate_instance = LibreTranslator(
                    source=config.get("translate", "start_lang"),
                    target=config.get("translate", "end_lang"),
                    api_key=key,
                    custom_url=url,
                )
            case "PapagoTranslator":
                translate_instance = PapagoTranslator(
                    source="auto",
                    target=config.get("translate", "end_lang"),
                    client_id=client_id,
                    secret_key=key,
                )
            case "ChatGptTranslator":
                translate_instance = ChatGptTranslator(
                    source="auto", target=config.get("translate", "end_lang")
                )
            case "BaiduTranslator":
                translate_instance = BaiduTranslator(
                    source=config.get("translate", "start_lang"),
                    target=config.get("translate", "end_lang"),
                    appid=appid,
                    appkey=key,
                )
        # elif translator == "DeepLearningTranslator":
        #     translate_instance = BaiduTranslator(source=config.get('translate', 'startlang'),
        #                                          target=config.get('translate', 'endLang'),
        #                                          appid=appid,
        #                                          appkey=key)
        logging.info("Translation Provider is {}".format(translator))
        logging.info(f'Text [{config.get("translate", "start_lang")}]: {text}')
        if config.get("translate", "end_lang") in [
            "ckb" "ku",
            "kmr",
            "kmr-TR",
            "ckb-IQ",
        ]:
            text = normalize_text(text)
        translation = translate_instance.translate(text)
        logging.info(
            f'Translation [{config.get("translate", "end_lang")}]: {translation}'
        )
        return translation
    except Exception as e:
        logging.error(f"Translation Error: {e}", exc_info=True)


def normalize_text(text: str):
    normalizedText = unicodedata.normalize("NFC", text)
    logging.info("Normalized Text: {}".format(normalizedText))
    return normalizedText


def remove_stale_temp_files(directory_path, ignore_pattern=".db"):
    config = utils.config
    start = time.perf_counter()
    current_time = time.time()
    day = int(config.get("appCache", "threshold"))
    time_threshold = current_time - day * 24 * 60 * 60
    file_list = []

    for root, dirs, files in os.walk(directory_path):
        for file in files:
            file_path = os.path.join(root, file)
            if (
                ignore_pattern
                and file.endswith(ignore_pattern)
                and file.endswith(".db-journal")
            ):
                continue
            try:
                file_modification_time = os.path.getmtime(file_path)
                if file_modification_time < time_threshold:
                    os.remove(file_path)
                    file_list.append(os.path.basename(file_path))
                    logging.info(f"Removed cache file: {file_path}")
            except Exception as e:
                logging.error(f"Error processing file {file_path}: {e}", exc_info=True)

    stop = time.perf_counter() - start
    utils.clear_history(file_list)
    logging.info(f"Cache clearing took {stop:0.5f} seconds.")


def clearCache():
    temp_folder = os.getenv("TEMP")
    size_limit = 5 * 1024  # 5KB in bytes
    # Scan the directory
    for root, dirs, files in os.walk(temp_folder):
        for file in files:
            if file.startswith("tmp"):
                file_path = os.path.join(root, file)
                if os.path.getsize(file_path) < size_limit:
                    current_time = time.time()
                    day = 7
                    time_threshold = current_time - day * 24 * 60 * 60
                    file_modification_time = os.path.getmtime(file_path)
                    if file_modification_time < time_threshold:
                        os.remove(file_path)


if __name__ == "__main__":
    # Check if another instance is already running
    if not check_single_instance():
        logging.error("Exiting: Another instance of AACSpeakHelper server is already running")

        # Show a message box if we're not in frozen mode (development)
        if not getattr(sys, "frozen", False):
            try:
                app = QApplication(sys.argv)
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Icon.Warning)
                msg.setWindowTitle("AACSpeakHelper Server")
                msg.setText("Another instance of AACSpeakHelper server is already running.")
                msg.setInformativeText("Please close the existing instance before starting a new one.")
                msg.exec()
            except Exception as e:
                logging.error(f"Error showing message box: {e}")

        sys.exit(1)

    clearCache()
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())
