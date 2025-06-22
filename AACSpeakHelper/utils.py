import logging
import os
import sys
import subprocess
import time
import io
import uuid
import posthog
from PySide6.QtWidgets import *
from PySide6.QtCore import *
import sqlite3
import wave
import pyaudio
import warnings

warnings.filterwarnings("ignore")
args = {
    "config": "",
    "listvoices": False,
    "preview": False,
    "style": "",
    "styledegree": None,
}
config_path = None
audio_files_path = None
config = None


def ynbox(message: str, header: str, timeout: int = 10000):
    """Display a QMessageBox with Yes or No Option.

    Args:
        message (str): Display the question answerable by Yes or No.
        header (str): Text which will be display on the Header.
        timeout (int): Time in milliseconds will take for the QMessageBox to close without user interaction.
    Returns: bool
    """
    try:

        ynInstance = QMessageBox(None)
        ynInstance.setWindowTitle(header)
        ynInstance.setText(message)
        ynInstance.setStandardButtons(QMessageBox.StandardButton.Yes)
        ynInstance.addButton(QMessageBox.StandardButton.No)
        timer = QTimer(None)
        timer.singleShot(
            timeout,
            lambda: ynInstance.button(QMessageBox.StandardButton.No).animateClick(),
        )
        return ynInstance.exec() == QMessageBox.StandardButton.Yes
    except Exception as error:
        logging.error("Message Error: {}".format(error), exc_info=True)


def msgbox(message: str, header: str, timeout: int = 10000):
    """Display a QMessageBox for Notification only.

    Args:
        message (str): Text which will be display as notification.
        header (str): Text which will be display on the Header.
        timeout (int): Time in milliseconds will take for the QMessageBox to close without user interaction.
    Returns: bool
    """
    try:
        msgInstance = QMessageBox(None)
        msgInstance.setWindowTitle(header)
        msgInstance.setText(message)
        msgInstance.setStandardButtons(QMessageBox.StandardButton.Ok)
        timer = QTimer(None)
        timer.singleShot(
            timeout,
            lambda: msgInstance.button(QMessageBox.StandardButton.Ok).animateClick(),
        )
        return msgInstance.exec() == QMessageBox.StandardButton.Ok
    except Exception as error:
        logging.error("Message Error: {}".format(error), exc_info=True)


def configure_app():
    """Determine if application is a script file or frozen exe.

    Returns: None
    """

    if getattr(sys, "frozen", False):
        application_path = os.path.dirname(sys.executable)
        exe_name = ""
        for root, dirs, files in os.walk(application_path):
            for file in files:
                if "Configure AACSpeakHelper" in file:
                    exe_name = file
        GUI_path = os.path.join(application_path, exe_name)
        # Use subprocess.Popen to run the executable
        process = subprocess.Popen(GUI_path)
        # Wait for the process to complete
        process.wait()
    elif __file__:
        application_path = os.path.dirname(__file__)
        GUI_script_path = os.path.join(
            application_path, "GUI_TranslateAndTTS", "widget.py"
        )
        process = subprocess.run(["python", GUI_script_path])


def get_paths(configuration_path=None):
    """Get all the paths the application will be used during runtime.
    Args:
        configuration_path: Configuration file path
    Returns: Tuple
    """
    if configuration_path and os.path.exists(configuration_path):
        audio_files_path = os.path.join(
            os.path.dirname(configuration_path), "Audio Files"
        )
    else:
        if getattr(sys, "frozen", False):
            home_directory = os.path.expanduser("~")
            application_path = os.path.join(
                home_directory, "AppData", "Roaming", "Ace Centre", "AACSpeakHelper"
            )
        else:
            application_path = os.path.dirname(__file__)

        audio_files_path = os.path.join(application_path, "Audio Files")
        configuration_path = os.path.join(application_path, "settings.cfg")

    # Ensure the audio files directory exists
    os.makedirs(audio_files_path, exist_ok=True)

    return configuration_path, audio_files_path


def play_audio(audio_bytes, file: bool = False):
    if file:
        with wave.open(audio_bytes, "rb") as wf:
            play_wave(wf)
    else:
        with wave.open(io.BytesIO(audio_bytes), "rb") as wf:
            play_wave(wf)


def play_wave(wf):
    p = pyaudio.PyAudio()

    def callback(in_data, frame_count, time_info, status):
        data = wf.readframes(frame_count)
        return data, pyaudio.paContinue

    stream = p.open(
        format=p.get_format_from_width(wf.getsampwidth()),
        channels=wf.getnchannels(),
        rate=wf.getframerate(),
        output=True,
        stream_callback=callback,
    )

    stream.start_stream()

    while stream.is_active():
        pass

    stream.stop_stream()
    stream.close()
    wf.close()

    p.terminate()


def save_audio(text: str, engine: str, file_format: str = "wav", tts=None):
    """Save text as audio file with specific file format then save this text and audio file name in the database.
    If text is synthesize again, it will be find first in the database if there is a match.
    Args:
        text (str): Text String
        engine (str): Name of the TTS Engine
        file_format (str): File Format of the Audio e.g. 'wav' or 'mp3'
        tts: Instance of TTS Engine
    Returns: None
    """
    timestr = time.strftime("%Y%m%d-%H%M%S.")
    filename = os.path.join(audio_files_path, timestr + file_format)

    # Handle different TTS engines with different method signatures
    try:
        # Check if this is SherpaOnnx TTS and handle it specially
        from tts_wrapper import SherpaOnnxClient

        if isinstance(tts, SherpaOnnxClient):
            # For SherpaOnnx TTS, use the basic speak_streamed method first
            # and try to save to file separately if supported
            try:
                tts.speak_streamed(text, save_to_file_path=filename)
            except (TypeError, AttributeError):
                # If save_to_file_path is not supported, just play the audio
                tts.speak_streamed(text)
                # Note: This won't save to file, but at least it will play the audio
        else:
            # Try the method with save parameters first (for engines that support it)
            tts.speak_streamed(text, save_to_file_path=filename, audio_format=file_format)
    except (TypeError, Exception) as e:
        # If that fails, try different approaches based on the TTS engine type
        from tts_wrapper import MicrosoftTTS

        if isinstance(tts, MicrosoftTTS):
            # For Azure TTS, use speak_to_file method if available
            if hasattr(tts, 'speak_to_file'):
                tts.speak_to_file(text, filename)
            else:
                # Fallback: get audio data and save manually
                audio_data = tts.speak(text)
                with open(filename, 'wb') as f:
                    f.write(audio_data)
        else:
            # For other engines, try the basic speak_streamed method
            tts.speak_streamed(text)
            # Note: This won't save to file, but at least it will play the audio

    sql = "INSERT INTO History(text, filename, engine) VALUES('{}','{}','{}')".format(
        text, filename, engine
    )
    try:
        connection = sqlite3.connect(os.path.join(audio_files_path, "cache_history.db"))
        connection.execute(sql)
        connection.commit()
        connection.close()
    except Exception as error:
        logging.error("Database Error: {}".format(error), exc_info=True)


def get_uuid():
    """Generates random UUID or loads UUID from configuration file.
    Remove uuid config every commit
    Code will raise an exception at first run due to blank uuid

    Returns: str
    """

    try:
        identifier = uuid.UUID(config.get("App", "uuid"))
    except Exception as error:
        identifier = uuid.uuid4()
        # config.set("App", "uuid", str(identifier))
        # with open(config_path, "w") as configfile:
        #     config.write(configfile)
        logging.error("Failed to get uuid: {}".format(error), exc_info=True)
    logging.info("uuid: {}".format(identifier), exc_info=True)
    return str(identifier)


def notify_posthog(id: str, event_name: str, properties: dict = {}):
    """Save text as audio file with specific file format then save this text and audio file name in the database.
    If text is synthesize again, it will be find first in the database if there is a match.
    Args:
        id (str): UUID
        event_name (str): Event Name
        properties (dict): Properties in dictionary format.
    Returns: None
    """

    try:
        posthog_client = posthog.Posthog(
            project_api_key="phc_q37FBcmTQD1hHtNBgqvs9wid45gKjGKEJGduRkPog0t",
            host="https://app.posthog.com",
        )
        # Attempt to send the event to PostHog
        posthog_client.capture(distinct_id=id, event=event_name, properties=properties)
        print(f"[notify-posthog] Event '{event_name}' captured successfully!")
    except Exception as e:
        # Handle the case when there's an issue with sending the event
        print(f"[notify-posthog] Failed to capture event '{event_name}': {e}")
        logging.error(
            "[notify-posthog] Failed to capture event '{}': {}".format(event_name, e),
            exc_info=True,
        )
        # You can add further logic here if needed, such as logging the error or continuing the script
        pass


def check_history(text: str):
    """Check for a matching string in the database and return the file path.
    If no database was found, the function will call the create_Database function and return None
    Args:
        text (str): Text to be match in the database
    Returns: str
    """
    try:
        if args["style"]:
            return None
        if os.path.isfile(os.path.join(audio_files_path, "cache_history.db")):
            sql = "SELECT filename FROM History WHERE text='{}'".format(text)
            connection = sqlite3.connect(
                os.path.join(audio_files_path, "cache_history.db")
            )
            cursor = connection.execute(sql)
            results = cursor.fetchone()
            base_name = results[0] if results is not None else None
            if base_name is not None:
                file = os.path.join(audio_files_path, base_name)
                connection.close()
                return file
            else:
                return None
        else:
            create_Database()
            return None
    except Exception:
        logging.error("Failed to connect to database: ", exc_info=True)
        return None


def clear_history(files: list):
    """Check for a matching string in the database and return the file path.
    If no database was found, the function will call the create_Database function and return None
    Args:
        files (list): List of file to be deleted.
    Returns: None
    """
    try:
        if (
            os.path.isfile(os.path.join(audio_files_path, "cache_history.db"))
            and len(files) > 0
        ):
            connection = sqlite3.connect(
                os.path.join(audio_files_path, "cache_history.db")
            )
            for file in files:
                sql = "DELETE FROM History WHERE filename='{}'".format(file)
                connection.execute(sql)
            connection.commit()
            connection.close()
    except Exception:
        logging.error("Failed to connect to database: ", exc_info=True)
        return None


def create_Database():
    """If no database was found, this function will create database function.

    Returns: None
    """
    try:
        if not os.path.isfile(os.path.join(audio_files_path, "cache_history.db")):
            sql1 = """CREATE TABLE IF NOT EXISTS "History" ("id"	INTEGER NOT NULL UNIQUE,
                                                            "text"	TEXT NOT NULL,
                                                            "filename"	TEXT NOT NULL,
                                                            "engine"	TEXT NOT NULL,
                                                            UNIQUE("id"),
                                                            PRIMARY KEY("id" AUTOINCREMENT));"""
            sql2 = (
                """CREATE UNIQUE INDEX IF NOT EXISTS "id_text" ON "History" ("text");"""
            )
            connection = sqlite3.connect(
                os.path.join(audio_files_path, "cache_history.db")
            )
            connection.execute(sql1)
            connection.execute(sql2)
            connection.close()
            logging.info("Cache database is created.")
        else:
            logging.info("Cache database is found")
    except Exception:
        logging.error("Failed to create database: ", exc_info=True)


def init(input_config, args):
    """Initialize configuration file path making it in memory instead of one time instance.

    Args:
        input_config: configuration file path.
    Returns: None
    """

    global config_path
    global audio_files_path
    global config

    # Enhanced debug logging for frozen executable issues
    logging.info("=== Utils.init() Debug Information ===")
    logging.info(f"Input config type: {type(input_config)}")
    logging.info(f"Input config sections: {list(input_config.keys()) if hasattr(input_config, 'keys') else 'Not dict-like'}")
    logging.info(f"Args: {args}")

    # Log App section details
    if "App" in input_config:
        logging.info(f"App section contents: {dict(input_config['App'])}")
    else:
        logging.error("No 'App' section found in input_config!")

    # Log TTS section details
    if "TTS" in input_config:
        logging.info(f"TTS section contents: {dict(input_config['TTS'])}")
    else:
        logging.error("No 'TTS' section found in input_config!")

    # Log azureTTS section details
    if "azureTTS" in input_config:
        logging.info(f"azureTTS section contents: {dict(input_config['azureTTS'])}")
    else:
        logging.error("No 'azureTTS' section found in input_config!")

    config_path = input_config["App"]["config_path"]
    audio_files_path = input_config["App"]["audio_files_path"]
    config = (
        input_config  # This assigns the passed config to the global config variable
    )

    logging.info(f"Initialized utils with config path: {config_path}")
    logging.info(f"Audio files path: {audio_files_path}")
    logging.info("=== End Utils.init() Debug Information ===")

    if config.getboolean("App", "collectstats"):
        distinct_id = get_uuid()
        event_name = "App Run"
        event_properties = {
            "uuid": distinct_id,
            "source": "helperApp",
            "version": 2.4,
            "fromLang": config.get("translate", "start_lang"),
            "toLang": config.get("translate", "end_lang"),
            "ttsengine": config.get("TTS", "engine"),
        }
        notify_posthog(distinct_id, event_name, event_properties)
