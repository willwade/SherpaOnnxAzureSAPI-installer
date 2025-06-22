import logging
import os.path
import sys
import time
from pathlib import Path

from tts_wrapper import (
    AbstractTTS,
    MicrosoftTTS,
    GoogleClient,
    GoogleTTS,
    SAPIClient,
    SAPIEngine,
    SherpaOnnxClient,
    GoogleTransTTS,
    ElevenLabsClient,
    ElevenLabsTTS,
    PlayHTClient,
    PlayHTTTS,
    PollyClient,
    PollyTTS,
    WatsonClient,
    WatsonTTS,
    OpenAIClient,
    WitAiClient,
    WitAiTTS,
)
import warnings
from threading import Thread
from configure_enc_utils import load_config, load_credentials



warnings.filterwarnings("ignore", category=RuntimeWarning)
utils = None
voices = None
ready = True
# Global dictionary to store TTS clients
tts_voiceid = {}

VALID_STYLES = [
    "advertisement_upbeat",
    "affectionate",
    "angry",
    "assistant",
    "calm",
    "chat",
    "cheerful",
    "customerservice",
    "depressed",
    "disgruntled",
    "documentary-narration",
    "embarrassed",
    "empathetic",
    "envious",
    "excited",
    "fearful",
    "friendly",
    "gentle",
    "hopeful",
    "lyrical",
    "narration-professional",
    "narration-relaxed",
    "newscast",
    "newscast-casual",
    "newscast-formal",
    "poetry-reading",
    "sad",
    "serious",
    "shouting",
    "sports_commentary",
    "sports_commentary_excited",
    "whispering",
    "terrified",
    "unfriendly",
]


def init(module):
    """Initialize utils module making it in memory instead of one time instance.

    Args:
        module: Instance of utils module.
    Returns: None
    """
    global utils
    utils = module


def init_azure_tts():
    """Initialize unique instance of MicrosoftTTS based on the changes in voiceid.

    Returns: MicrosoftTTS
    """
    logging.info("=== init_azure_tts() Debug Information ===")

    # Debug: Check if utils.config exists and has azureTTS section
    if not hasattr(utils, 'config'):
        logging.error("utils.config is not available!")
        return None

    if not utils.config.has_section("azureTTS"):
        logging.error("azureTTS section not found in utils.config!")
        logging.info(f"Available sections: {utils.config.sections()}")
        return None

    # Get configuration values with debug logging
    key = utils.config.get("azureTTS", "key")
    logging.info(f"Azure TTS key: {key[:10]}... (length: {len(key)})")

    config = load_config()
    if key == "":
        key = config.get("azureTTS")["key"]
        logging.info(f"Using fallback key from load_config: {key[:10]}...")

    location = utils.config.get("azureTTS", "location")
    logging.info(f"Azure TTS location: {location}")
    if location == "":
        location = config.get("azureTTS")["location"]
        logging.info(f"Using fallback location from load_config: {location}")

    voiceid = utils.config.get("azureTTS", "voice_id")
    logging.info(f"Azure TTS voice_id: {voiceid}")

    # Create TTS directly with credentials
    logging.info("Creating MicrosoftTTS instance...")
    try:
        tts = MicrosoftTTS(credentials=(key, location))
        logging.info("MicrosoftTTS instance created successfully")

        logging.info(f"Setting voice to: {voiceid}")
        tts.set_voice(voiceid)
        logging.info("Voice set successfully")

        logging.info("=== init_azure_tts() completed successfully ===")
        return tts
    except Exception as e:
        logging.error(f"Error creating MicrosoftTTS: {e}", exc_info=True)
        return None


def init_google_tts():
    """Initialize unique instance of GoogleTTS based on the changes in voiceid.

    Returns: GoogleTTS
    """
    gcreds = utils.config.get("googleTTS", "creds")
    path = Path(gcreds)
    if not path.exists():
        path = Path("google_creds.enc")
    file_format = path.suffix
    if file_format == ".enc":
        gcreds = load_credentials(path)
        logging.info(f"Google TTS credentials file location: {path}")
    logging.info(f"Google TTS credentials file location: {path}")
    if not isinstance(gcreds, dict) and os.path.isfile(gcreds):
        logging.info(f"Google TTS credentials file: {gcreds}")
    voiceid = utils.config.get("googleTTS", "voice_id")
    client = GoogleClient(credentials=gcreds)
    tts = GoogleTTS(client=client, voice=voiceid)
    return tts


def init_sapi_tts():
    """Initialize unique instance of SAPIEngine based on the changes in voiceid.

    Returns: SAPIEngine
    """
    voiceid = utils.config.get("sapi5TTS", "voice_id")
    client = SAPIClient()
    client._client.setProperty("voice", voiceid)
    client._client.setProperty("rate", utils.config.get("TTS", "rate"))
    client._client.setProperty("volume", utils.config.get("TTS", "volume"))
    return SAPIEngine(client=client)


def init_onnx_tts():
    """Initialize unique instance of SherpaOnnxClient based on the changes in voiceid.

    Returns: SherpaOnnxClient
    """
    try:
        voiceid = utils.config.get("SherpaOnnxTTS", "voice_id")
        logging.info(f"Initializing SherpaOnnx TTS with voice_id: {voiceid}")

        if getattr(sys, "frozen", False):
            home_directory = os.path.expanduser("~")
            onnx_cache_path = os.path.join(
                home_directory,
                "AppData",
                "Roaming",
                "Ace Centre",
                "AACSpeakHelper",
                "models",
            )
        elif __file__:
            app_data_path = os.path.abspath(os.path.dirname(__file__))
            onnx_cache_path = os.path.join(app_data_path, "models")

        logging.info(f"Using model cache path: {onnx_cache_path}")

        if not os.path.isdir(onnx_cache_path):
            os.mkdir(onnx_cache_path)
            logging.info(f"Created model cache directory: {onnx_cache_path}")

        logging.info("Creating SherpaOnnxClient...")
        client = SherpaOnnxClient(model_path=onnx_cache_path, tokens_path=None)
        logging.info("SherpaOnnxClient created successfully")

        logging.info(f"Setting voice to: {voiceid}")
        client.set_voice(voice_id=voiceid)
        logging.info("Voice set successfully")

        return client
    except Exception as e:
        logging.error(f"Error in init_onnx_tts: {e}", exc_info=True)
        raise


def init_googleTrans_tts():
    """Initialize unique instance of GoogleTransTTS based on the changes in voiceid.

    Returns: GoogleTransTTS
    """
    try:
        voiceid = utils.config.get("googleTransTTS", "voice_id")
        logging.info(f"Initializing GoogleTrans TTS with voice: {voiceid}")

        # GoogleTransTTS takes voice_id directly in constructor
        if voiceid:
            tts = GoogleTransTTS(voice_id=voiceid)
            logging.info(f"GoogleTrans TTS initialized with voice: {voiceid}")
        else:
            # Use default voice if none specified
            tts = GoogleTransTTS(voice_id="en")
            logging.info("GoogleTrans TTS initialized with default voice: en")

        return tts
    except Exception as e:
        logging.error(f"Error initializing GoogleTrans TTS: {e}", exc_info=True)
        raise


def init_elevenlabs_tts():
    """Initialize unique instance of ElevenLabsTTS based on the changes in voiceid.

    Returns: ElevenLabsTTS
    """
    import os
    api_key = utils.config.get("ElevenLabsTTS", "api_key", fallback="")
    if not api_key:
        api_key = os.getenv("ELEVENLABS_API_KEY", "")
    voiceid = utils.config.get("ElevenLabsTTS", "voice_id")
    client = ElevenLabsClient(credentials=(api_key,))
    tts = ElevenLabsTTS(client=client, voice=voiceid)
    return tts


def init_playht_tts():
    """Initialize unique instance of PlayHTTTS based on the changes in voiceid.

    Returns: PlayHTTTS
    """
    import os
    api_key = utils.config.get("PlayHTTTS", "api_key", fallback="")
    user_id = utils.config.get("PlayHTTTS", "user_id", fallback="")
    if not api_key:
        api_key = os.getenv("PLAYHT_API_KEY", "")
    if not user_id:
        user_id = os.getenv("PLAYHT_USER_ID", "")
    voiceid = utils.config.get("PlayHTTTS", "voice_id")
    client = PlayHTClient(credentials=(api_key, user_id))
    tts = PlayHTTTS(client=client, voice=voiceid)
    return tts


def init_polly_tts():
    """Initialize unique instance of PollyTTS based on the changes in voiceid.

    Returns: PollyTTS
    """
    import os
    region = utils.config.get("PollyTTS", "region", fallback="")
    aws_key_id = utils.config.get("PollyTTS", "aws_key_id", fallback="")
    aws_access_key = utils.config.get("PollyTTS", "aws_access_key", fallback="")
    if not region:
        region = os.getenv("POLLY_REGION", "us-east-1")
    if not aws_key_id:
        aws_key_id = os.getenv("POLLY_AWS_KEY_ID", "")
    if not aws_access_key:
        aws_access_key = os.getenv("POLLY_AWS_ACCESS_KEY", "")
    voiceid = utils.config.get("PollyTTS", "voice_id")
    client = PollyClient(credentials=(region, aws_key_id, aws_access_key))
    tts = PollyTTS(client=client, voice=voiceid)
    return tts


def init_watson_tts():
    """Initialize unique instance of WatsonTTS based on the changes in voiceid.

    Returns: WatsonTTS
    """
    import os
    api_key = utils.config.get("WatsonTTS", "api_key", fallback="")
    region = utils.config.get("WatsonTTS", "region", fallback="")
    instance_id = utils.config.get("WatsonTTS", "instance_id", fallback="")
    if not api_key:
        api_key = os.getenv("WATSON_API_KEY", "")
    if not region:
        region = os.getenv("WATSON_REGION", "eu-gb")
    if not instance_id:
        instance_id = os.getenv("WATSON_INSTANCE_ID", "")
    voiceid = utils.config.get("WatsonTTS", "voice_id")
    client = WatsonClient(credentials=(api_key, region, instance_id))
    tts = WatsonTTS(client=client, voice=voiceid)
    return tts


def init_openai_tts():
    """Initialize unique instance of OpenAIClient based on the changes in voiceid.

    Returns: OpenAIClient
    """
    import os
    api_key = utils.config.get("OpenAITTS", "api_key", fallback="")
    if not api_key:
        api_key = os.getenv("OPENAI_API_KEY", "")
    voiceid = utils.config.get("OpenAITTS", "voice_id")
    client = OpenAIClient(credentials=(api_key,))
    if voiceid:
        client.set_voice(voiceid)
    return client


def init_witai_tts():
    """Initialize unique instance of WitAiTTS based on the changes in voiceid.

    Returns: WitAiTTS
    """
    import os
    token = utils.config.get("WitAiTTS", "token", fallback="")
    if not token:
        token = os.getenv("WITAI_TOKEN", "")
    voiceid = utils.config.get("WitAiTTS", "voice_id")
    client = WitAiClient(credentials=(token,))
    tts = WitAiTTS(client=client, voice=voiceid)
    return tts


def speak(text="", list_voices=False):
    """Speak function convert text parameter to speech. This function decides which TTS Engine will be used
    base on the config file received.
    Then, it will call the specific function that will create the TTS Engine Instance.

    Args:
        text (str): String to be spoken by specific TTS Engine.
        list_voices (bool): Use to return all available voices only instead of speech function .
    Returns: None
    """


def speak_to_bytes(text="", list_voices=False):
    """Convert text to speech and return audio bytes instead of playing.
    This is used for SAPI integration where audio data needs to be returned.
    Uses tts.speak(text) method from tts-wrapper to get bytes (not tts.speak_streamed()).

    Args:
        text (str): String to be converted to speech.
        list_voices (bool): If True, list available voices instead of speaking.
    Returns: bytes: Audio data as bytes, or None if failed
    """
    global voices
    global ready
    ready = False

    logging.info("=== speak_to_bytes() Debug Information ===")
    logging.info(f"Text to synthesize: {text}")
    logging.info(f"List voices: {list_voices}")

    try:
        # Reuse the same logic as speak() function but call tts.speak() instead of tts.speak_streamed()
        ttsengine = utils.config.get("TTS", "engine")
        logging.info(f"TTS engine from config: {ttsengine}")

        voice_id = utils.config.get(ttsengine, "voice_id")
        if not voice_id:
            voice_id = utils.config.get("TTS", "voice_id")
        logging.info(f"Voice ID: {voice_id}")

        # Get or initialize TTS client (same as speak() function)
        if ttsengine in tts_voiceid and voice_id in tts_voiceid[ttsengine]:
            tts_client = tts_voiceid[ttsengine][voice_id]
        else:
            match ttsengine:
                case "azureTTS":
                    tts_client = init_azure_tts()
                case "googleTTS":
                    tts_client = init_google_tts()
                case "sapi5":
                    tts_client = init_sapi_tts()
                case "SherpaOnnxTTS":
                    tts_client = init_onnx_tts()
                case "googleTransTTS":
                    tts_client = init_googleTrans_tts()
                case "ElevenLabsTTS":
                    tts_client = init_elevenlabs_tts()
                case "PlayHTTTS":
                    tts_client = init_playht_tts()
                case "PollyTTS":
                    tts_client = init_polly_tts()
                case "WatsonTTS":
                    tts_client = init_watson_tts()
                case "OpenAITTS":
                    tts_client = init_openai_tts()
                case "WitAiTTS":
                    tts_client = init_witai_tts()
                case _:
                    logging.error(f"Unsupported TTS engine: {ttsengine}")
                    return None

            if not tts_client:
                logging.error(f"Failed to initialize TTS client for {ttsengine}")
                return None

            # Cache the client
            if ttsengine not in tts_voiceid:
                tts_voiceid[ttsengine] = {voice_id: tts_client}
            else:
                tts_voiceid[ttsengine][voice_id] = tts_client

        # Handle list voices request
        if list_voices:
            try:
                voices = tts_client.get_voices()
                return voices
            except Exception as e:
                logging.error(f"Error getting voices: {e}")
                return None

        # Prepare text with SSML if needed (same as azureSpeak function)
        processed_text = text
        if ttsengine == "azureTTS" and hasattr(utils, 'args') and utils.args.get("style"):
            style = utils.args["style"]
            styledegree = utils.args.get("styledegree")
            if style and style in VALID_STYLES:
                ssml = f'<mstts:express-as style="{style}"'
                if styledegree:
                    ssml += f' styledegree="{styledegree}"'
                ssml += f">{text}</mstts:express-as>"
                processed_text = ssml
                logging.info(f"Using SSML with style: {style}")

        # Try different methods to get audio bytes without playback
        logging.info(f"Attempting to get audio bytes for {ttsengine} without playback")

        # Method 1: Try synth_to_bytes if available (most direct)
        try:
            if hasattr(tts_client, 'synth_to_bytes'):
                logging.info("Trying synth_to_bytes method...")
                audio_bytes = tts_client.synth_to_bytes(processed_text)
                if audio_bytes:
                    logging.info(f"✅ Successfully got {len(audio_bytes)} bytes using synth_to_bytes() method")
                    return audio_bytes
        except Exception as e:
            logging.warning(f"synth_to_bytes() method failed: {e}")

        # Method 2: Try direct speak() method (for Azure TTS this should return bytes)
        try:
            logging.info("Trying speak() method...")
            audio_bytes = tts_client.speak(processed_text)
            if audio_bytes:
                logging.info(f"✅ Successfully got {len(audio_bytes)} bytes using speak() method")
                return audio_bytes
        except Exception as e:
            logging.warning(f"speak() method failed: {e}")

        # Method 3: Use speak_streamed with save_to_file_path (will play audio but also save)
        logging.info("Falling back to speak_streamed with file capture...")

        # Create a temporary file to capture the audio
        import tempfile
        import os

        with tempfile.NamedTemporaryFile(suffix='.wav', delete=False) as temp_file:
            temp_filename = temp_file.name

        try:
            # Use speak_streamed with save_to_file_path
            tts_client.speak_streamed(
                processed_text,
                save_to_file_path=temp_filename,
                audio_format="wav",
                wait_for_completion=True  # Wait for synthesis to complete
            )

            # Read the audio file and return bytes
            if os.path.exists(temp_filename) and os.path.getsize(temp_filename) > 0:
                with open(temp_filename, 'rb') as f:
                    audio_bytes = f.read()
                logging.info(f"✅ Successfully captured {len(audio_bytes)} bytes of audio from file (with playback)")
                return audio_bytes
            else:
                logging.error("No audio file was created or file is empty")
                return None

        except (TypeError, AttributeError) as e:
            logging.error(f"speak_streamed with save_to_file_path failed: {e}")
            return None

        finally:
            # Clean up temporary file
            try:
                if os.path.exists(temp_filename):
                    os.unlink(temp_filename)
            except Exception as e:
                logging.warning(f"Failed to clean up temp file {temp_filename}: {e}")

    except Exception as e:
        logging.error(f"Error in speak_to_bytes: {e}", exc_info=True)
        return None
    finally:
        ready = True


def speak(text="", list_voices=False):
    """Speak function convert text parameter to speech. This function decides which TTS Engine will be used
    base on the config file received.
    Then, it will call the specific function that will create the TTS Engine Instance.

    Args:
        text (str): String to be spoken by specific TTS Engine.
        list_voices (bool): Use to return all available voices only instead of speech function .
    Returns: None
    """
    global voices
    global ready
    ready = False

    logging.info("=== speak() Debug Information ===")
    logging.info(f"Text to speak: {text}")
    logging.info(f"List voices: {list_voices}")

    # Debug: Check if utils.config exists
    if not hasattr(utils, 'config'):
        logging.error("utils.config is not available in speak()!")
        ready = True
        return

    logging.info(f"Available config sections: {utils.config.sections()}")

    try:
        ttsengine = utils.config.get("TTS", "engine")
        logging.info(f"TTS engine from config: {ttsengine}")

        voice_id = utils.config.get(ttsengine, "voice_id")
        logging.info(f"Voice ID from {ttsengine} section: {voice_id}")

        if not voice_id:
            voice_id = utils.config.get("TTS", "voice_id")
            logging.info(f"Fallback voice ID from TTS section: {voice_id}")
    except Exception as e:
        logging.error(f"Error getting TTS engine or voice ID: {e}", exc_info=True)
        ready = True
        return

    try:
        file = utils.check_history(text)
        if file is not None and os.path.isfile(file):
            if list_voices:
                tts_client = tts_voiceid[ttsengine][voice_id]
                voices = tts_client.get_voices()
                return
            else:
                voices = None
            utils.play_audio(file, file=True)
            logging.info(f"Speech synthesized for text [{text}] from cache.")
            ready = True  # Make sure to update ready after cache playback
            return
    except Exception as e:
        logging.error(f"Error checking history or playing audio: {e}")
        ready = True  # In case of error, ensure ready is set back
        return

    logging.info(f"Speech synthesized for text [{text}].")

    try:
        if ttsengine in tts_voiceid and voice_id in tts_voiceid[ttsengine]:
            tts_client = tts_voiceid[ttsengine][voice_id]
        else:
            match ttsengine:
                case "azureTTS":
                    tts_client = init_azure_tts()
                case "googleTTS":
                    tts_client = init_google_tts()
                case "sapi5":
                    tts_client = init_sapi_tts()
                case "SherpaOnnxTTS":
                    tts_client = init_onnx_tts()
                case "googleTransTTS":
                    tts_client = init_googleTrans_tts()
                case "ElevenLabsTTS":
                    tts_client = init_elevenlabs_tts()
                case "PlayHTTTS":
                    tts_client = init_playht_tts()
                case "PollyTTS":
                    tts_client = init_polly_tts()
                case "WatsonTTS":
                    tts_client = init_watson_tts()
                case "OpenAITTS":
                    tts_client = init_openai_tts()
                case "WitAiTTS":
                    tts_client = init_witai_tts()
                case _:
                    logging.error(f"Unsupported TTS engine: {ttsengine}")
                    return
    except Exception as e:
        logging.error(f"Error initializing TTS client: {e}")
        return

    try:
        if ttsengine not in tts_voiceid:
            tts_voiceid[ttsengine] = {voice_id: tts_client}
        else:
            if voice_id not in tts_voiceid[ttsengine]:
                tts_voiceid[ttsengine][voice_id] = tts_client
    except Exception as e:
        logging.error(f"Error storing TTS client: {e}")
        return

    if list_voices:
        try:
            voices = tts_client.get_voices()
            return
        except Exception as e:
            logging.error(f"Error getting voices: {e}")
            return
    else:
        voices = None

    try:
        match ttsengine:
            case "azureTTS":
                if utils.args["style"]:
                    azureSpeak(
                        text,
                        ttsengine,
                        tts_client,
                        utils.args["style"],
                        utils.args["styledegree"],
                    )
                else:
                    azureSpeak(text, ttsengine, tts_client)
            case "googleTTS":
                googleSpeak(text, ttsengine, tts_client)
            case "sapi5":
                sapiSpeak(text, ttsengine, tts_client)
            case "SherpaOnnxTTS":
                onnxSpeak(text, ttsengine, tts_client)
            case "googleTransTTS":
                googleTransSpeak(text, ttsengine, tts_client)
            case "ElevenLabsTTS":
                elevenlabsSpeak(text, ttsengine, tts_client)
            case "PlayHTTTS":
                playhtSpeak(text, ttsengine, tts_client)
            case "PollyTTS":
                pollySpeak(text, ttsengine, tts_client)
            case "WatsonTTS":
                watsonSpeak(text, ttsengine, tts_client)
            case "OpenAITTS":
                openaiSpeak(text, ttsengine, tts_client)
            case "WitAiTTS":
                witaiSpeak(text, ttsengine, tts_client)
            case _:
                logging.error(f"Unsupported TTS engine in speak function: {ttsengine}")
                return
    except Exception as e:
        logging.error(f"Error during TTS processing: {e}")
        ready = True


def onnxSpeak(text: str, engine, tts_client):
    """This function received the input parameters and make necessary modification (if needed). Then, those parameter
    will be pass to ttsWrapperSpeak.

    Args:
        text (str): String to be spoken by the TTS Engine.
        engine (str): Name of the TTS Engine.
        tts_client: Instance of TTS Engine.
    Returns: None
    """

    ttsWrapperSpeak(text, tts_client, engine)


def azureSpeak(
    text: str, engine, tts_client, style: str = None, styledegree: float = None
):
    """This function received the input parameters and make necessary modification (if needed). Then, those parameter
    will be pass to ttsWrapperSpeak.

    Args:
        text (str): String to be spoken by the TTS Engine.
        engine (str): Name of the TTS Engine.
        tts_client: Instance of TTS Engine.
        style (str): Set the SSML style format and wrap the text string.
        styledegree (float): Set the SSML style degree format and wrap the text string.
    Returns: None
    """

    if style:
        # Check if the provided style is in the valid styles array
        if style in VALID_STYLES:
            # Construct SSML with the specified style
            ssml = f'<mstts:express-as style="{style}"'
            if styledegree:
                ssml += f' styledegree="{styledegree}"'
            ssml += f">{text}</mstts:express-as>"
        else:
            # Style is not valid, use default SSML without style
            ssml = text
    else:
        # Use default SSML without style
        ssml = text

    ttsWrapperSpeak(ssml, tts_client, engine)


def googleSpeak(text: str, engine, tts_client):
    """This function received the input parameters and make necessary modification (if needed). Then, those parameter
    will be pass to ttsWrapperSpeak.

    Args:
        text (str): String to be spoken by the TTS Engine.
        engine (str): Name of the TTS Engine.
        tts_client: Instance of TTS Engine.
    Returns: None
    """
    ttsWrapperSpeak(text, tts_client, engine)


def googleTransSpeak(text: str, engine, tts_client):
    """This function received the input parameters and make necessary modification (if needed). Then, those parameter
    will be pass to ttsWrapperSpeak.

    Args:
        text (str): String to be spoken by the TTS Engine.
        engine (str): Name of the TTS Engine.
        tts_client: Instance of TTS Engine.
    Returns: None
    """
    ttsWrapperSpeak(text, tts_client, engine)


def sapiSpeak(text: str, engine, tts_client):
    """This function received the input parameters and make necessary modification (if needed). Then, those parameter
    will be pass to ttsWrapperSpeak.

    Args:
        text (str): String to be spoken by the TTS Engine.
        engine (str): Name of the TTS Engine.
        tts_client: Instance of TTS Engine.
    Returns: None
    """
    ttsWrapperSpeak(text, tts_client, engine)


def elevenlabsSpeak(text: str, engine, tts_client):
    """This function received the input parameters and make necessary modification (if needed). Then, those parameter
    will be pass to ttsWrapperSpeak.

    Args:
        text (str): String to be spoken by the TTS Engine.
        engine (str): Name of the TTS Engine.
        tts_client: Instance of TTS Engine.
    Returns: None
    """
    ttsWrapperSpeak(text, tts_client, engine)


def playhtSpeak(text: str, engine, tts_client):
    """This function received the input parameters and make necessary modification (if needed). Then, those parameter
    will be pass to ttsWrapperSpeak.

    Args:
        text (str): String to be spoken by the TTS Engine.
        engine (str): Name of the TTS Engine.
        tts_client: Instance of TTS Engine.
    Returns: None
    """
    ttsWrapperSpeak(text, tts_client, engine)


def pollySpeak(text: str, engine, tts_client):
    """This function received the input parameters and make necessary modification (if needed). Then, those parameter
    will be pass to ttsWrapperSpeak.

    Args:
        text (str): String to be spoken by the TTS Engine.
        engine (str): Name of the TTS Engine.
        tts_client: Instance of TTS Engine.
    Returns: None
    """
    ttsWrapperSpeak(text, tts_client, engine)


def watsonSpeak(text: str, engine, tts_client):
    """This function received the input parameters and make necessary modification (if needed). Then, those parameter
    will be pass to ttsWrapperSpeak.

    Args:
        text (str): String to be spoken by the TTS Engine.
        engine (str): Name of the TTS Engine.
        tts_client: Instance of TTS Engine.
    Returns: None
    """
    ttsWrapperSpeak(text, tts_client, engine)


def openaiSpeak(text: str, engine, tts_client):
    """This function received the input parameters and make necessary modification (if needed). Then, those parameter
    will be pass to ttsWrapperSpeak.

    Args:
        text (str): String to be spoken by the TTS Engine.
        engine (str): Name of the TTS Engine.
        tts_client: Instance of TTS Engine.
    Returns: None
    """
    ttsWrapperSpeak(text, tts_client, engine)


def witaiSpeak(text: str, engine, tts_client):
    """This function received the input parameters and make necessary modification (if needed). Then, those parameter
    will be pass to ttsWrapperSpeak.

    Args:
        text (str): String to be spoken by the TTS Engine.
        engine (str): Name of the TTS Engine.
        tts_client: Instance of TTS Engine.
    Returns: None
    """
    ttsWrapperSpeak(text, tts_client, engine)


def ttsWrapperSpeak(text: str, tts, engine):
    """This function identifies the TTS Instance and set format of the text and audio format.
    Then, create a Thread that synthesize text to audio.

    Args:
        text (str): String to be spoken by the TTS Engine.
        tts: Instance of TTS Engine.
        engine (str): Name of the TTS Engine.
    Returns: None
    """
    fmt = "wav"
    match tts:
        case SherpaOnnxClient():
            pass
        case GoogleTransTTS():
            fmt = "mp3"
        case SAPIEngine():
            pass
        case AbstractTTS():
            tts.ssml.clear_ssml()
            text = tts.ssml.add(text)
    try:
        playText = Thread(target=playSpeech, args=(text, engine, fmt, tts))
        playText.start()
    except Exception as e:
        print(e)


def playSpeech(text, engine, file_format, tts):
    """This function is run by a Thread which synthesize text to audio.
    While audio is streaming, the audio is also saving in parallel.

    Args:
        text (str): String to be spoken by the TTS Engine.
        engine (str): Name of the TTS Engine.
        file_format (str): Audio Format.
        tts: Instance of TTS Engine.
    Returns: None
    """
    global ready
    ready = False  # Start in an unready state

    start = time.perf_counter()
    try:
        save_audio_file = utils.config.getboolean("TTS", "save_audio_file")
        if save_audio_file:
            utils.save_audio(text=text, engine=engine, file_format=file_format, tts=tts)
            logging.info(f"Speech synthesized for text [{text}] saved in cache.")
        else:
            tts.speak_streamed(text)
    except Exception as e:
        logging.error(f"Error during TTS processing: {e}")
        return
    finally:
        stop = time.perf_counter() - start
        logging.info(f"Speech synthesis runtime is {stop:0.5f} seconds.")
        ready = True  # Ensure ready is set to True even after an exception
