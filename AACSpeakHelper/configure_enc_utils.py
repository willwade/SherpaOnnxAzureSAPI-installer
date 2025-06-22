import io
import json
import logging
import os
import pickle
from pathlib import Path
import configparser
from cryptography.fernet import Fernet
import sys
import argparse
import base64

# Ensure CONFIG_ENCRYPTION_KEY is set
if "CONFIG_ENCRYPTION_KEY" not in os.environ:
    os.environ["CONFIG_ENCRYPTION_KEY"] = "YOUR_ENCRYPTION_KEY"


# use it like this
#  Encode the JSON file and save it to a variable
# base64 google_creds.json > google_creds_base64.txt


def create_google_creds_file(filename):
    # Fetch the Base64 encoded JSON string from the environment variable
    google_creds_base64 = os.getenv("GOOGLE_CREDS_JSON")
    if not google_creds_base64:
        raise ValueError("GOOGLE_CREDS_JSON environment variable is not set")

    try:
        # Decode the Base64 string to get the JSON string
        google_creds_json = base64.b64decode(google_creds_base64).decode("utf-8")
        # Parse the JSON string to ensure it's formatted correctly
        google_creds_dict = json.loads(google_creds_json)

        # Write the JSON dictionary back to a file if needed
        with open(filename, "w") as f:
            json.dump(google_creds_dict, f, indent=4)

    except (json.JSONDecodeError, base64.binascii.Error) as e:
        raise ValueError(f"Failed to decode and parse GOOGLE_CREDS_BASE64: {e}")


def load_encryption_key():
    """Loads the encryption key from the environment variable."""
    encryption_key = os.getenv("CONFIG_ENCRYPTION_KEY")
    if not encryption_key:
        logging.error("CONFIG_ENCRYPTION_KEY environment variable is not set.")
        raise EnvironmentError("CONFIG_ENCRYPTION_KEY environment variable is not set.")
    return encryption_key.encode()


def prepare_config_enc(output_path="config.enc"):
    """Prepares and encrypts the configuration from environment variables into config.enc."""
    required_keys = [
        "MICROSOFT_TOKEN",
        "MICROSOFT_REGION",
        "MICROSOFT_TOKEN_TRANS",
        "GOOGLE_CREDS_JSON",
    ]
    config = {}

    # Collect configuration from environment variables
    for key in required_keys:
        value = os.getenv(key)
        if not value:
            logging.error(f"Missing required environment variable: {key}")
            raise EnvironmentError(f"Missing required environment variable: {key}")
        config[key] = value

    # Convert the config dictionary to JSON bytes
    config_json = json.dumps(config).encode()

    # Load the encryption key and encrypt the configuration
    encryption_key = load_encryption_key()
    fernet = Fernet(encryption_key)
    encrypted_config = fernet.encrypt(config_json)

    # Save the encrypted configuration to the file
    with open(output_path, "wb") as config_file:
        config_file.write(encrypted_config)
    logging.info(f"Encrypted configuration saved to {output_path}.")


def generate_key():
    """Generate a new Fernet key and save it"""
    key = Fernet.generate_key()
    key_file = os.path.join(get_config_dir(), ".key")
    with open(key_file, "wb") as f:
        f.write(key)
    return key


def get_config_dir():
    """Get the configuration directory path"""
    if getattr(sys, "frozen", False):
        config_dir = os.path.join(
            os.path.expanduser("~"),
            "AppData",
            "Roaming",
            "Ace Centre",
            "AACSpeakHelper"
        )
    else:
        config_dir = os.path.dirname(os.path.abspath(__file__))
    
    os.makedirs(config_dir, exist_ok=True)
    return config_dir


def create_default_config():
    """Create a default configuration"""
    import uuid
    config = configparser.ConfigParser()

    # App section
    config["App"] = {
        "collectstats": "True",
        "uuid": str(uuid.uuid4())
    }

    # Translate section
    config["translate"] = {
        "no_translate": "False",
        "start_lang": "en",
        "end_lang": "en",
        "replace_pb": "True",
        "provider": "GoogleTranslator",
        "microsoft_translator_secret_key": "",
        "papago_translator_client_id": "",
        "papago_translator_secret_key": "",
        "my_memory_translator_secret_key": "",
        "email": "",
        "libre_translator_secret_key": "",
        "url": "",
        "deep_l_translator_secret_key": "",
        "deepl_pro": "false",
        "region": "",
        "yandex_translator_secret_key": "",
        "qcri_translator_secret_key": "",
        "baidu_translator_appid": "",
        "baidu_translator_secret_key": ""
    }

    # TTS section
    config["TTS"] = {
        "engine": "SherpaOnnxTTS",
        "bypass_tts": "False",
        "save_audio_file": "True",
        "rate": "0",
        "volume": "100",
        "voice_id": "eng"
    }

    # Azure TTS section
    config["azureTTS"] = {
        "key": "",
        "location": "",
        "voice_id": "en-US-JennyNeural"
    }

    # Google TTS section
    config["googleTTS"] = {
        "creds": "",
        "voice_id": "en-US-Wavenet-C"
    }

    # Google Trans TTS section
    config["googleTransTTS"] = {
        "voice_id": ""
    }

    # SAPI5 TTS section
    config["sapi5TTS"] = {
        "voice_id": "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\TTS_MS_EN-US_DAVID_11.0"
    }

    # Sherpa ONNX TTS section
    config["SherpaOnnxTTS"] = {
        "voice_id": "eng"
    }

    # App Cache section
    config["appCache"] = {
        "threshold": "7"
    }

    return config


def load_config(custom_config_path=None):
    """Load configuration from file, supporting both plain text and encrypted formats.

    Args:
        custom_config_path (str, optional): Path to a custom configuration file.
                                          If not provided, uses default location.

    Returns:
        configparser.ConfigParser: Loaded configuration object.
    """
    try:
        # Determine config file path
        if custom_config_path and os.path.exists(custom_config_path):
            config_file = custom_config_path
            logging.info(f"Using custom config path: {config_file}")
        else:
            config_file = os.path.join(get_config_dir(), "settings.cfg")
            logging.info(f"Using default config path: {config_file}")

        # If no config exists, create default
        if not os.path.exists(config_file):
            logging.info("No configuration found, creating default")
            config = create_default_config()
            return config

        # First, try to load as plain text configuration
        try:
            config = configparser.ConfigParser()
            config.read(config_file)

            # Verify it's a valid config by checking if it has sections
            if config.sections():
                logging.info(f"Successfully loaded plain text configuration from {config_file}")
                return config
            else:
                logging.info("Config file appears to be empty or invalid, trying encrypted format")
        except Exception as e:
            logging.info(f"Failed to load as plain text config: {e}, trying encrypted format")

        # If plain text loading failed, try encrypted format
        try:
            key_file = os.path.join(get_config_dir(), ".key")

            # If no key exists, generate one
            if not os.path.exists(key_file):
                logging.info("No encryption key found, generating new key")
                encryption_key = generate_key()
            else:
                with open(key_file, "rb") as f:
                    encryption_key = f.read()

            fernet = Fernet(encryption_key)

            # Try to load existing encrypted config
            with open(config_file, "rb") as f:
                encrypted_data = f.read()
            decrypted_data = fernet.decrypt(encrypted_data)
            config = configparser.ConfigParser()
            config.read_string(decrypted_data.decode())
            logging.info(f"Successfully loaded encrypted configuration from {config_file}")
            return config

        except Exception as e:
            logging.error(f"Failed to load encrypted configuration: {e}")
            logging.info("Creating new default configuration")
            return create_default_config()

    except Exception as e:
        logging.error(f"Failed to load configuration: {e}")
        raise


def load_credentials(fp: str) -> object:
    encryption_key = load_encryption_key()
    fernet = Fernet(encryption_key)
    with open(fp, "rb") as f:
        return pickle.loads(fernet.decrypt(f.read()))


def save_credentials(obj: object, fp: str):
    encryption_key = load_encryption_key()
    fernet = Fernet(encryption_key)
    with open(fp, "wb") as f:
        f.write(fernet.encrypt(pickle.dumps(obj)))


# Example usage
# prepare_config_enc()  # Run this to create config.enc
# config = load_config()  # Load config when needed


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="AACSpeakHelper Encryption Utility")
    parser.add_argument(
        "-i", "--input", help="Path to a defined JSON file", required=False
    )
    parser.add_argument(
        "-use-env",
        action="store_true",
        help="Create JSON from environment variables instead of input file",
    )
    args = vars(parser.parse_args())

    # Check if -use-env is provided
    if args["use_env"]:
        # Set a default filename for the JSON file to be created from the environment variable
        filename = Path("google_creds.json")
        create_google_creds_file(filename)
        prepare_config_enc()
    else:
        # Handle the input flag normally
        if not args["input"]:
            print("Either --input or --use-env must be specified.")
            sys.exit(1)

        filename = Path(args["input"])

    file_path = filename.resolve().parent
    with io.open(filename, "r", encoding="utf-8") as json_file:
        json_dict = json.load(json_file)
        new_file = filename.with_suffix(".enc")
        save_credentials(json_dict, os.path.join(file_path, new_file))
        files = os.listdir(".")
        print(files)
