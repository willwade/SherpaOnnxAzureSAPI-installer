import os
import json
from cryptography.fernet import Fernet
import sys
import logging
from pathlib import Path


def get_google_creds_path():
    """
    Determines the path for google_creds.json based on whether the application is frozen.

    Returns:
        str: The path to google_creds.json.
    """
    if getattr(sys, "frozen", False):
        # Running as a bundled executable
        if os.name == "nt":
            # Windows
            app_data = (
                Path.home() / "AppData" / "Roaming" / "Ace Centre" / "AACSpeakHelper"
            )
        else:
            # macOS/Linux
            app_data = Path.home() / ".config" / "AceCentre" / "AACSpeakHelper"
    else:
        # Running as a script (development)
        app_data = Path(__file__).parent  # Root directory of the repo

    return str(app_data / "google_creds.json")


def encrypt_config(output_path, encryption_key):
    """
    Encrypt the configuration and save to the specified output path.

    Args:
        output_path (str): Path to save the encrypted configuration.
        encryption_key (str): Encryption key for Fernet.
    """
    google_creds_json = os.getenv("GOOGLE_CREDS_JSON")
    if not google_creds_json:
        raise ValueError("GOOGLE_CREDS_JSON environment variable is not set")

    config = {
        "MICROSOFT_TOKEN": os.getenv("MICROSOFT_TOKEN"),
        "MICROSOFT_REGION": os.getenv("MICROSOFT_REGION"),
        "MICROSOFT_TOKEN_TRANS": os.getenv("MICROSOFT_TOKEN_TRANS"),
        "GOOGLE_CREDS_JSON": google_creds_json,
        "GOOGLE_CREDS_PATH": get_google_creds_path(),
    }

    # Validate all config values are present
    missing = [k for k, v in config.items() if not v]
    if missing:
        raise ValueError(f"Missing environment variables: {', '.join(missing)}")

    # Serialize to JSON
    config_json = json.dumps(config).encode()

    # Encrypt
    fernet = Fernet(encryption_key.encode())
    encrypted_config = fernet.encrypt(config_json)

    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Write encrypted config
    with open(output_path, "wb") as f:
        f.write(encrypted_config)

    logging.info(f"Encrypted configuration saved to {output_path}")


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    try:
        encryption_key = os.getenv("CONFIG_ENCRYPTION_KEY")
        if not encryption_key:
            raise EnvironmentError(
                "CONFIG_ENCRYPTION_KEY environment variable is not set."
            )

        output_path = os.path.join(os.path.dirname(__file__), "config.enc")
        encrypt_config(output_path, encryption_key)
    except Exception as e:
        logging.error(f"Encryption failed: {e}")
        sys.exit(1)
