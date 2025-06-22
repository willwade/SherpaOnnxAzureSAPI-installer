# prepare_config_enc.py
import os
from configure_enc_utils import prepare_config_enc
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


def main():
    """
    Main function to prepare the encrypted configuration.
    """
    # Optionally, allow specifying the output path via an environment variable
    output_path = os.getenv("CONFIG_OUTPUT_PATH", "config.enc")
    prepare_config_enc(output_path)


if __name__ == "__main__":
    main()
