import os
import sys
from cryptography.fernet import Fernet
import logging


def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s — %(levelname)s — %(message)s",
    )
    return logging.getLogger()


logger = setup_logging()


def generate_key():
    """
    Generates a Fernet encryption key.
    """
    key = Fernet.generate_key()
    return key


def save_key(key, file_path):
    """
    Saves the encryption key to a file with restricted permissions.
    """
    try:
        with open(file_path, "wb") as key_file:
            key_file.write(key)
        logger.info(f"Encryption key saved to '{file_path}'.")
    except Exception as e:
        logger.error(f"Failed to save encryption key: {e}")
        sys.exit(1)


def main():
    # Define the key file path
    key_file = "encryption_key.key"

    # Check if the key file already exists to prevent overwriting
    if os.path.exists(key_file):
        logger.error(
            f"Encryption key file '{key_file}' already exists. Generation aborted to prevent overwriting."
        )
        sys.exit(1)

    # Generate the encryption key
    logger.info("Generating encryption key...")
    key = generate_key()

    # Save the key to the file
    save_key(key, key_file)

    # Optionally, set file permissions (Windows and Unix/Linux)
    try:
        if os.name == "nt":
            # On Windows, use the built-in `icacls` tool to restrict permissions
            os.system(f'icacls "{key_file}" /inheritance:r /grant:r "%USERNAME%:F"')
            logger.info(
                f"File permissions set to allow only the current user to access '{key_file}'."
            )
        else:
            # On Unix/Linux, set the file permission to read/write for the user only
            os.chmod(key_file, 0o600)
            logger.info(
                f"File permissions set to read/write for the user only on '{key_file}'."
            )
    except Exception as e:
        logger.error(f"Failed to set file permissions: {e}")
        sys.exit(1)

    # Display the key to the user (optional and should be handled securely)
    logger.info("Encryption key generation complete.")
    print(
        f"Your encryption key has been saved to '{key_file}'. Please keep it secure and do not share it."
    )


if __name__ == "__main__":
    main()
