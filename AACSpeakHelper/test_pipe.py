#!/usr/bin/env python3
"""
Simple test script for AACSpeakHelper pipe service.
Replaces the more complex client.py with essential testing functionality.
"""

import sys
import os
import json
import logging
import win32file
import pywintypes
import time
import argparse

def setup_logging():
    """Setup basic logging"""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

def send_to_pipe(text, engine="SherpaOnnxTTS", voice="en_GB-jenny_dioco-medium", return_audio=False):
    """
    Send a simple test message to the AACSpeakHelper pipe.
    
    Args:
        text (str): Text to synthesize
        engine (str): TTS engine to use
        voice (str): Voice ID to use
        return_audio (bool): Whether to request audio bytes back
    
    Returns:
        bool: True if successful, False otherwise
    """
    pipe_name = r"\\.\pipe\AACSpeakHelper"
    
    # Create test message
    data = {
        "clipboard_text": text,
        "args": {
            "engine": engine,
            "voice": voice,
            "rate": 0,
            "volume": 100,
            "listvoices": False,
            "return_audio_bytes": return_audio,
            "verbose": True
        },
        "config": {
            "TTS": {
                "engine": engine,
                "rate": "0",
                "volume": "100"
            },
            engine: {
                "voice_id": voice
            }
        }
    }
    
    try:
        # Connect to pipe
        handle = win32file.CreateFile(
            pipe_name,
            win32file.GENERIC_READ | win32file.GENERIC_WRITE,
            0,
            None,
            win32file.OPEN_EXISTING,
            0,
            None,
        )
        
        # Send message
        message = json.dumps(data).encode()
        win32file.WriteFile(handle, message)
        logging.info(f"Sent message: {text}")
        
        # Read response if expecting audio
        if return_audio:
            try:
                result, response = win32file.ReadFile(handle, 128 * 1024)
                if result == 0 and len(response) > 0:
                    logging.info(f"Received {len(response)} bytes of audio data")
                else:
                    logging.warning("No audio data received")
            except Exception as e:
                logging.error(f"Error reading response: {e}")
        
        win32file.CloseHandle(handle)
        return True
        
    except pywintypes.error as e:
        logging.error(f"Pipe communication error: {e}")
        return False
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        return False

def main():
    """Main function"""
    parser = argparse.ArgumentParser(description="Test AACSpeakHelper pipe service")
    parser.add_argument("-t", "--text", default="Hello, this is a test.", 
                       help="Text to synthesize")
    parser.add_argument("-e", "--engine", default="SherpaOnnxTTS",
                       help="TTS engine to use")
    parser.add_argument("-v", "--voice", default="en_GB-jenny_dioco-medium",
                       help="Voice ID to use")
    parser.add_argument("-b", "--bytes", action="store_true",
                       help="Request audio bytes back")
    
    args = parser.parse_args()
    
    setup_logging()
    
    logging.info("Testing AACSpeakHelper pipe service...")
    logging.info(f"Text: {args.text}")
    logging.info(f"Engine: {args.engine}")
    logging.info(f"Voice: {args.voice}")
    
    success = send_to_pipe(args.text, args.engine, args.voice, args.bytes)
    
    if success:
        logging.info("✅ Test completed successfully")
        sys.exit(0)
    else:
        logging.error("❌ Test failed")
        sys.exit(1)

if __name__ == "__main__":
    main()
