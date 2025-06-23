#!/usr/bin/env python3
"""
test_pipe.py - Simple standalone test client for AACSpeakHelper server

This script tests the Windows named pipe communication with the AACSpeakHelper server.
It sends TTS requests and receives audio data back through the pipe.

Usage:
    python test_pipe.py "Hello world"
    python test_pipe.py "Hello world" --engine azure --voice en-GB-LibbyNeural
    python test_pipe.py "Hello world" --engine sherpaonnx --voice en_GB-jenny_dioco-medium
    python test_pipe.py --help

Author: Ace Centre
"""

import argparse
import json
import struct
import sys
import time
import win32file
import win32pipe
import win32api


def connect_to_pipe(pipe_name, timeout=5000):
    """Connect to the named pipe with timeout"""
    try:
        # Wait for pipe to be available
        win32pipe.WaitNamedPipe(pipe_name, timeout)
        
        # Open the pipe
        pipe_handle = win32file.CreateFile(
            pipe_name,
            win32file.GENERIC_READ | win32file.GENERIC_WRITE,
            0,
            None,
            win32file.OPEN_EXISTING,
            0,
            None
        )
        
        return pipe_handle
    except Exception as e:
        print(f"❌ Failed to connect to pipe: {e}")
        return None


def send_tts_request(pipe_handle, text, engine="sherpaonnx", voice=None):
    """Send TTS request through the pipe"""
    try:
        # Prepare the request data
        request_data = {
            "text": text,
            "args": {
                "engine": engine,
                "voice": voice,
                "return_audio_bytes": True
            }
        }
        
        # Convert to JSON and encode
        json_data = json.dumps(request_data)
        message_bytes = json_data.encode('utf-8')
        
        print(f"📤 Sending request: {len(message_bytes)} bytes")
        print(f"   Text: {text}")
        print(f"   Engine: {engine}")
        if voice:
            print(f"   Voice: {voice}")
        
        # Send the message
        win32file.WriteFile(pipe_handle, message_bytes)
        
        return True
    except Exception as e:
        print(f"❌ Failed to send request: {e}")
        return False


def receive_audio_response(pipe_handle, timeout=30):
    """Receive audio response from the pipe"""
    try:
        print(f"📥 Waiting for response (timeout: {timeout}s)...")
        
        # Read the length prefix (4 bytes)
        result, length_data = win32file.ReadFile(pipe_handle, 4)
        if result != 0:
            print(f"❌ Failed to read length prefix: {result}")
            return None
        
        # Unpack the length
        audio_length = struct.unpack('<I', length_data)[0]
        print(f"📊 Expected audio length: {audio_length} bytes")
        
        if audio_length == 0:
            print("⚠️ Received empty audio response")
            return None
        
        # Read the audio data
        audio_data = b""
        bytes_remaining = audio_length
        
        while bytes_remaining > 0:
            chunk_size = min(bytes_remaining, 64 * 1024)  # Read in 64KB chunks
            result, chunk = win32file.ReadFile(pipe_handle, chunk_size)
            
            if result != 0:
                print(f"❌ Failed to read audio chunk: {result}")
                return None
            
            audio_data += chunk
            bytes_remaining -= len(chunk)
            
            if len(chunk) == 0:
                break
        
        print(f"✅ Received {len(audio_data)} bytes of audio data")
        return audio_data
        
    except Exception as e:
        print(f"❌ Failed to receive response: {e}")
        return None


def save_audio_file(audio_data, filename="test_output.wav"):
    """Save audio data to a file"""
    try:
        with open(filename, 'wb') as f:
            f.write(audio_data)
        print(f"💾 Audio saved to: {filename}")
        return True
    except Exception as e:
        print(f"❌ Failed to save audio: {e}")
        return False


def play_audio_file(filename):
    """Play audio file using multiple methods"""
    import subprocess
    import os

    # Method 1: Try Windows Media Player
    try:
        subprocess.run(['wmplayer', filename], timeout=10, capture_output=True)
        print(f"🔊 Played audio with Windows Media Player: {filename}")
        return True
    except:
        pass

    # Method 2: Try PowerShell SoundPlayer
    try:
        cmd = f'(New-Object Media.SoundPlayer "{os.path.abspath(filename)}").PlaySync()'
        result = subprocess.run(['powershell', '-c', cmd],
                              capture_output=True, timeout=10, text=True)
        if result.returncode == 0:
            print(f"🔊 Played audio with PowerShell: {filename}")
            return True
        else:
            print(f"PowerShell error: {result.stderr}")
    except Exception as e:
        print(f"PowerShell failed: {e}")

    # Method 3: Try start command (opens with default player)
    try:
        subprocess.run(['start', '', filename], shell=True, timeout=5)
        print(f"🔊 Opened audio with default player: {filename}")
        return True
    except Exception as e:
        print(f"Start command failed: {e}")

    # Method 4: Try pygame (if available)
    try:
        import pygame
        pygame.mixer.init()
        pygame.mixer.music.load(filename)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            pygame.time.wait(100)
        print(f"🔊 Played audio with pygame: {filename}")
        return True
    except ImportError:
        print("pygame not available")
    except Exception as e:
        print(f"pygame failed: {e}")

    print(f"⚠️ Could not play audio with any method. File saved as: {filename}")
    print(f"💡 Try manually opening: {os.path.abspath(filename)}")
    return False

def test_pipe_communication(text, engine="sherpaonnx", voice=None, save_file=True, play_audio=False):
    """Test the complete pipe communication"""
    pipe_name = r"\\.\pipe\AACSpeakHelper"

    print("🔧 AACSpeakHelper Pipe Test Client")
    print("=" * 50)

    # Connect to pipe
    print("🔌 Connecting to pipe...")
    pipe_handle = connect_to_pipe(pipe_name)
    if not pipe_handle:
        print("❌ Could not connect to pipe. Is the server running?")
        return False

    try:
        # Send request
        if not send_tts_request(pipe_handle, text, engine, voice):
            return False

        # Receive response
        audio_data = receive_audio_response(pipe_handle)
        if not audio_data:
            return False

        # Save audio file
        filename = None
        if save_file:
            timestamp = int(time.time())
            filename = f"test_output_{timestamp}.wav"
            save_audio_file(audio_data, filename)

        # Play audio if requested
        if play_audio and filename:
            play_audio_file(filename)

        print("✅ Test completed successfully!")
        return True

    finally:
        # Close the pipe
        win32file.CloseHandle(pipe_handle)
        print("🔌 Pipe connection closed")


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description="Test client for AACSpeakHelper pipe server",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python test_pipe.py "Hello world"
  python test_pipe.py "Hello world" --engine azure
  python test_pipe.py "Hello world" --engine azure --voice en-GB-LibbyNeural
  python test_pipe.py "Hello world" --engine sherpaonnx --voice en_GB-jenny_dioco-medium
  python test_pipe.py "Testing Azure TTS" --engine azure --no-save
        """
    )
    
    parser.add_argument("text", help="Text to synthesize")
    parser.add_argument("--engine", choices=["azure", "sherpaonnx"], default="sherpaonnx",
                        help="TTS engine to use (default: sherpaonnx)")
    parser.add_argument("--voice", help="Voice ID to use (optional)")
    parser.add_argument("--no-save", action="store_true", help="Don't save audio to file")
    parser.add_argument("--play", action="store_true", help="Play the generated audio")
    
    args = parser.parse_args()
    
    # Run the test
    success = test_pipe_communication(
        text=args.text,
        engine=args.engine,
        voice=args.voice,
        save_file=not args.no_save,
        play_audio=args.play
    )
    
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
