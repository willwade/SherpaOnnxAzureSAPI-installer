#!/usr/bin/env python3
"""
PyAudio Sound Generator Test
This script tests PyAudio functionality with intensive logging.
It will generate a simple tone and try to play it through the audio system.
"""

import sys
import os
import traceback
import logging
from datetime import datetime

# Set up intensive logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pyaudio_test.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)

def log_system_info():
    """Log detailed system information"""
    logger.info("=== SYSTEM INFORMATION ===")
    logger.info(f"Python version: {sys.version}")
    logger.info(f"Platform: {sys.platform}")
    logger.info(f"Executable: {sys.executable}")
    
    if getattr(sys, 'frozen', False):
        logger.info("Running from FROZEN executable")
        logger.info(f"Frozen executable path: {sys.executable}")
        if hasattr(sys, '_MEIPASS'):
            logger.info(f"PyInstaller temp dir: {sys._MEIPASS}")
    else:
        logger.info("Running from Python interpreter")
    
    logger.info(f"Current working directory: {os.getcwd()}")
    logger.info(f"Script directory: {os.path.dirname(os.path.abspath(__file__))}")

def test_pyaudio_import():
    """Test PyAudio import with detailed logging"""
    logger.info("=== TESTING PYAUDIO IMPORT ===")
    
    try:
        logger.debug("Attempting to import pyaudio...")
        import pyaudio
        logger.info("[OK] PyAudio imported successfully")
        logger.info(f"PyAudio version: {pyaudio.__version__}")
        logger.info(f"PyAudio location: {pyaudio.__file__}")

        # Check for the _portaudio module
        try:
            logger.debug("Checking _portaudio module...")
            import pyaudio._portaudio
            logger.info("[OK] _portaudio module accessible")
        except Exception as e:
            logger.error(f"[ERROR] _portaudio module issue: {e}")
            logger.debug(traceback.format_exc())

        return pyaudio

    except ImportError as e:
        logger.error(f"[ERROR] Failed to import PyAudio: {e}")
        logger.debug(traceback.format_exc())
        return None
    except Exception as e:
        logger.error(f"[ERROR] Unexpected error importing PyAudio: {e}")
        logger.debug(traceback.format_exc())
        return None

def test_pyaudio_initialization(pyaudio):
    """Test PyAudio initialization with detailed logging"""
    logger.info("=== TESTING PYAUDIO INITIALIZATION ===")
    
    try:
        logger.debug("Creating PyAudio instance...")
        p = pyaudio.PyAudio()
        logger.info("[OK] PyAudio instance created successfully")
        
        # Get device count
        logger.debug("Getting device count...")
        device_count = p.get_device_count()
        logger.info(f"[OK] Found {device_count} audio devices")
        
        # List all devices with detailed info
        logger.info("=== AUDIO DEVICE DETAILS ===")
        output_devices = []
        
        for i in range(device_count):
            try:
                device_info = p.get_device_info_by_index(i)
                logger.info(f"Device {i}: {device_info['name']}")
                logger.debug(f"  Max input channels: {device_info['maxInputChannels']}")
                logger.debug(f"  Max output channels: {device_info['maxOutputChannels']}")
                logger.debug(f"  Default sample rate: {device_info['defaultSampleRate']}")
                logger.debug(f"  Host API: {device_info['hostApi']}")
                
                if device_info['maxOutputChannels'] > 0:
                    output_devices.append(i)
                    logger.info("  [OK] Output device available")
                
            except Exception as e:
                logger.error(f"  [ERROR] Error getting device {i} info: {e}")
        
        logger.info(f"Found {len(output_devices)} output devices: {output_devices}")
        
        # Get default devices
        try:
            logger.debug("Getting default input device...")
            default_input = p.get_default_input_device_info()
            logger.info(f"Default input device: {default_input['name']}")
        except Exception as e:
            logger.warning(f"No default input device: {e}")
        
        try:
            logger.debug("Getting default output device...")
            default_output = p.get_default_output_device_info()
            logger.info(f"Default output device: {default_output['name']}")
            logger.info(f"Default output device index: {default_output['index']}")
        except Exception as e:
            logger.error(f"[ERROR] No default output device: {e}")
            p.terminate()
            return None, None
        
        return p, default_output
        
    except Exception as e:
        logger.error(f"[ERROR] PyAudio initialization failed: {e}")
        logger.debug(traceback.format_exc())
        return None, None

def generate_tone(frequency=440, duration=1.0, sample_rate=44100, amplitude=0.3):
    """Generate a sine wave tone"""
    logger.debug(f"Generating tone: {frequency}Hz, {duration}s, {sample_rate}Hz sample rate")
    
    try:
        import numpy as np
        import math
        
        # Generate time array
        t = np.linspace(0, duration, int(sample_rate * duration), False)
        
        # Generate sine wave
        tone = amplitude * np.sin(2 * np.pi * frequency * t)
        
        # Convert to 16-bit integers
        tone_int16 = (tone * 32767).astype(np.int16)
        
        logger.info(f"[OK] Generated {len(tone_int16)} samples")
        return tone_int16.tobytes()
        
    except ImportError:
        logger.warning("NumPy not available, generating tone manually...")
        
        # Manual tone generation without NumPy
        import math
        samples = []
        num_samples = int(sample_rate * duration)
        
        for i in range(num_samples):
            t = i / sample_rate
            sample = amplitude * math.sin(2 * math.pi * frequency * t)
            sample_int16 = int(sample * 32767)
            # Convert to bytes (little-endian 16-bit)
            samples.append(sample_int16.to_bytes(2, byteorder='little', signed=True))
        
        tone_bytes = b''.join(samples)
        logger.info(f"[OK] Generated {len(tone_bytes)} bytes manually")
        return tone_bytes
        
    except Exception as e:
        logger.error(f"[ERROR] Tone generation failed: {e}")
        logger.debug(traceback.format_exc())
        return None

def test_audio_stream(pyaudio, p, device_info):
    """Test audio stream creation and playback"""
    logger.info("=== TESTING AUDIO STREAM ===")
    
    try:
        # Generate test tone
        logger.debug("Generating test tone...")
        tone_data = generate_tone(frequency=440, duration=2.0)
        if tone_data is None:
            return False
        
        # Try different audio configurations
        configs = [
            {"format": pyaudio.paInt16, "channels": 1, "rate": 44100},
            {"format": pyaudio.paInt16, "channels": 2, "rate": 44100},
            {"format": pyaudio.paInt16, "channels": 1, "rate": 22050},
        ]
        
        for i, config in enumerate(configs):
            logger.info(f"--- Testing configuration {i+1}: {config} ---")
            
            try:
                logger.debug("Creating audio stream...")
                stream = p.open(
                    format=config["format"],
                    channels=config["channels"],
                    rate=config["rate"],
                    output=True,
                    output_device_index=device_info['index'],
                    frames_per_buffer=1024
                )
                
                logger.info("[OK] Audio stream created successfully")
                logger.debug(f"Stream info: active={stream.is_active()}, stopped={stream.is_stopped()}")
                
                # Prepare audio data for the channel configuration
                if config["channels"] == 2 and len(tone_data) > 0:
                    # Duplicate mono to stereo
                    import struct
                    mono_samples = struct.unpack('<' + 'h' * (len(tone_data) // 2), tone_data)
                    stereo_samples = []
                    for sample in mono_samples:
                        stereo_samples.extend([sample, sample])  # Left and right channels
                    audio_data = struct.pack('<' + 'h' * len(stereo_samples), *stereo_samples)
                else:
                    audio_data = tone_data
                
                logger.info("[AUDIO] Playing test tone...")
                logger.debug(f"Audio data length: {len(audio_data)} bytes")
                
                # Play the tone in chunks
                chunk_size = 1024 * config["channels"] * 2  # 1024 frames * channels * 2 bytes per sample
                for chunk_start in range(0, len(audio_data), chunk_size):
                    chunk = audio_data[chunk_start:chunk_start + chunk_size]
                    if len(chunk) > 0:
                        stream.write(chunk)
                        logger.debug(f"Wrote chunk: {len(chunk)} bytes")
                
                logger.info("[OK] Audio playback completed")
                
                # Clean up
                logger.debug("Stopping stream...")
                stream.stop_stream()
                logger.debug("Closing stream...")
                stream.close()
                logger.info("[OK] Stream closed successfully")
                
                return True
                
            except Exception as e:
                logger.error(f"[ERROR] Stream test {i+1} failed: {e}")
                logger.debug(traceback.format_exc())
                
                # Try to clean up
                try:
                    if 'stream' in locals():
                        stream.close()
                except:
                    pass
        
        logger.error("[ERROR] All stream configurations failed")
        return False
        
    except Exception as e:
        logger.error(f"[ERROR] Audio stream test failed: {e}")
        logger.debug(traceback.format_exc())
        return False

def main():
    """Main test function"""
    logger.info("[AUDIO] PyAudio Sound Generator Test Starting")
    logger.info("=" * 60)
    
    start_time = datetime.now()
    logger.info(f"Test started at: {start_time}")
    
    # Log system information
    log_system_info()
    
    # Test PyAudio import
    pyaudio = test_pyaudio_import()
    if pyaudio is None:
        logger.error("[FAIL] Cannot proceed - PyAudio import failed")
        return 1
    
    # Test PyAudio initialization
    p, device_info = test_pyaudio_initialization(pyaudio)
    if p is None:
        logger.error("[FAIL] Cannot proceed - PyAudio initialization failed")
        return 1
    
    # Test audio stream
    try:
        success = test_audio_stream(pyaudio, p, device_info)
        
        if success:
            logger.info("[SUCCESS] ALL TESTS PASSED - Audio system is working!")
            result = 0
        else:
            logger.error("[FAIL] AUDIO STREAM TESTS FAILED")
            result = 1
            
    finally:
        # Clean up
        logger.debug("Terminating PyAudio...")
        try:
            p.terminate()
            logger.info("[OK] PyAudio terminated successfully")
        except Exception as e:
            logger.error(f"[ERROR] Error terminating PyAudio: {e}")
    
    end_time = datetime.now()
    duration = end_time - start_time
    logger.info(f"Test completed at: {end_time}")
    logger.info(f"Total duration: {duration}")
    logger.info("=" * 60)
    
    # Keep console open if running as executable
    if getattr(sys, 'frozen', False):
        input("Press Enter to exit...")
    
    return result

if __name__ == "__main__":
    sys.exit(main())
