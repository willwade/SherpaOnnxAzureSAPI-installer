#!/usr/bin/env python3
"""
Diagnostic script to check audio functionality in the installed AACSpeakHelper.
This script should be run in the same directory as the installed AACSpeakHelperServer.exe
"""

import os
import subprocess

def check_installed_dlls():
    """Check if the required DLLs are present in the installed version"""
    print("=== Checking Installed DLLs ===")
    
    # Get the directory where this script is running (should be the install directory)
    install_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"Install directory: {install_dir}")
    
    # Look for the _internal directory (PyInstaller onedir structure)
    internal_dir = os.path.join(install_dir, "_internal")
    if not os.path.exists(internal_dir):
        print("❌ _internal directory not found - this might not be a PyInstaller build")
        return False
    
    print(f"Internal directory: {internal_dir}")
    
    # Check for required audio DLLs
    required_files = [
        "_sounddevice_data/portaudio-binaries/libportaudio64bit.dll",
        "_sounddevice_data/portaudio-binaries/libportaudio64bit-asio.dll", 
        "pyaudio/_portaudio.cp311-win_amd64.pyd"
    ]
    
    found_files = []
    for file_path in required_files:
        full_path = os.path.join(internal_dir, file_path)
        if os.path.exists(full_path):
            found_files.append(file_path)
            print(f"✅ Found: {file_path}")
        else:
            print(f"❌ Missing: {file_path}")
    
    print(f"DLL check: {len(found_files)}/{len(required_files)} found")
    return len(found_files) == len(required_files)

def test_server_audio():
    """Test if the AACSpeakHelper server can handle audio requests"""
    print("\n=== Testing Server Audio ===")
    
    # Look for the server executable
    install_dir = os.path.dirname(os.path.abspath(__file__))
    server_exe = os.path.join(install_dir, "AACSpeakHelperServer.exe")
    
    if not os.path.exists(server_exe):
        print(f"❌ Server executable not found: {server_exe}")
        return False
    
    print(f"Server executable: {server_exe}")
    
    # Try to start the server and send a test request
    try:
        print("🔊 Starting server for audio test...")
        
        # Start the server in the background
        server_process = subprocess.Popen(
            [server_exe],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            cwd=install_dir
        )
        
        # Give it a moment to start
        import time
        time.sleep(2)
        
        # Check if the server is still running
        if server_process.poll() is not None:
            stdout, stderr = server_process.communicate()
            print("❌ Server exited immediately")
            print(f"STDOUT: {stdout.decode()}")
            print(f"STDERR: {stderr.decode()}")
            return False
        
        print("✅ Server started successfully")
        
        # Try to connect and send a test message
        try:
            import win32file
            
            pipe_name = r'\\.\pipe\AACSpeakHelper'
            print(f"Connecting to pipe: {pipe_name}")
            
            # Try to connect to the named pipe
            handle = win32file.CreateFile(
                pipe_name,
                win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                0,
                None,
                win32file.OPEN_EXISTING,
                0,
                None
            )
            
            print("✅ Connected to server pipe")
            
            # Send a test message
            test_message = "hello"
            win32file.WriteFile(handle, test_message.encode())
            print(f"✅ Sent test message: {test_message}")
            
            # Give it time to process
            time.sleep(3)
            
            win32file.CloseHandle(handle)
            print("✅ Test completed")
            
        except Exception as e:
            print(f"❌ Pipe communication failed: {e}")
            return False
        finally:
            # Clean up the server process
            try:
                server_process.terminate()
                server_process.wait(timeout=5)
            except:
                server_process.kill()
        
        return True
        
    except Exception as e:
        print(f"❌ Server test failed: {e}")
        return False

def check_log_files():
    """Check the log files for audio errors"""
    print("\n=== Checking Log Files ===")
    
    # Common log file locations
    log_locations = [
        os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Ace Centre", "AACSpeakHelper", "app.log"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.log"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "userslogs", "app.log"),
    ]
    
    for log_path in log_locations:
        if os.path.exists(log_path):
            print(f"📄 Found log file: {log_path}")
            
            # Read the last few lines to check for audio errors
            try:
                with open(log_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    
                # Look for recent audio errors
                audio_errors = []
                for line in lines[-50:]:  # Check last 50 lines
                    if "Device unavailable" in line or "errno -9985" in line.lower():
                        audio_errors.append(line.strip())
                
                if audio_errors:
                    print(f"❌ Found {len(audio_errors)} audio errors in log:")
                    for error in audio_errors[-3:]:  # Show last 3 errors
                        print(f"   {error}")
                else:
                    print("✅ No recent audio errors found in log")
                    
            except Exception as e:
                print(f"❌ Could not read log file: {e}")
        else:
            print(f"📄 Log file not found: {log_path}")

def main():
    """Main diagnostic function"""
    print("🔊 AACSpeakHelper Audio Diagnostic")
    print("=" * 50)
    print("This script checks for audio issues in the installed AACSpeakHelper")
    print()
    
    # Run diagnostics
    tests = [
        ("DLL Check", check_installed_dlls),
        ("Log File Check", check_log_files),
        ("Server Audio Test", test_server_audio),
    ]
    
    passed = 0
    total = len(tests)
    
    for test_name, test_func in tests:
        print(f"Running: {test_name}")
        try:
            if test_func():
                passed += 1
        except Exception as e:
            print(f"❌ Test failed with exception: {e}")
        print()
    
    print("=" * 50)
    print(f"Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All tests passed! Audio should be working.")
    else:
        print("💥 Some tests failed. Audio issues detected.")
        print("\nTroubleshooting suggestions:")
        print("1. Make sure no other applications are using the audio device")
        print("2. Try restarting the Windows Audio service")
        print("3. Check if the AACSpeakHelper was built with the latest audio fixes")
        print("4. Run this diagnostic as Administrator")
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
