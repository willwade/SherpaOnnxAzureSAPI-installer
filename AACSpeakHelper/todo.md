Complete the SAPI integration by implementing audio byte streaming between the COM wrapper and AACSpeakHelper service. Currently, the COM wrapper successfully communicates with AACSpeakHelper, but SAPI requires audio bytes to be returned through the interface.

📋 Task List
Phase 1: AACSpeakHelper Audio Streaming Enhancement
Task 1.1: Add audio bytes return parameter
Modify AACSpeakHelperServer.py to accept return_audio_bytes parameter in JSON requests
Update message parsing to detect when audio bytes should be returned instead of played
Add logging to track when byte streaming mode is requested
Task 1.2: Implement speak_to_bytes function
Create speak_to_bytes() function in tts_utils.py that returns audio bytes
Use tts.speak(text) method instead of tts.speak_streamed(text) to get byte data
Support all TTS engines (Azure TTS, SherpaOnnx, Google TTS, etc.)
Add proper error handling and logging
Task 1.3: Implement pipe audio streaming
Modify pipe server to send audio bytes back through the named pipe
Add proper byte length headers for reliable data transmission
Implement chunked transmission for large audio files
Add timeout handling for pipe write operations
Phase 2: COM Wrapper Audio Reception
Task 2.1: Update COM wrapper message format
Add return_audio_bytes: true parameter to JSON messages sent to AACSpeakHelper
Ensure all required parameters (listvoices, engine config) are included
Update message creation functions to handle the new parameter
Task 2.2: Implement pipe audio reception
Replace placeholder audio generation with actual pipe audio reading
Implement chunked audio data reception with proper buffering
Add timeout handling for pipe read operations
Validate received audio data format and size
Task 2.3: Audio format conversion
Ensure received audio matches SAPI's expected format (22050Hz, 16-bit, mono)
Add audio format validation and conversion if needed
Handle different audio formats from different TTS engines
Add proper WAV header validation
Phase 3: Testing and Validation
Task 3.1: Unit testing
Test speak_to_bytes() function with all supported TTS engines
Verify audio byte generation for Azure TTS and SherpaOnnx
Test pipe communication with various message sizes
Add automated tests for audio format validation
Task 3.2: Integration testing
Test complete SAPI workflow: voice selection → text input → audio output
Verify audio quality and timing with real SAPI applications
Test with multiple concurrent SAPI requests
Validate memory usage and performance under load
Task 3.3: Real-world application testing
Test with Windows Narrator and other screen readers
Test with Notepad "Speak selected text" feature
Test with third-party SAPI applications
Verify voice switching and rate/volume controls work correctly
Phase 4: Performance Optimization
Task 4.1: Streaming optimization
Implement real-time audio streaming instead of buffering entire audio
Add audio chunk streaming for faster response times
Optimize pipe buffer sizes for best performance
Add connection pooling for multiple SAPI requests
Task 4.2: Caching and efficiency
Implement audio caching for repeated text synthesis
Add TTS engine connection pooling to reduce initialization overhead
Optimize JSON message parsing and creation
Add performance monitoring and logging
Phase 5: Error Handling and Robustness
Task 5.1: Comprehensive error handling
Add graceful fallbacks when AACSpeakHelper is unavailable
Implement retry logic for failed pipe communications
Add proper error codes and messages for SAPI applications
Handle TTS engine failures gracefully
Task 5.2: Service reliability
Add automatic AACSpeakHelper service restart on failure
Implement health checking for TTS engines
Add proper cleanup for abandoned pipe connections
Add logging and monitoring for production deployment
Phase 6: Documentation and Deployment
Task 6.1: Update documentation
Document the new audio streaming architecture
Add troubleshooting guide for audio streaming issues
Update installation instructions with new requirements
Add performance tuning guide
Task 6.2: Deployment preparation
Create automated installer that sets up both COM wrapper and AACSpeakHelper
Add service installation scripts for AACSpeakHelper
Create uninstaller that properly cleans up all components
Add version compatibility checking
🎯 Priority Order
High Priority: Tasks 1.2, 2.1, 2.2 (Core audio streaming functionality)
Medium Priority: Tasks 1.1, 1.3, 2.3, 3.2 (Integration and format handling)
Low Priority: Tasks 3.1, 4.x, 5.x (Testing, optimization, robustness)
🔧 Technical Notes
tts-wrapper methods: Use tts.speak(text) for bytes, tts.speak_streamed(text) for playback
Pipe communication: Named pipe \\.\pipe\AACSpeakHelper for IPC
Audio format: SAPI expects 22050Hz, 16-bit, mono PCM audio
Message format: JSON with return_audio_bytes: true parameter
⚠️ Dependencies
AACSpeakHelper service must be running
tts-wrapper library with byte streaming support
Proper COM wrapper registration with correct CLSID
Window