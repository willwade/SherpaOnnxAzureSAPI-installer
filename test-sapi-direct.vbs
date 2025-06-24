' Test our ISpVoice through Windows SAPI

WScript.Echo "=== Testing OpenSpeechSpVoice through Windows SAPI ===" 

On Error Resume Next

' Test 1: Create standard Windows SAPI SpVoice
WScript.Echo ""
WScript.Echo "1. Creating standard Windows SAPI SpVoice..."
Set sapi = CreateObject("SAPI.SpVoice")
If Err.Number <> 0 Then
    WScript.Echo "FAILED: Could not create SAPI.SpVoice"
    WScript.Echo "Error: " & Err.Description
    WScript.Quit 1
End If
WScript.Echo "SUCCESS: Standard SAPI SpVoice created"

' Test 2: Get available voices
WScript.Echo ""
WScript.Echo "2. Getting available SAPI voices..."
Set voices = sapi.GetVoices()
WScript.Echo "Found " & voices.Count & " voices:"

For i = 0 To voices.Count - 1
    Set voice = voices.Item(i)
    voiceName = voice.GetDescription()
    WScript.Echo "  Voice " & i & ": " & voiceName
    
    ' Check if this is one of our voices
    If InStr(voiceName, "OpenSpeech") > 0 Or InStr(voiceName, "NativeTTS") > 0 Or InStr(voiceName, "Azure") > 0 Then
        WScript.Echo "    ^ This might be one of our voices!"
    End If
Next

' Test 3: Try to use our voice if we can find it
WScript.Echo ""
WScript.Echo "3. Testing speech with default voice..."
Err.Clear
sapi.Speak "Hello from Windows SAPI! This is a test to see if our voice integration is working.", 0
If Err.Number <> 0 Then
    WScript.Echo "FAILED: Speech test failed"
    WScript.Echo "Error: " & Err.Description
Else
    WScript.Echo "SUCCESS: Speech test completed with default voice"
End If

' Test 4: Try to create our COM object and see if SAPI can use it
WScript.Echo ""
WScript.Echo "4. Testing if SAPI can work with our COM object..."

' Try to create our object
Set ourVoice = GetObject("new:{F2E8B6A1-3C4D-4E5F-8A7B-9C1D2E3F4A5B}")
If Err.Number <> 0 Then
    WScript.Echo "FAILED: Could not create our COM object"
    WScript.Echo "Error: " & Err.Description
Else
    WScript.Echo "SUCCESS: Our COM object created"
    
    ' Try to see if we can use it with SAPI somehow
    ' This is experimental - trying to see if there's any connection
    WScript.Echo "Object created, but SAPI integration requires proper voice registration"
End If

' Test 5: Check voice tokens
WScript.Echo ""
WScript.Echo "5. Checking voice tokens in detail..."
For i = 0 To voices.Count - 1
    Set voice = voices.Item(i)
    voiceName = voice.GetDescription()
    
    ' Try to get more details about each voice
    Err.Clear
    Set attributes = voice.GetAttributes()
    If Err.Number = 0 Then
        WScript.Echo "Voice " & i & " (" & voiceName & ") attributes:"
        ' Try to get CLSID
        clsid = attributes.GetStringValue("CLSID")
        If Err.Number = 0 Then
            WScript.Echo "  CLSID: " & clsid
            If clsid = "{F2E8B6A1-3C4D-4E5F-8A7B-9C1D2E3F4A5B}" Then
                WScript.Echo "  ^ This is our OpenSpeechSpVoice!"
            End If
        End If
    End If
Next

Set sapi = Nothing
Set ourVoice = Nothing

WScript.Echo ""
WScript.Echo "=== SAPI Direct Test Complete ==="
