# Test interface querying to see if SAPI can find our interfaces
Write-Host "Testing Interface Query on COM Objects" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Green

# Create our COM object
Write-Host "1. Creating our COM object..." -ForegroundColor Cyan
try {
    $ourObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    Write-Host "   ✅ Our COM object created successfully" -ForegroundColor Green
    
    # Test if we can cast to the interfaces
    Write-Host "2. Testing interface casting..." -ForegroundColor Cyan
    
    # Get the type and check interfaces
    $type = $ourObject.GetType()
    Write-Host "   Object type: $($type.FullName)" -ForegroundColor White
    
    $interfaces = $type.GetInterfaces()
    Write-Host "   Implemented interfaces:" -ForegroundColor White
    foreach($interface in $interfaces) {
        Write-Host "     - $($interface.Name) ($($interface.GUID))" -ForegroundColor Gray
    }
    
    # Try to get specific methods
    Write-Host "3. Checking specific methods..." -ForegroundColor Cyan
    
    $speakMethod = $type.GetMethod("Speak")
    if($speakMethod) {
        Write-Host "   ✅ Speak method found" -ForegroundColor Green
        Write-Host "     Parameters: $($speakMethod.GetParameters().Length)" -ForegroundColor Gray
    } else {
        Write-Host "   ❌ Speak method not found" -ForegroundColor Red
    }
    
    $getOutputFormatMethod = $type.GetMethod("GetOutputFormat")
    if($getOutputFormatMethod) {
        Write-Host "   ✅ GetOutputFormat method found" -ForegroundColor Green
        Write-Host "     Parameters: $($getOutputFormatMethod.GetParameters().Length)" -ForegroundColor Gray
    } else {
        Write-Host "   ❌ GetOutputFormat method not found" -ForegroundColor Red
    }
    
    $setObjectTokenMethod = $type.GetMethod("SetObjectToken")
    if($setObjectTokenMethod) {
        Write-Host "   ✅ SetObjectToken method found" -ForegroundColor Green
        Write-Host "     Parameters: $($setObjectTokenMethod.GetParameters().Length)" -ForegroundColor Gray
    } else {
        Write-Host "   ❌ SetObjectToken method not found" -ForegroundColor Red
    }
    
} catch {
    Write-Host "❌ Error creating our COM object: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "4. Testing SAPI's ability to create our object..." -ForegroundColor Cyan

try {
    # Try to create our object using the CLSID directly
    $clsid = "{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}"
    Write-Host "   Attempting to create object with CLSID: $clsid" -ForegroundColor Yellow
    
    # This is how SAPI would create our object
    $sapiObject = [System.Activator]::CreateInstance([System.Type]::GetTypeFromCLSID([System.Guid]::new($clsid)))
    
    if($sapiObject) {
        Write-Host "   ✅ SAPI-style object creation successful" -ForegroundColor Green
        
        # Check if this object has the same interfaces
        $sapiType = $sapiObject.GetType()
        Write-Host "   SAPI-created object type: $($sapiType.FullName)" -ForegroundColor White
        
        # Try to call SetObjectToken on this object
        $setTokenMethod = $sapiType.GetMethod("SetObjectToken")
        if($setTokenMethod) {
            Write-Host "   ✅ SetObjectToken method accessible via SAPI creation" -ForegroundColor Green
        } else {
            Write-Host "   ❌ SetObjectToken method not accessible via SAPI creation" -ForegroundColor Red
        }
    } else {
        Write-Host "   ❌ SAPI-style object creation failed" -ForegroundColor Red
    }
    
} catch {
    Write-Host "   ❌ Error in SAPI-style creation: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== ANALYSIS ===" -ForegroundColor Cyan
Write-Host "This test helps determine if:" -ForegroundColor Yellow
Write-Host "1. Our COM object can be created properly" -ForegroundColor White
Write-Host "2. Our interfaces are accessible" -ForegroundColor White
Write-Host "3. SAPI can create our object the same way we can" -ForegroundColor White
Write-Host "4. There are differences in how SAPI vs PowerShell creates objects" -ForegroundColor White
