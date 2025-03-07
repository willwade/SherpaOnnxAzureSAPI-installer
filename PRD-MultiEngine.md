# Product Requirements Document: Multi-Engine TTS Integration Framework

## Overview
This document outlines the requirements and implementation details for refactoring the OpenSpeech TTS SAPI Installer into a plugin-based architecture that supports multiple TTS engines. The goal is to create a flexible, extensible framework that can easily incorporate new TTS engines such as ElevenLabs and PlayHT alongside the existing Sherpa ONNX and Azure TTS engines.

## Project Goals

1. Create a plugin-based architecture for TTS engines
2. Refactor existing code to use the new architecture
3. Implement ElevenLabs and PlayHT as new TTS engines
4. Standardize configuration management across all engines
5. Ensure backward compatibility with existing installations
6. Provide a clean, intuitive user experience for managing multiple TTS engines
7. Eliminate redundant code and improve maintainability

## Key Components

### 1. Core Framework

#### 1.1 Engine Interface and Base Classes
- Create `ITtsEngine` interface defining common functionality for all TTS engines
- Develop `TtsEngineBase` abstract class with shared implementation
- Define standard voice information structure (`TtsVoiceInfo`)
- Implement engine discovery and registration system

#### 1.2 Engine Manager
- Create `TtsEngineManager` to handle engine registration and retrieval
- Implement engine lifecycle management
- Provide centralized access to all registered engines

#### 1.3 Configuration Management
- Design unified configuration system for all engines
- Implement secure credential storage with encryption
- Support environment variables for sensitive information
- Create configuration file templates for each engine

#### 1.4 SAPI Integration
- Standardize SAPI registration process across engines
- Create base SAPI implementation classes
- Define common registry structure with engine-specific extensions

### 2. Engine Implementations

#### 2.1 Refactor Existing Engines
- Migrate Sherpa ONNX to new architecture
- Migrate Azure TTS to new architecture
- Ensure backward compatibility with existing installations

#### 2.2 New Engine: ElevenLabs
- Implement ElevenLabs API integration
- Create ElevenLabs SAPI voice implementation
- Support ElevenLabs-specific features (voice cloning, etc.)

#### 2.3 New Engine: PlayHT
- Implement PlayHT API integration
- Create PlayHT SAPI voice implementation
- Support PlayHT-specific features (voice customization, etc.)

### 3. User Interface

#### 3.1 Command-Line Interface
- Update CLI to support multiple engines
- Implement engine-specific command options
- Create unified help system

#### 3.2 Interactive Mode
- Redesign interactive flow for multiple engines
- Implement engine selection menu
- Create engine-specific configuration prompts

### 4. Installation and Deployment

#### 4.1 Plugin System
- Implement plugin discovery and loading
- Create plugin directory structure
- Support dynamic loading of engine assemblies

#### 4.2 Installer Updates
- Update WiX installer to support multiple engines
- Create engine-specific installation options
- Implement plugin deployment

## Implementation Phases

### Phase 1: Architecture and Planning (2 weeks)
- Design detailed architecture
- Create interface and base class definitions
- Plan migration strategy for existing code
- Set up project structure

### Phase 2: Core Framework Implementation (3 weeks)
- Implement `ITtsEngine` interface and `TtsEngineBase` class
- Develop `TtsEngineManager`
- Create configuration management system
- Implement plugin loading system

### Phase 3: Existing Engine Migration (2 weeks)
- Refactor Sherpa ONNX to use new architecture
- Refactor Azure TTS to use new architecture
- Update tests for refactored engines
- Ensure backward compatibility

### Phase 4: New Engine Implementation (4 weeks)
- Implement ElevenLabs engine (2 weeks)
- Implement PlayHT engine (2 weeks)
- Create tests for new engines

### Phase 5: User Interface Updates (2 weeks)
- Update command-line interface
- Redesign interactive mode
- Implement engine-specific help

### Phase 6: Cleanup and Optimization (1 week)
- Remove redundant code
- Optimize performance
- Improve error handling
- Update documentation

### Phase 7: Testing and Deployment (2 weeks)
- Comprehensive testing across all engines
- Update installer
- Prepare release

## Detailed Requirements

### 1. Core Framework

#### 1.1 Engine Interface (`ITtsEngine`)

```csharp
public interface ITtsEngine
{
    // Basic properties
    string EngineName { get; }
    string EngineVersion { get; }
    string EngineDescription { get; }
    bool RequiresSsml { get; }
    bool RequiresAuthentication { get; }
    bool SupportsOfflineUsage { get; }
    
    // Configuration
    IEnumerable<ConfigurationParameter> GetRequiredParameters();
    bool ValidateConfiguration(Dictionary<string, string> config);
    
    // Voice management
    Task<IEnumerable<TtsVoiceInfo>> GetAvailableVoicesAsync(Dictionary<string, string> config);
    Task<bool> TestVoiceAsync(string voiceId, Dictionary<string, string> config);
    
    // Speech synthesis
    Task<byte[]> SynthesizeSpeechAsync(string text, string voiceId, Dictionary<string, string> parameters);
    Task<Stream> SynthesizeSpeechToStreamAsync(string text, string voiceId, Dictionary<string, string> parameters);
    
    // SAPI integration
    Guid GetEngineClsid();
    Type GetSapiImplementationType();
    void RegisterVoice(TtsVoiceInfo voice, Dictionary<string, string> config, string dllPath);
    void UnregisterVoice(string voiceId);
    
    // Lifecycle
    void Initialize();
    void Shutdown();
}
```

#### 1.2 Base Engine Class (`TtsEngineBase`)

```csharp
public abstract class TtsEngineBase : ITtsEngine
{
    // Common implementation for all engines
    public abstract string EngineName { get; }
    public virtual string EngineVersion => "1.0";
    public virtual string EngineDescription => $"{EngineName} Text-to-Speech Engine";
    public virtual bool RequiresSsml => false;
    public virtual bool RequiresAuthentication => true;
    public virtual bool SupportsOfflineUsage => false;
    
    // Default implementations that can be overridden
    public virtual IEnumerable<ConfigurationParameter> GetRequiredParameters() => new List<ConfigurationParameter>();
    public virtual Guid GetEngineClsid() => Guid.NewGuid();
    
    // Common functionality
    protected virtual void LogError(string message, Exception ex = null) 
    {
        // Common logging implementation
    }
    
    // Other shared functionality
}
```

#### 1.3 Voice Information (`TtsVoiceInfo`)

```csharp
public class TtsVoiceInfo
{
    // Basic properties
    public string Id { get; set; }
    public string Name { get; set; }
    public string DisplayName { get; set; }
    public string Gender { get; set; }
    public string Locale { get; set; }
    public string Age { get; set; } = "Adult";
    
    // Engine-specific properties
    public string EngineName { get; set; }
    public Dictionary<string, string> AdditionalAttributes { get; set; } = new Dictionary<string, string>();
    
    // Feature support
    public bool SupportsStyles { get; set; }
    public List<string> SupportedStyles { get; set; } = new List<string>();
    public bool SupportsRoles { get; set; }
    public List<string> SupportedRoles { get; set; } = new List<string>();
    
    // Selected options
    public string SelectedStyle { get; set; }
    public string SelectedRole { get; set; }
}
```

#### 1.4 Configuration Parameter (`ConfigurationParameter`)

```csharp
public class ConfigurationParameter
{
    public string Name { get; set; }
    public string DisplayName { get; set; }
    public string Description { get; set; }
    public bool IsRequired { get; set; } = true;
    public bool IsSecret { get; set; } = false;
    public string DefaultValue { get; set; }
    public List<string> AllowedValues { get; set; } = new List<string>();
    public string ValidationRegex { get; set; }
}
```

### 2. Configuration Management

#### 2.1 Configuration File Structure

```json
{
  "engines": {
    "SherpaOnnx": {
      "enabled": true,
      "parameters": {}
    },
    "AzureTTS": {
      "enabled": true,
      "parameters": {
        "subscriptionKey": "ENCRYPTED:AbC123XyZ...",
        "region": "eastus"
      }
    },
    "ElevenLabs": {
      "enabled": true,
      "parameters": {
        "apiKey": "ENCRYPTED:DeF456UvW..."
      }
    },
    "PlayHT": {
      "enabled": true,
      "parameters": {
        "apiKey": "ENCRYPTED:GhI789RsT...",
        "userId": "user123"
      }
    }
  },
  "defaultEngine": "SherpaOnnx",
  "secureStorage": true,
  "lastUpdated": "2023-06-15T12:34:56Z"
}
```

#### 2.2 Configuration Manager (`ConfigurationManager`)

```csharp
public class ConfigurationManager
{
    // Load configuration
    public EngineConfiguration LoadConfiguration();
    
    // Save configuration
    public void SaveConfiguration(EngineConfiguration config);
    
    // Get engine-specific configuration
    public Dictionary<string, string> GetEngineConfiguration(string engineName);
    
    // Update engine configuration
    public void UpdateEngineConfiguration(string engineName, Dictionary<string, string> parameters);
    
    // Encryption/decryption
    public string EncryptValue(string value);
    public string DecryptValue(string encryptedValue);
    
    // Environment variable support
    public string ResolveEnvironmentVariables(string value);
}
```

### 3. Engine Implementations

#### 3.1 Sherpa ONNX Engine

```csharp
public class SherpaOnnxEngine : TtsEngineBase
{
    public override string EngineName => "SherpaOnnx";
    public override bool RequiresAuthentication => false;
    public override bool SupportsOfflineUsage => true;
    
    public override Guid GetEngineClsid() => new Guid("3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2");
    public override Type GetSapiImplementationType() => typeof(SherpaOnnxSapi5VoiceImpl);
    
    // Implement other methods
}
```

#### 3.2 Azure TTS Engine

```csharp
public class AzureTtsEngine : TtsEngineBase
{
    public override string EngineName => "AzureTTS";
    public override bool RequiresSsml => true;
    
    public override Guid GetEngineClsid() => new Guid("3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3");
    public override Type GetSapiImplementationType() => typeof(AzureSapi5VoiceImpl);
    
    public override IEnumerable<ConfigurationParameter> GetRequiredParameters()
    {
        return new List<ConfigurationParameter>
        {
            new ConfigurationParameter
            {
                Name = "subscriptionKey",
                DisplayName = "Subscription Key",
                Description = "Azure Cognitive Services subscription key",
                IsSecret = true
            },
            new ConfigurationParameter
            {
                Name = "region",
                DisplayName = "Region",
                Description = "Azure region (e.g., eastus, westus)"
            }
        };
    }
    
    // Implement other methods
}
```

#### 3.3 ElevenLabs Engine

```csharp
public class ElevenLabsEngine : TtsEngineBase
{
    public override string EngineName => "ElevenLabs";
    public override string EngineDescription => "ElevenLabs AI Voice Generation";
    
    public override Guid GetEngineClsid() => new Guid("3d8f5c5f-9d6b-4b92-a12b-1a6dff80b6b4");
    public override Type GetSapiImplementationType() => typeof(ElevenLabsSapi5VoiceImpl);
    
    public override IEnumerable<ConfigurationParameter> GetRequiredParameters()
    {
        return new List<ConfigurationParameter>
        {
            new ConfigurationParameter
            {
                Name = "apiKey",
                DisplayName = "API Key",
                Description = "ElevenLabs API key",
                IsSecret = true
            },
            new ConfigurationParameter
            {
                Name = "modelId",
                DisplayName = "Model ID",
                Description = "ElevenLabs model ID (optional)",
                IsRequired = false,
                DefaultValue = "eleven_monolingual_v1"
            }
        };
    }
    
    // Implement other methods
}
```

#### 3.4 PlayHT Engine

```csharp
public class PlayHTEngine : TtsEngineBase
{
    public override string EngineName => "PlayHT";
    public override string EngineDescription => "PlayHT AI Voice Generation";
    
    public override Guid GetEngineClsid() => new Guid("3d8f5c60-9d6b-4b92-a12b-1a6dff80b6b5");
    public override Type GetSapiImplementationType() => typeof(PlayHTSapi5VoiceImpl);
    
    public override IEnumerable<ConfigurationParameter> GetRequiredParameters()
    {
        return new List<ConfigurationParameter>
        {
            new ConfigurationParameter
            {
                Name = "apiKey",
                DisplayName = "API Key",
                Description = "PlayHT API key",
                IsSecret = true
            },
            new ConfigurationParameter
            {
                Name = "userId",
                DisplayName = "User ID",
                Description = "PlayHT user ID"
            }
        };
    }
    
    // Implement other methods
}
```

### 4. Plugin System

#### 4.1 Plugin Discovery and Loading

```csharp
public class PluginLoader
{
    private readonly string _pluginDirectory;
    
    public PluginLoader(string pluginDirectory)
    {
        _pluginDirectory = pluginDirectory;
    }
    
    public IEnumerable<ITtsEngine> DiscoverEngines()
    {
        var engines = new List<ITtsEngine>();
        
        // Load built-in engines
        engines.Add(new SherpaOnnxEngine());
        engines.Add(new AzureTtsEngine());
        
        // Load external engines from plugin directory
        if (Directory.Exists(_pluginDirectory))
        {
            foreach (var file in Directory.GetFiles(_pluginDirectory, "*.dll"))
            {
                try
                {
                    var assembly = Assembly.LoadFrom(file);
                    
                    foreach (var type in assembly.GetTypes()
                        .Where(t => typeof(ITtsEngine).IsAssignableFrom(t) && !t.IsAbstract))
                    {
                        var engine = (ITtsEngine)Activator.CreateInstance(type);
                        engines.Add(engine);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading plugin {file}: {ex.Message}");
                }
            }
        }
        
        return engines;
    }
}
```

#### 4.2 Plugin Directory Structure

```
C:\Program Files\OpenAssistive\OpenSpeech\
├── OpenSpeechTTS.dll
├── plugins\
│   ├── ElevenLabs\
│   │   ├── ElevenLabsEngine.dll
│   │   └── dependencies\
│   ├── PlayHT\
│   │   ├── PlayHTEngine.dll
│   │   └── dependencies\
│   └── Custom\
│       └── CustomEngine.dll
└── models\
    └── sherpa-onnx\
        └── ...
```

### 5. User Interface Updates

#### 5.1 Command-Line Interface

```
# List available engines
TTSInstaller.exe list-engines

# List configuration parameters for an engine
TTSInstaller.exe list-params <engine-name>

# Configure an engine
TTSInstaller.exe configure <engine-name> --param1 value1 --param2 value2

# List available voices for an engine
TTSInstaller.exe list-voices <engine-name>

# Install a voice
TTSInstaller.exe install <engine-name> <voice-id> [--param1 value1] [--param2 value2]

# Uninstall a voice
TTSInstaller.exe uninstall <voice-id>

# Uninstall all voices for an engine
TTSInstaller.exe uninstall-engine <engine-name>

# Test a voice
TTSInstaller.exe test <voice-id> [--text "Text to speak"]
```

#### 5.2 Interactive Mode Flow

1. Main Menu
   - List installed engines
   - Configure engines
   - Install voices
   - Uninstall voices
   - Test voices
   - Exit

2. Engine Configuration
   - Select engine
   - Enter configuration parameters
   - Test configuration
   - Save configuration

3. Voice Installation
   - Select engine
   - Search for voices
   - Select voice
   - Configure voice-specific options
   - Install voice

4. Voice Testing
   - List installed voices
   - Select voice
   - Enter text to speak
   - Test voice

### 6. Registry Structure

```
HKLM\SOFTWARE\Microsoft\SPEECH\Voices\Tokens\<voice-name>
    CLSID = "{engine-specific-clsid}"
    Path = "C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll"
    
    \Attributes
        Language = "0409"  # LCID
        Gender = "Female"
        Age = "Adult"
        Vendor = "Vendor Name"
        Version = "1.0"
        Name = "<voice-name>"
        VoiceType = "EngineName"  # e.g., "SherpaOnnx", "AzureTTS", "ElevenLabs", "PlayHT"
        
        # Engine-specific attributes stored as needed
```

## Cleanup and Reorganization

### 1. Code Reorganization

#### 1.1 Directory Structure

```
/Installer
  /Core
    /Interfaces
      ITtsEngine.cs
      ISapiVoiceImpl.cs
    /Base
      TtsEngineBase.cs
      SapiVoiceImplBase.cs
    /Models
      TtsVoiceInfo.cs
      ConfigurationParameter.cs
    /Managers
      TtsEngineManager.cs
      ConfigurationManager.cs
      PluginLoader.cs
  /Engines
    /SherpaOnnx
      SherpaOnnxEngine.cs
      SherpaOnnxSapi5VoiceImpl.cs
    /Azure
      AzureTtsEngine.cs
      AzureSapi5VoiceImpl.cs
    /ElevenLabs
      ElevenLabsEngine.cs
      ElevenLabsSapi5VoiceImpl.cs
    /PlayHT
      PlayHTEngine.cs
      PlayHTSapi5VoiceImpl.cs
  /UI
    Program.cs
    CommandLineParser.cs
    InteractiveMode.cs
  /Utilities
    RegistryHelper.cs
    LoggingHelper.cs
    SecurityHelper.cs
```

#### 1.2 File Cleanup

1. Identify and remove redundant files:
   - Duplicate model classes
   - Deprecated implementations
   - Unused utility classes

2. Consolidate similar functionality:
   - Merge similar utility methods
   - Create shared base classes
   - Extract common code into helper methods

3. Remove dead code:
   - Unused methods
   - Commented-out code
   - Debug-only functionality

### 2. Configuration Consolidation

1. Create a unified configuration system:
   - Migrate Azure config to new format
   - Create templates for new engines
   - Implement secure storage for all engines

2. Standardize configuration access:
   - Use ConfigurationManager for all engines
   - Implement environment variable resolution
   - Support command-line overrides

### 3. Testing and Validation

1. Create comprehensive tests:
   - Unit tests for each engine
   - Integration tests for SAPI registration
   - End-to-end tests for voice installation

2. Validate backward compatibility:
   - Test with existing Sherpa ONNX installations
   - Test with existing Azure TTS installations
   - Ensure smooth migration path

## Security Considerations

### 1. Credential Storage

1. Encrypt sensitive information:
   - Use Windows DPAPI for local encryption
   - Support custom encryption providers
   - Never store credentials in plain text

2. Provide multiple storage options:
   - Configuration file with encryption
   - Windows Credential Manager
   - Environment variables

3. Implement secure deletion:
   - Securely clear credentials from memory
   - Provide option to remove stored credentials

### 2. API Key Management

1. Validate API keys before use:
   - Test connectivity with minimal API calls
   - Provide clear error messages for invalid keys
   - Support key rotation

2. Implement rate limiting:
   - Respect API rate limits
   - Implement exponential backoff
   - Cache results when appropriate

### 3. Plugin Security

1. Validate plugins before loading:
   - Check digital signatures
   - Verify against known plugins
   - Sandbox plugin execution

2. Limit plugin capabilities:
   - Restrict file system access
   - Control network access
   - Prevent registry modifications outside allowed paths

## Testing Plan

### 1. Unit Testing

1. Core components:
   - Engine interfaces and base classes
   - Configuration management
   - Plugin loading

2. Engine implementations:
   - SherpaOnnx engine
   - Azure TTS engine
   - ElevenLabs engine
   - PlayHT engine

### 2. Integration Testing

1. SAPI integration:
   - Voice registration
   - Voice unregistration
   - SAPI voice selection

2. Configuration management:
   - Configuration loading/saving
   - Encryption/decryption
   - Environment variable resolution

### 3. End-to-End Testing

1. Installation scenarios:
   - Fresh installation
   - Upgrade from previous version
   - Multiple engines installation

2. Voice management:
   - Voice installation
   - Voice uninstallation
   - Voice testing

3. User interface:
   - Command-line interface
   - Interactive mode
   - Error handling

## Documentation

### 1. User Documentation

1. Installation guide:
   - System requirements
   - Installation steps
   - Troubleshooting

2. Engine configuration:
   - Configuration parameters
   - API key acquisition
   - Security considerations

3. Voice management:
   - Voice installation
   - Voice uninstallation
   - Voice testing

### 2. Developer Documentation

1. Architecture overview:
   - Component diagram
   - Interaction flow
   - Extension points

2. Plugin development:
   - Interface implementation
   - Plugin packaging
   - Testing guidelines

3. API reference:
   - Core interfaces
   - Base classes
   - Utility methods

## Success Criteria

1. All existing functionality is preserved
2. ElevenLabs and PlayHT engines are fully implemented
3. Configuration management is unified across all engines
4. Plugin system supports dynamic loading of new engines
5. User interface is intuitive and consistent
6. Documentation is comprehensive and up-to-date
7. All tests pass with high code coverage
8. No redundant or deprecated code remains

## Future Considerations

1. Additional engines:
   - Google Cloud TTS
   - Amazon Polly
   - IBM Watson TTS
   - OpenAI TTS

2. Advanced features:
   - Voice cloning
   - Emotion control
   - Pronunciation customization
   - Batch processing

3. Integration options:
   - Web API for remote access
   - Integration with assistive technology
   - Mobile app support

## Conclusion

This project will transform the OpenSpeech TTS SAPI Installer into a flexible, extensible framework for multiple TTS engines. By implementing a plugin-based architecture, we will enable easy integration of new engines while maintaining backward compatibility with existing installations. The unified configuration management system will provide a consistent user experience across all engines, with secure storage of credentials and support for environment variables.

The addition of ElevenLabs and PlayHT engines will significantly expand the voice options available to users, while the cleanup and reorganization phase will ensure the codebase remains maintainable and efficient. With comprehensive testing and documentation, we will ensure a smooth transition to the new architecture and provide a solid foundation for future enhancements. 