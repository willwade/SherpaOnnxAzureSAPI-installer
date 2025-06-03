#pragma once

#include <windows.h>
#include <vector>
#include <string>
#include <memory>

namespace NativeTTS {

    /// <summary>
    /// Abstract interface for TTS engines
    /// This allows us to support multiple TTS backends (SherpaOnnx, Azure, future engines)
    /// while maintaining a consistent interface for the COM wrapper
    /// </summary>
    class ITTSEngine {
    public:
        virtual ~ITTSEngine() = default;

        /// <summary>
        /// Initialize the TTS engine with configuration
        /// </summary>
        /// <param name="config">JSON configuration string for the engine</param>
        /// <returns>S_OK on success, error HRESULT on failure</returns>
        virtual HRESULT Initialize(const std::wstring& config) = 0;

        /// <summary>
        /// Generate audio samples from text
        /// </summary>
        /// <param name="text">Text to synthesize</param>
        /// <param name="samples">Output audio samples (32-bit float, normalized -1.0 to 1.0)</param>
        /// <param name="sampleRate">Output sample rate in Hz</param>
        /// <returns>S_OK on success, error HRESULT on failure</returns>
        virtual HRESULT Generate(const std::wstring& text, std::vector<float>& samples, int& sampleRate) = 0;

        /// <summary>
        /// Shutdown the engine and release resources
        /// </summary>
        /// <returns>S_OK on success, error HRESULT on failure</returns>
        virtual HRESULT Shutdown() = 0;

        /// <summary>
        /// Check if the engine is initialized and ready
        /// </summary>
        /// <returns>true if initialized, false otherwise</returns>
        virtual bool IsInitialized() const = 0;

        /// <summary>
        /// Get engine-specific information (name, version, etc.)
        /// </summary>
        /// <returns>Engine information string</returns>
        virtual std::wstring GetEngineInfo() const = 0;

        /// <summary>
        /// Get supported audio format information
        /// </summary>
        /// <param name="sampleRate">Preferred sample rate</param>
        /// <param name="channels">Number of channels (1=mono, 2=stereo)</param>
        /// <param name="bitsPerSample">Bits per sample (16 or 32)</param>
        /// <returns>S_OK if format is supported</returns>
        virtual HRESULT GetSupportedFormat(int& sampleRate, int& channels, int& bitsPerSample) const = 0;
    };

    /// <summary>
    /// Engine types supported by the system
    /// </summary>
    enum class EngineType {
        SherpaOnnx,
        Azure,
        Plugin,
        Mock  // For testing
    };

    /// <summary>
    /// Factory for creating TTS engines
    /// </summary>
    class TTSEngineFactory {
    public:
        /// <summary>
        /// Create a TTS engine of the specified type
        /// </summary>
        /// <param name="type">Type of engine to create</param>
        /// <param name="config">Configuration for the engine</param>
        /// <returns>Unique pointer to the created engine, or nullptr on failure</returns>
        static std::unique_ptr<ITTSEngine> CreateEngine(EngineType type, const std::wstring& config);

        /// <summary>
        /// Get engine type from string name
        /// </summary>
        /// <param name="typeName">String name of the engine type</param>
        /// <returns>EngineType enum value</returns>
        static EngineType GetEngineTypeFromString(const std::wstring& typeName);

        /// <summary>
        /// Get string name from engine type
        /// </summary>
        /// <param name="type">EngineType enum value</param>
        /// <returns>String name of the engine type</returns>
        static std::wstring GetStringFromEngineType(EngineType type);
    };

} // namespace NativeTTS
