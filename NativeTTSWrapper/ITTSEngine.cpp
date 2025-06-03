#include "stdafx.h"
#include "ITTSEngine.h"
#include "SherpaOnnxEngine.h"
#include "AzureTTSEngine.h"
#include <memory>

namespace NativeTTS {

    std::unique_ptr<ITTSEngine> TTSEngineFactory::CreateEngine(EngineType type, const std::wstring& config) {
        switch (type) {
            case EngineType::SherpaOnnx:
                return std::make_unique<SherpaOnnxEngine>();
                
            case EngineType::Azure:
                return std::make_unique<AzureTTSEngine>();
                
            case EngineType::Plugin:
                // TODO: Implement plugin engine
                return nullptr;
                
            case EngineType::Mock:
                // TODO: Implement mock engine for testing
                return nullptr;
                
            default:
                return nullptr;
        }
    }

    EngineType TTSEngineFactory::GetEngineTypeFromString(const std::wstring& typeName) {
        if (typeName == L"sherpaonnx" || typeName == L"sherpa" || typeName == L"SherpaOnnx") {
            return EngineType::SherpaOnnx;
        }
        else if (typeName == L"azure" || typeName == L"Azure" || typeName == L"AzureTTS") {
            return EngineType::Azure;
        }
        else if (typeName == L"plugin" || typeName == L"Plugin") {
            return EngineType::Plugin;
        }
        else if (typeName == L"mock" || typeName == L"Mock") {
            return EngineType::Mock;
        }
        else {
            // Default to SherpaOnnx for backward compatibility
            return EngineType::SherpaOnnx;
        }
    }

    std::wstring TTSEngineFactory::GetStringFromEngineType(EngineType type) {
        switch (type) {
            case EngineType::SherpaOnnx:
                return L"sherpaonnx";
            case EngineType::Azure:
                return L"azure";
            case EngineType::Plugin:
                return L"plugin";
            case EngineType::Mock:
                return L"mock";
            default:
                return L"unknown";
        }
    }

} // namespace NativeTTS
