#include <windows.h>
#include <memory>
#include <string>
#include <iostream>
#include <fstream>

struct HandlePair {
    HMODULE module;
    void* tts;
};

extern "C" {
    typedef struct {
        float* samples;
        int32_t num_samples;
    } AudioData;

    typedef struct {
        const char* model;
        const char* tokens;
        const char* lexicon;
        float noise_scale;
        float noise_scale_w;
        float length_scale;
    } VitsConfig;

    typedef void* TTSHandle;

    void LogError(const char* message) {
        std::ofstream log("sherpa_native.log", std::ios::app);
        log << "Error: " << message << std::endl;
    }

    void LogAvailableFunctions(HMODULE module) {
        std::ofstream log("sherpa_native.log", std::ios::app);
        log << "Available functions in sherpa-onnx.dll:" << std::endl;

        // Get the DOS header
        PIMAGE_DOS_HEADER dosHeader = (PIMAGE_DOS_HEADER)module;
        if (dosHeader->e_magic != IMAGE_DOS_SIGNATURE) {
            log << "Not a valid PE file" << std::endl;
            return;
        }

        // Get the NT headers
        PIMAGE_NT_HEADERS ntHeader = (PIMAGE_NT_HEADERS)((BYTE*)module + dosHeader->e_lfanew);
        if (ntHeader->Signature != IMAGE_NT_SIGNATURE) {
            log << "Not a valid NT header" << std::endl;
            return;
        }

        // Get the export directory
        DWORD exportDirRVA = ntHeader->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT].VirtualAddress;
        if (exportDirRVA == 0) {
            log << "No export directory found" << std::endl;
            return;
        }

        PIMAGE_EXPORT_DIRECTORY exportDir = (PIMAGE_EXPORT_DIRECTORY)((BYTE*)module + exportDirRVA);
        PDWORD functions = (PDWORD)((BYTE*)module + exportDir->AddressOfFunctions);
        PDWORD names = (PDWORD)((BYTE*)module + exportDir->AddressOfNames);
        PWORD ordinals = (PWORD)((BYTE*)module + exportDir->AddressOfNameOrdinals);

        for (DWORD i = 0; i < exportDir->NumberOfNames; i++) {
            const char* name = (const char*)((BYTE*)module + names[i]);
            log << name << std::endl;
        }
    }

    __declspec(dllexport) TTSHandle __stdcall CreateTTS(const VitsConfig* config) {
        try {
            // Load sherpa-onnx.dll dynamically
            HMODULE sherpa = LoadLibraryA("sherpa-onnx.dll");
            if (!sherpa) {
                DWORD error = GetLastError();
                char buffer[256];
                sprintf_s(buffer, "Failed to load sherpa-onnx.dll. Error code: %d", error);
                LogError(buffer);
                return nullptr;
            }

            // Log available functions
            LogAvailableFunctions(sherpa);

            // Try different function names
            auto create = reinterpret_cast<TTSHandle(__stdcall*)(const VitsConfig*)>(
                GetProcAddress(sherpa, "CreateOfflineTts"));
            if (!create) {
                create = reinterpret_cast<TTSHandle(__stdcall*)(const VitsConfig*)>(
                    GetProcAddress(sherpa, "CreateTts"));
            }
            if (!create) {
                create = reinterpret_cast<TTSHandle(__stdcall*)(const VitsConfig*)>(
                    GetProcAddress(sherpa, "CreateVitsTts"));
            }

            if (!create) {
                LogError("Failed to get TTS creation function");
                FreeLibrary(sherpa);
                return nullptr;
            }

            // Create the TTS instance
            TTSHandle tts = create(config);
            if (!tts) {
                LogError("Failed to create TTS instance");
                FreeLibrary(sherpa);
                return nullptr;
            }

            // Store both the module handle and TTS handle
            auto* pair = new HandlePair{ sherpa, tts };
            return pair;
        }
        catch (const std::exception& e) {
            LogError(e.what());
            return nullptr;
        }
    }

    __declspec(dllexport) AudioData* __stdcall GenerateAudio(TTSHandle handle, const char* text, float speed, int32_t speaker_id) {
        if (!handle) {
            LogError("Invalid handle");
            return nullptr;
        }

        try {
            auto* pair = static_cast<HandlePair*>(handle);
            
            // Try different function names
            auto generate = reinterpret_cast<AudioData*(__stdcall*)(TTSHandle, const char*, float, int32_t)>(
                GetProcAddress(pair->module, "GenerateAudio"));
            if (!generate) {
                generate = reinterpret_cast<AudioData*(__stdcall*)(TTSHandle, const char*, float, int32_t)>(
                    GetProcAddress(pair->module, "Generate"));
            }
            if (!generate) {
                generate = reinterpret_cast<AudioData*(__stdcall*)(TTSHandle, const char*, float, int32_t)>(
                    GetProcAddress(pair->module, "GenerateVits"));
            }

            if (!generate) {
                LogError("Failed to get audio generation function");
                return nullptr;
            }

            // Generate audio
            return generate(pair->tts, text, speed, speaker_id);
        }
        catch (const std::exception& e) {
            LogError(e.what());
            return nullptr;
        }
    }

    __declspec(dllexport) void __stdcall DestroyTTS(TTSHandle handle) {
        if (!handle) return;

        try {
            auto* pair = static_cast<HandlePair*>(handle);
            
            // Try different function names
            auto destroy = reinterpret_cast<void(__stdcall*)(TTSHandle)>(
                GetProcAddress(pair->module, "DestroyOfflineTts"));
            if (!destroy) {
                destroy = reinterpret_cast<void(__stdcall*)(TTSHandle)>(
                    GetProcAddress(pair->module, "DestroyTts"));
            }
            if (!destroy) {
                destroy = reinterpret_cast<void(__stdcall*)(TTSHandle)>(
                    GetProcAddress(pair->module, "DestroyVitsTts"));
            }

            if (destroy) {
                destroy(pair->tts);
            }

            FreeLibrary(pair->module);
            delete pair;
        }
        catch (const std::exception& e) {
            LogError(e.what());
        }
    }
}
