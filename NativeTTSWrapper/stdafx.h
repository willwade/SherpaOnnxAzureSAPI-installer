#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN
#pragma warning(disable: 4996) // Disable deprecation warnings
#include <windows.h>

#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>

#include <sapi.h>
#include <sphelper.h>

#include <string>
#include <vector>
#include <memory>
#include <fstream>
#include <sstream>
#include <map>
#include <mutex>
#include <codecvt>
#include <locale>

// Third-party libraries
#include <nlohmann/json.hpp>
#include <spdlog/spdlog.h>

// TTS Engine interfaces
#include "ITTSEngine.h"

using namespace ATL;
