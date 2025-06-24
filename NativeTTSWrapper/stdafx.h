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

// Include generated COM interface definitions
#include "NativeTTSWrapper_i.h"

#include <string>
#include <vector>
#include <memory>
#include <fstream>
#include <sstream>
#include <map>
#include <mutex>
#include <codecvt>
#include <locale>

// Pure pipe client - no third-party libraries needed

using namespace ATL;
