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

using namespace ATL;
