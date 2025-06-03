#pragma once

#include "resource.h"
#include "NativeTTSWrapper_i.h"

class CNativeTTSWrapperModule : public ATL::CAtlDllModuleT< CNativeTTSWrapperModule >
{
public :
    DECLARE_LIBID(LIBID_NativeTTSWrapperLib)
    DECLARE_REGISTRY_APPID_RESOURCEID(IDR_NATIVETTSRAPPER, "{B8F4A8E2-9C3D-4A5E-8F7C-2D1B3E4F5A6B}")
};

extern class CNativeTTSWrapperModule _AtlModule;
