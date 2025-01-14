
//  Microsoft (R) .NET Framework IL Disassembler.  Version 4.8.3928.0
//  Copyright (c) Microsoft Corporation.  All rights reserved.



// Metadata version: v4.0.30319
.module extern 'sherpa-onnx-c-api'
.assembly extern netstandard
{
  .publickeytoken = (CC 7B 13 FF CD 2D DD 51 )                         // .{...-.Q
  .ver 2:0:0:0
}
.assembly 'sherpa-onnx'
{
  .custom instance void [netstandard]System.Runtime.CompilerServices.CompilationRelaxationsAttribute::.ctor(int32) = ( 01 00 08 00 00 00 00 00 ) 
  .custom instance void [netstandard]System.Runtime.CompilerServices.RuntimeCompatibilityAttribute::.ctor() = ( 01 00 01 00 54 02 16 57 72 61 70 4E 6F 6E 45 78   // ....T..WrapNonEx
                                                                                                                63 65 70 74 69 6F 6E 54 68 72 6F 77 73 01 )       // ceptionThrows.

  // --- The following custom attribute is added automatically, do not uncomment -------
  //  .custom instance void [netstandard]System.Diagnostics.DebuggableAttribute::.ctor(valuetype [netstandard]System.Diagnostics.DebuggableAttribute/DebuggingModes) = ( 01 00 02 00 00 00 00 00 ) 

  .custom instance void [netstandard]System.Runtime.Versioning.TargetFrameworkAttribute::.ctor(string) = ( 01 00 19 2E 4E 45 54 53 74 61 6E 64 61 72 64 2C   // ....NETStandard,
                                                                                                           56 65 72 73 69 6F 6E 3D 76 32 2E 30 01 00 54 0E   // Version=v2.0..T.
                                                                                                           14 46 72 61 6D 65 77 6F 72 6B 44 69 73 70 6C 61   // .FrameworkDispla
                                                                                                           79 4E 61 6D 65 11 2E 4E 45 54 20 53 74 61 6E 64   // yName..NET Stand
                                                                                                           61 72 64 20 32 2E 30 )                            // ard 2.0
  .custom instance void [netstandard]System.Reflection.AssemblyCompanyAttribute::.ctor(string) = ( 01 00 12 58 69 61 6F 6D 69 20 43 6F 72 70 6F 72   // ...Xiaomi Corpor
                                                                                                   61 74 69 6F 6E 00 00 )                            // ation..
  .custom instance void [netstandard]System.Reflection.AssemblyConfigurationAttribute::.ctor(string) = ( 01 00 07 52 65 6C 65 61 73 65 00 00 )             // ...Release..
  .custom instance void [netstandard]System.Reflection.AssemblyCopyrightAttribute::.ctor(string) = ( 01 00 26 43 6F 70 79 72 69 67 68 74 20 32 30 31   // ..&Copyright 201
                                                                                                     39 2D 32 30 32 33 20 58 69 61 6F 6D 69 20 43 6F   // 9-2023 Xiaomi Co
                                                                                                     72 70 6F 72 61 74 69 6F 6E 00 00 )                // rporation..
  .custom instance void [netstandard]System.Reflection.AssemblyDescriptionAttribute::.ctor(string) = ( 01 00 81 8B 73 68 65 72 70 61 2D 6F 6E 6E 78 20   // ....sherpa-onnx 
                                                                                                       69 73 20 61 6E 20 6F 70 65 6E 2D 73 6F 75 72 63   // is an open-sourc
                                                                                                       65 20 72 65 61 6C 2D 74 69 6D 65 20 73 70 65 65   // e real-time spee
                                                                                                       63 68 20 72 65 63 6F 67 6E 69 74 69 6F 6E 20 74   // ch recognition t
                                                                                                       6F 6F 6C 6B 69 74 20 64 65 76 65 6C 6F 70 65 64   // oolkit developed
                                                                                                       0A 20 20 20 20 62 79 20 74 68 65 20 4E 65 78 74   // .    by the Next
                                                                                                       2D 67 65 6E 20 4B 61 6C 64 69 20 74 65 61 6D 2E   // -gen Kaldi team.
                                                                                                       20 49 74 20 73 75 70 70 6F 72 74 73 20 73 74 72   //  It supports str
                                                                                                       65 61 6D 69 6E 67 20 72 65 63 6F 67 6E 69 74 69   // eaming recogniti
                                                                                                       6F 6E 20 6F 6E 20 61 20 76 61 72 69 65 74 79 20   // on on a variety 
                                                                                                       6F 66 0A 20 20 20 20 70 6C 61 74 66 6F 72 6D 73   // of.    platforms
                                                                                                       20 73 75 63 68 20 61 73 20 41 6E 64 72 6F 69 64   //  such as Android
                                                                                                       2C 20 69 4F 53 2C 20 52 61 73 70 62 65 72 72 79   // , iOS, Raspberry
                                                                                                       2C 20 4C 69 6E 75 78 2C 20 57 69 6E 64 6F 77 73   // , Linux, Windows
                                                                                                       2C 20 6D 61 63 4F 53 2C 20 65 74 63 2E 0A 0A 20   // , macOS, etc... 
                                                                                                       20 20 20 49 74 20 64 6F 65 73 20 6E 6F 74 20 72   //    It does not r
                                                                                                       65 71 75 69 72 65 20 49 6E 74 65 72 6E 65 74 20   // equire Internet 
                                                                                                       63 6F 6E 6E 65 63 74 69 6F 6E 20 64 75 72 69 6E   // connection durin
                                                                                                       67 20 72 65 63 6F 67 6E 69 74 69 6F 6E 2E 0A 0A   // g recognition...
                                                                                                       20 20 20 20 53 65 65 20 74 68 65 20 64 6F 63 75   //     See the docu
                                                                                                       6D 65 6E 74 61 74 69 6F 6E 20 68 74 74 70 73 3A   // mentation https:
                                                                                                       2F 2F 6B 32 2D 66 73 61 2E 67 69 74 68 75 62 2E   // //k2-fsa.github.
                                                                                                       69 6F 2F 73 68 65 72 70 61 2F 6F 6E 6E 78 2F 69   // io/sherpa/onnx/i
                                                                                                       6E 64 65 78 2E 68 74 6D 6C 0A 20 20 20 20 66 6F   // ndex.html.    fo
                                                                                                       72 20 64 65 74 61 69 6C 73 2E 0A 20 20 20 20 00   // r details..    .
                                                                                                       00 ) 
  .custom instance void [netstandard]System.Reflection.AssemblyFileVersionAttribute::.ctor(string) = ( 01 00 08 31 2E 39 2E 31 32 2E 30 00 00 )          // ...1.9.12.0..
  .custom instance void [netstandard]System.Reflection.AssemblyInformationalVersionAttribute::.ctor(string) = ( 01 00 06 31 2E 39 2E 31 32 00 00 )                // ...1.9.12..
  .custom instance void [netstandard]System.Reflection.AssemblyProductAttribute::.ctor(string) = ( 01 00 0B 73 68 65 72 70 61 2D 6F 6E 6E 78 00 00 ) // ...sherpa-onnx..
  .custom instance void [netstandard]System.Reflection.AssemblyTitleAttribute::.ctor(string) = ( 01 00 0B 73 68 65 72 70 61 2D 6F 6E 6E 78 00 00 ) // ...sherpa-onnx..
  .custom instance void [netstandard]System.Reflection.AssemblyMetadataAttribute::.ctor(string,
                                                                                        string) = ( 01 00 0D 52 65 70 6F 73 69 74 6F 72 79 55 72 6C   // ...RepositoryUrl
                                                                                                    25 68 74 74 70 73 3A 2F 2F 67 69 74 68 75 62 2E   // %https://github.
                                                                                                    63 6F 6D 2F 6B 32 2D 66 73 61 2F 73 68 65 72 70   // com/k2-fsa/sherp
                                                                                                    61 2D 6F 6E 6E 78 00 00 )                         // a-onnx..
  .permissionset reqmin
             = {[netstandard]System.Security.Permissions.SecurityPermissionAttribute = {property bool 'SkipVerification' = bool(true)}}
  .hash algorithm 0x00008004
  .ver 1:9:12:0
}
.module 'sherpa-onnx.dll'
// MVID: {77FD7CD0-6167-448D-B160-3B5E4AAE8CDF}
.custom instance void [netstandard]System.Security.UnverifiableCodeAttribute::.ctor() = ( 01 00 00 00 ) 
.imagebase 0x10000000
.file alignment 0x00000200
.stackreserve 0x00100000
.subsystem 0x0003       // WINDOWS_CUI
.corflags 0x00000001    //  ILONLY
// Image base: 0x000001B1A6F50000


// =============== CLASS MEMBERS DECLARATION ===================

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OfflineTtsVitsModelConfig
       extends [netstandard]System.ValueType
{
  .field public  marshal( lpstr) string Model
  .field public  marshal( lpstr) string Lexicon
  .field public  marshal( lpstr) string Tokens
  .field public  marshal( lpstr) string DataDir
  .field public float32 NoiseScale
  .field public float32 NoiseScaleW
  .field public float32 LengthScale
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       78 (0x4e)
    .maxstack  2
    IL_0000:  ldarg.0
    IL_0001:  ldstr      ""
    IL_0006:  stfld      string SherpaOnnx.OfflineTtsVitsModelConfig::Model
    IL_000b:  ldarg.0
    IL_000c:  ldstr      ""
    IL_0011:  stfld      string SherpaOnnx.OfflineTtsVitsModelConfig::Lexicon
    IL_0016:  ldarg.0
    IL_0017:  ldstr      ""
    IL_001c:  stfld      string SherpaOnnx.OfflineTtsVitsModelConfig::Tokens
    IL_0021:  ldarg.0
    IL_0022:  ldstr      ""
    IL_0027:  stfld      string SherpaOnnx.OfflineTtsVitsModelConfig::DataDir
    IL_002c:  ldarg.0
    IL_002d:  ldc.r4     0.667
    IL_0032:  stfld      float32 SherpaOnnx.OfflineTtsVitsModelConfig::NoiseScale
    IL_0037:  ldarg.0
    IL_0038:  ldc.r4     0.80000001
    IL_003d:  stfld      float32 SherpaOnnx.OfflineTtsVitsModelConfig::NoiseScaleW
    IL_0042:  ldarg.0
    IL_0043:  ldc.r4     1.
    IL_0048:  stfld      float32 SherpaOnnx.OfflineTtsVitsModelConfig::LengthScale
    IL_004d:  ret
  } // end of method OfflineTtsVitsModelConfig::.ctor

} // end of class SherpaOnnx.OfflineTtsVitsModelConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OfflineTtsModelConfig
       extends [netstandard]System.ValueType
{
  .field public valuetype SherpaOnnx.OfflineTtsVitsModelConfig Vits
  .field public int32 NumThreads
  .field public int32 Debug
  .field public  marshal( lpstr) string Provider
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       37 (0x25)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  newobj     instance void SherpaOnnx.OfflineTtsVitsModelConfig::.ctor()
    IL_0006:  stfld      valuetype SherpaOnnx.OfflineTtsVitsModelConfig SherpaOnnx.OfflineTtsModelConfig::Vits
    IL_000b:  ldarg.0
    IL_000c:  ldc.i4.1
    IL_000d:  stfld      int32 SherpaOnnx.OfflineTtsModelConfig::NumThreads
    IL_0012:  ldarg.0
    IL_0013:  ldc.i4.0
    IL_0014:  stfld      int32 SherpaOnnx.OfflineTtsModelConfig::Debug
    IL_0019:  ldarg.0
    IL_001a:  ldstr      "cpu"
    IL_001f:  stfld      string SherpaOnnx.OfflineTtsModelConfig::Provider
    IL_0024:  ret
  } // end of method OfflineTtsModelConfig::.ctor

} // end of class SherpaOnnx.OfflineTtsModelConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OfflineTtsConfig
       extends [netstandard]System.ValueType
{
  .field public valuetype SherpaOnnx.OfflineTtsModelConfig Model
  .field public  marshal( lpstr) string RuleFsts
  .field public int32 MaxNumSentences
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       30 (0x1e)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  newobj     instance void SherpaOnnx.OfflineTtsModelConfig::.ctor()
    IL_0006:  stfld      valuetype SherpaOnnx.OfflineTtsModelConfig SherpaOnnx.OfflineTtsConfig::Model
    IL_000b:  ldarg.0
    IL_000c:  ldstr      ""
    IL_0011:  stfld      string SherpaOnnx.OfflineTtsConfig::RuleFsts
    IL_0016:  ldarg.0
    IL_0017:  ldc.i4.1
    IL_0018:  stfld      int32 SherpaOnnx.OfflineTtsConfig::MaxNumSentences
    IL_001d:  ret
  } // end of method OfflineTtsConfig::.ctor

} // end of class SherpaOnnx.OfflineTtsConfig

.class public auto ansi beforefieldinit SherpaOnnx.OfflineTtsGeneratedAudio
       extends [netstandard]System.Object
{
  .class sequential ansi sealed nested private beforefieldinit Impl
         extends [netstandard]System.ValueType
  {
    .field public native int Samples
    .field public int32 NumSamples
    .field public int32 SampleRate
  } // end of class Impl

  .field private valuetype [netstandard]System.Runtime.InteropServices.HandleRef _handle
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(native int p) cil managed
  {
    // Code size       20 (0x14)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [netstandard]System.Object::.ctor()
    IL_0006:  ldarg.0
    IL_0007:  ldarg.0
    IL_0008:  ldarg.1
    IL_0009:  newobj     instance void [netstandard]System.Runtime.InteropServices.HandleRef::.ctor(object,
                                                                                                    native int)
    IL_000e:  stfld      valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineTtsGeneratedAudio::_handle
    IL_0013:  ret
  } // end of method OfflineTtsGeneratedAudio::.ctor

  .method public hidebysig instance bool 
          SaveToWaveFile(string filename) cil managed
  {
    // Code size       55 (0x37)
    .maxstack  4
    .locals init (valuetype SherpaOnnx.OfflineTtsGeneratedAudio/Impl V_0)
    IL_0000:  ldarg.0
    IL_0001:  call       instance native int SherpaOnnx.OfflineTtsGeneratedAudio::get_Handle()
    IL_0006:  ldtoken    SherpaOnnx.OfflineTtsGeneratedAudio/Impl
    IL_000b:  call       class [netstandard]System.Type [netstandard]System.Type::GetTypeFromHandle(valuetype [netstandard]System.RuntimeTypeHandle)
    IL_0010:  call       object [netstandard]System.Runtime.InteropServices.Marshal::PtrToStructure(native int,
                                                                                                    class [netstandard]System.Type)
    IL_0015:  unbox.any  SherpaOnnx.OfflineTtsGeneratedAudio/Impl
    IL_001a:  stloc.0
    IL_001b:  ldloc.0
    IL_001c:  ldfld      native int SherpaOnnx.OfflineTtsGeneratedAudio/Impl::Samples
    IL_0021:  ldloc.0
    IL_0022:  ldfld      int32 SherpaOnnx.OfflineTtsGeneratedAudio/Impl::NumSamples
    IL_0027:  ldloc.0
    IL_0028:  ldfld      int32 SherpaOnnx.OfflineTtsGeneratedAudio/Impl::SampleRate
    IL_002d:  ldarg.1
    IL_002e:  call       int32 SherpaOnnx.OfflineTtsGeneratedAudio::SherpaOnnxWriteWave(native int,
                                                                                        int32,
                                                                                        int32,
                                                                                        string)
    IL_0033:  ldc.i4.1
    IL_0034:  ceq
    IL_0036:  ret
  } // end of method OfflineTtsGeneratedAudio::SaveToWaveFile

  .method family hidebysig virtual instance void 
          Finalize() cil managed
  {
    .override [netstandard]System.Object::Finalize
    // Code size       16 (0x10)
    .maxstack  1
    .try
    {
      IL_0000:  ldarg.0
      IL_0001:  call       instance void SherpaOnnx.OfflineTtsGeneratedAudio::Cleanup()
      IL_0006:  leave.s    IL_000f

    }  // end .try
    finally
    {
      IL_0008:  ldarg.0
      IL_0009:  call       instance void [netstandard]System.Object::Finalize()
      IL_000e:  endfinally
    }  // end handler
    IL_000f:  ret
  } // end of method OfflineTtsGeneratedAudio::Finalize

  .method public hidebysig instance void 
          Dispose() cil managed
  {
    // Code size       13 (0xd)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void SherpaOnnx.OfflineTtsGeneratedAudio::Cleanup()
    IL_0006:  ldarg.0
    IL_0007:  call       void [netstandard]System.GC::SuppressFinalize(object)
    IL_000c:  ret
  } // end of method OfflineTtsGeneratedAudio::Dispose

  .method private hidebysig instance void 
          Cleanup() cil managed
  {
    // Code size       29 (0x1d)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance native int SherpaOnnx.OfflineTtsGeneratedAudio::get_Handle()
    IL_0006:  call       void SherpaOnnx.OfflineTtsGeneratedAudio::SherpaOnnxDestroyOfflineTtsGeneratedAudio(native int)
    IL_000b:  ldarg.0
    IL_000c:  ldarg.0
    IL_000d:  ldsfld     native int [netstandard]System.IntPtr::Zero
    IL_0012:  newobj     instance void [netstandard]System.Runtime.InteropServices.HandleRef::.ctor(object,
                                                                                                    native int)
    IL_0017:  stfld      valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineTtsGeneratedAudio::_handle
    IL_001c:  ret
  } // end of method OfflineTtsGeneratedAudio::Cleanup

  .method public hidebysig specialname instance native int 
          get_Handle() cil managed
  {
    // Code size       12 (0xc)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineTtsGeneratedAudio::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  ret
  } // end of method OfflineTtsGeneratedAudio::get_Handle

  .method public hidebysig specialname instance int32 
          get_NumSamples() cil managed
  {
    // Code size       32 (0x20)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance native int SherpaOnnx.OfflineTtsGeneratedAudio::get_Handle()
    IL_0006:  ldtoken    SherpaOnnx.OfflineTtsGeneratedAudio/Impl
    IL_000b:  call       class [netstandard]System.Type [netstandard]System.Type::GetTypeFromHandle(valuetype [netstandard]System.RuntimeTypeHandle)
    IL_0010:  call       object [netstandard]System.Runtime.InteropServices.Marshal::PtrToStructure(native int,
                                                                                                    class [netstandard]System.Type)
    IL_0015:  unbox.any  SherpaOnnx.OfflineTtsGeneratedAudio/Impl
    IL_001a:  ldfld      int32 SherpaOnnx.OfflineTtsGeneratedAudio/Impl::NumSamples
    IL_001f:  ret
  } // end of method OfflineTtsGeneratedAudio::get_NumSamples

  .method public hidebysig specialname instance int32 
          get_SampleRate() cil managed
  {
    // Code size       32 (0x20)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance native int SherpaOnnx.OfflineTtsGeneratedAudio::get_Handle()
    IL_0006:  ldtoken    SherpaOnnx.OfflineTtsGeneratedAudio/Impl
    IL_000b:  call       class [netstandard]System.Type [netstandard]System.Type::GetTypeFromHandle(valuetype [netstandard]System.RuntimeTypeHandle)
    IL_0010:  call       object [netstandard]System.Runtime.InteropServices.Marshal::PtrToStructure(native int,
                                                                                                    class [netstandard]System.Type)
    IL_0015:  unbox.any  SherpaOnnx.OfflineTtsGeneratedAudio/Impl
    IL_001a:  ldfld      int32 SherpaOnnx.OfflineTtsGeneratedAudio/Impl::SampleRate
    IL_001f:  ret
  } // end of method OfflineTtsGeneratedAudio::get_SampleRate

  .method public hidebysig specialname instance float32[] 
          get_Samples() cil managed
  {
    // Code size       60 (0x3c)
    .maxstack  4
    .locals init (valuetype SherpaOnnx.OfflineTtsGeneratedAudio/Impl V_0,
             float32[] V_1)
    IL_0000:  ldarg.0
    IL_0001:  call       instance native int SherpaOnnx.OfflineTtsGeneratedAudio::get_Handle()
    IL_0006:  ldtoken    SherpaOnnx.OfflineTtsGeneratedAudio/Impl
    IL_000b:  call       class [netstandard]System.Type [netstandard]System.Type::GetTypeFromHandle(valuetype [netstandard]System.RuntimeTypeHandle)
    IL_0010:  call       object [netstandard]System.Runtime.InteropServices.Marshal::PtrToStructure(native int,
                                                                                                    class [netstandard]System.Type)
    IL_0015:  unbox.any  SherpaOnnx.OfflineTtsGeneratedAudio/Impl
    IL_001a:  stloc.0
    IL_001b:  ldloc.0
    IL_001c:  ldfld      int32 SherpaOnnx.OfflineTtsGeneratedAudio/Impl::NumSamples
    IL_0021:  newarr     [netstandard]System.Single
    IL_0026:  stloc.1
    IL_0027:  ldloc.0
    IL_0028:  ldfld      native int SherpaOnnx.OfflineTtsGeneratedAudio/Impl::Samples
    IL_002d:  ldloc.1
    IL_002e:  ldc.i4.0
    IL_002f:  ldloc.0
    IL_0030:  ldfld      int32 SherpaOnnx.OfflineTtsGeneratedAudio/Impl::NumSamples
    IL_0035:  call       void [netstandard]System.Runtime.InteropServices.Marshal::Copy(native int,
                                                                                        float32[],
                                                                                        int32,
                                                                                        int32)
    IL_003a:  ldloc.1
    IL_003b:  ret
  } // end of method OfflineTtsGeneratedAudio::get_Samples

  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          void  SherpaOnnxDestroyOfflineTtsGeneratedAudio(native int handle) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          int32  SherpaOnnxWriteWave(native int samples,
                                     int32 n,
                                     int32 sample_rate,
                                     string  marshal( lpstr) filename) cil managed preservesig
  {
  }
  .property instance native int Handle()
  {
    .get instance native int SherpaOnnx.OfflineTtsGeneratedAudio::get_Handle()
  } // end of property OfflineTtsGeneratedAudio::Handle
  .property instance int32 NumSamples()
  {
    .get instance int32 SherpaOnnx.OfflineTtsGeneratedAudio::get_NumSamples()
  } // end of property OfflineTtsGeneratedAudio::NumSamples
  .property instance int32 SampleRate()
  {
    .get instance int32 SherpaOnnx.OfflineTtsGeneratedAudio::get_SampleRate()
  } // end of property OfflineTtsGeneratedAudio::SampleRate
  .property instance float32[] Samples()
  {
    .get instance float32[] SherpaOnnx.OfflineTtsGeneratedAudio::get_Samples()
  } // end of property OfflineTtsGeneratedAudio::Samples
} // end of class SherpaOnnx.OfflineTtsGeneratedAudio

.class public auto ansi sealed SherpaOnnx.OfflineTtsCallback
       extends [netstandard]System.MulticastDelegate
{
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(object 'object',
                               native int 'method') runtime managed
  {
  } // end of method OfflineTtsCallback::.ctor

  .method public hidebysig newslot virtual 
          instance void  Invoke(native int samples,
                                int32 n) runtime managed
  {
  } // end of method OfflineTtsCallback::Invoke

  .method public hidebysig newslot virtual 
          instance class [netstandard]System.IAsyncResult 
          BeginInvoke(native int samples,
                      int32 n,
                      class [netstandard]System.AsyncCallback callback,
                      object 'object') runtime managed
  {
  } // end of method OfflineTtsCallback::BeginInvoke

  .method public hidebysig newslot virtual 
          instance void  EndInvoke(class [netstandard]System.IAsyncResult result) runtime managed
  {
  } // end of method OfflineTtsCallback::EndInvoke

} // end of class SherpaOnnx.OfflineTtsCallback

.class public auto ansi beforefieldinit SherpaOnnx.OfflineTts
       extends [netstandard]System.Object
       implements [netstandard]System.IDisposable
{
  .field private valuetype [netstandard]System.Runtime.InteropServices.HandleRef _handle
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(valuetype SherpaOnnx.OfflineTtsConfig config) cil managed
  {
    // Code size       28 (0x1c)
    .maxstack  3
    .locals init (native int V_0)
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [netstandard]System.Object::.ctor()
    IL_0006:  ldarga.s   config
    IL_0008:  call       native int SherpaOnnx.OfflineTts::SherpaOnnxCreateOfflineTts(valuetype SherpaOnnx.OfflineTtsConfig&)
    IL_000d:  stloc.0
    IL_000e:  ldarg.0
    IL_000f:  ldarg.0
    IL_0010:  ldloc.0
    IL_0011:  newobj     instance void [netstandard]System.Runtime.InteropServices.HandleRef::.ctor(object,
                                                                                                    native int)
    IL_0016:  stfld      valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineTts::_handle
    IL_001b:  ret
  } // end of method OfflineTts::.ctor

  .method public hidebysig instance class SherpaOnnx.OfflineTtsGeneratedAudio 
          Generate(string text,
                   float32 speed,
                   int32 speakerId) cil managed
  {
    // Code size       25 (0x19)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineTts::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  ldarg.1
    IL_000c:  ldarg.3
    IL_000d:  ldarg.2
    IL_000e:  call       native int SherpaOnnx.OfflineTts::SherpaOnnxOfflineTtsGenerate(native int,
                                                                                        string,
                                                                                        int32,
                                                                                        float32)
    IL_0013:  newobj     instance void SherpaOnnx.OfflineTtsGeneratedAudio::.ctor(native int)
    IL_0018:  ret
  } // end of method OfflineTts::Generate

  .method public hidebysig instance class SherpaOnnx.OfflineTtsGeneratedAudio 
          GenerateWithCallback(string text,
                               float32 speed,
                               int32 speakerId,
                               class SherpaOnnx.OfflineTtsCallback callback) cil managed
  {
    // Code size       27 (0x1b)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineTts::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  ldarg.1
    IL_000c:  ldarg.3
    IL_000d:  ldarg.2
    IL_000e:  ldarg.s    callback
    IL_0010:  call       native int SherpaOnnx.OfflineTts::SherpaOnnxOfflineTtsGenerateWithCallback(native int,
                                                                                                    string,
                                                                                                    int32,
                                                                                                    float32,
                                                                                                    class SherpaOnnx.OfflineTtsCallback)
    IL_0015:  newobj     instance void SherpaOnnx.OfflineTtsGeneratedAudio::.ctor(native int)
    IL_001a:  ret
  } // end of method OfflineTts::GenerateWithCallback

  .method public hidebysig newslot virtual final 
          instance void  Dispose() cil managed
  {
    // Code size       13 (0xd)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void SherpaOnnx.OfflineTts::Cleanup()
    IL_0006:  ldarg.0
    IL_0007:  call       void [netstandard]System.GC::SuppressFinalize(object)
    IL_000c:  ret
  } // end of method OfflineTts::Dispose

  .method family hidebysig virtual instance void 
          Finalize() cil managed
  {
    .override [netstandard]System.Object::Finalize
    // Code size       16 (0x10)
    .maxstack  1
    .try
    {
      IL_0000:  ldarg.0
      IL_0001:  call       instance void SherpaOnnx.OfflineTts::Cleanup()
      IL_0006:  leave.s    IL_000f

    }  // end .try
    finally
    {
      IL_0008:  ldarg.0
      IL_0009:  call       instance void [netstandard]System.Object::Finalize()
      IL_000e:  endfinally
    }  // end handler
    IL_000f:  ret
  } // end of method OfflineTts::Finalize

  .method private hidebysig instance void 
          Cleanup() cil managed
  {
    // Code size       34 (0x22)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineTts::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  call       void SherpaOnnx.OfflineTts::SherpaOnnxDestroyOfflineTts(native int)
    IL_0010:  ldarg.0
    IL_0011:  ldarg.0
    IL_0012:  ldsfld     native int [netstandard]System.IntPtr::Zero
    IL_0017:  newobj     instance void [netstandard]System.Runtime.InteropServices.HandleRef::.ctor(object,
                                                                                                    native int)
    IL_001c:  stfld      valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineTts::_handle
    IL_0021:  ret
  } // end of method OfflineTts::Cleanup

  .method public hidebysig specialname instance int32 
          get_SampleRate() cil managed
  {
    // Code size       17 (0x11)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineTts::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  call       int32 SherpaOnnx.OfflineTts::SherpaOnnxOfflineTtsSampleRate(native int)
    IL_0010:  ret
  } // end of method OfflineTts::get_SampleRate

  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          native int  SherpaOnnxCreateOfflineTts(valuetype SherpaOnnx.OfflineTtsConfig& config) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          void  SherpaOnnxDestroyOfflineTts(native int handle) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          int32  SherpaOnnxOfflineTtsSampleRate(native int handle) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          native int  SherpaOnnxOfflineTtsGenerate(native int handle,
                                                   string  marshal( lpstr) text,
                                                   int32 sid,
                                                   float32 speed) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" cdecl) 
          native int  SherpaOnnxOfflineTtsGenerateWithCallback(native int handle,
                                                               string  marshal( lpstr) text,
                                                               int32 sid,
                                                               float32 speed,
                                                               class SherpaOnnx.OfflineTtsCallback callback) cil managed preservesig
  {
  }
  .property instance int32 SampleRate()
  {
    .get instance int32 SherpaOnnx.OfflineTts::get_SampleRate()
  } // end of property OfflineTts::SampleRate
} // end of class SherpaOnnx.OfflineTts

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OfflineTransducerModelConfig
       extends [netstandard]System.ValueType
{
  .field public  marshal( lpstr) string Encoder
  .field public  marshal( lpstr) string Decoder
  .field public  marshal( lpstr) string Joiner
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       34 (0x22)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldstr      ""
    IL_0006:  stfld      string SherpaOnnx.OfflineTransducerModelConfig::Encoder
    IL_000b:  ldarg.0
    IL_000c:  ldstr      ""
    IL_0011:  stfld      string SherpaOnnx.OfflineTransducerModelConfig::Decoder
    IL_0016:  ldarg.0
    IL_0017:  ldstr      ""
    IL_001c:  stfld      string SherpaOnnx.OfflineTransducerModelConfig::Joiner
    IL_0021:  ret
  } // end of method OfflineTransducerModelConfig::.ctor

} // end of class SherpaOnnx.OfflineTransducerModelConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OfflineParaformerModelConfig
       extends [netstandard]System.ValueType
{
  .field public  marshal( lpstr) string Model
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       12 (0xc)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldstr      ""
    IL_0006:  stfld      string SherpaOnnx.OfflineParaformerModelConfig::Model
    IL_000b:  ret
  } // end of method OfflineParaformerModelConfig::.ctor

} // end of class SherpaOnnx.OfflineParaformerModelConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OfflineNemoEncDecCtcModelConfig
       extends [netstandard]System.ValueType
{
  .field public  marshal( lpstr) string Model
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       12 (0xc)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldstr      ""
    IL_0006:  stfld      string SherpaOnnx.OfflineNemoEncDecCtcModelConfig::Model
    IL_000b:  ret
  } // end of method OfflineNemoEncDecCtcModelConfig::.ctor

} // end of class SherpaOnnx.OfflineNemoEncDecCtcModelConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OfflineWhisperModelConfig
       extends [netstandard]System.ValueType
{
  .field public  marshal( lpstr) string Encoder
  .field public  marshal( lpstr) string Decoder
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       23 (0x17)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldstr      ""
    IL_0006:  stfld      string SherpaOnnx.OfflineWhisperModelConfig::Encoder
    IL_000b:  ldarg.0
    IL_000c:  ldstr      ""
    IL_0011:  stfld      string SherpaOnnx.OfflineWhisperModelConfig::Decoder
    IL_0016:  ret
  } // end of method OfflineWhisperModelConfig::.ctor

} // end of class SherpaOnnx.OfflineWhisperModelConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OfflineTdnnModelConfig
       extends [netstandard]System.ValueType
{
  .field public  marshal( lpstr) string Model
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       12 (0xc)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldstr      ""
    IL_0006:  stfld      string SherpaOnnx.OfflineTdnnModelConfig::Model
    IL_000b:  ret
  } // end of method OfflineTdnnModelConfig::.ctor

} // end of class SherpaOnnx.OfflineTdnnModelConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OfflineLMConfig
       extends [netstandard]System.ValueType
{
  .field public  marshal( lpstr) string Model
  .field public float32 Scale
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       23 (0x17)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldstr      ""
    IL_0006:  stfld      string SherpaOnnx.OfflineLMConfig::Model
    IL_000b:  ldarg.0
    IL_000c:  ldc.r4     0.5
    IL_0011:  stfld      float32 SherpaOnnx.OfflineLMConfig::Scale
    IL_0016:  ret
  } // end of method OfflineLMConfig::.ctor

} // end of class SherpaOnnx.OfflineLMConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OfflineModelConfig
       extends [netstandard]System.ValueType
{
  .field public valuetype SherpaOnnx.OfflineTransducerModelConfig Transducer
  .field public valuetype SherpaOnnx.OfflineParaformerModelConfig Paraformer
  .field public valuetype SherpaOnnx.OfflineNemoEncDecCtcModelConfig NeMoCtc
  .field public valuetype SherpaOnnx.OfflineWhisperModelConfig Whisper
  .field public valuetype SherpaOnnx.OfflineTdnnModelConfig Tdnn
  .field public  marshal( lpstr) string Tokens
  .field public int32 NumThreads
  .field public int32 Debug
  .field public  marshal( lpstr) string Provider
  .field public  marshal( lpstr) string ModelType
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       103 (0x67)
    .maxstack  2
    IL_0000:  ldarg.0
    IL_0001:  newobj     instance void SherpaOnnx.OfflineTransducerModelConfig::.ctor()
    IL_0006:  stfld      valuetype SherpaOnnx.OfflineTransducerModelConfig SherpaOnnx.OfflineModelConfig::Transducer
    IL_000b:  ldarg.0
    IL_000c:  newobj     instance void SherpaOnnx.OfflineParaformerModelConfig::.ctor()
    IL_0011:  stfld      valuetype SherpaOnnx.OfflineParaformerModelConfig SherpaOnnx.OfflineModelConfig::Paraformer
    IL_0016:  ldarg.0
    IL_0017:  newobj     instance void SherpaOnnx.OfflineNemoEncDecCtcModelConfig::.ctor()
    IL_001c:  stfld      valuetype SherpaOnnx.OfflineNemoEncDecCtcModelConfig SherpaOnnx.OfflineModelConfig::NeMoCtc
    IL_0021:  ldarg.0
    IL_0022:  newobj     instance void SherpaOnnx.OfflineWhisperModelConfig::.ctor()
    IL_0027:  stfld      valuetype SherpaOnnx.OfflineWhisperModelConfig SherpaOnnx.OfflineModelConfig::Whisper
    IL_002c:  ldarg.0
    IL_002d:  newobj     instance void SherpaOnnx.OfflineTdnnModelConfig::.ctor()
    IL_0032:  stfld      valuetype SherpaOnnx.OfflineTdnnModelConfig SherpaOnnx.OfflineModelConfig::Tdnn
    IL_0037:  ldarg.0
    IL_0038:  ldstr      ""
    IL_003d:  stfld      string SherpaOnnx.OfflineModelConfig::Tokens
    IL_0042:  ldarg.0
    IL_0043:  ldc.i4.1
    IL_0044:  stfld      int32 SherpaOnnx.OfflineModelConfig::NumThreads
    IL_0049:  ldarg.0
    IL_004a:  ldc.i4.0
    IL_004b:  stfld      int32 SherpaOnnx.OfflineModelConfig::Debug
    IL_0050:  ldarg.0
    IL_0051:  ldstr      "cpu"
    IL_0056:  stfld      string SherpaOnnx.OfflineModelConfig::Provider
    IL_005b:  ldarg.0
    IL_005c:  ldstr      ""
    IL_0061:  stfld      string SherpaOnnx.OfflineModelConfig::ModelType
    IL_0066:  ret
  } // end of method OfflineModelConfig::.ctor

} // end of class SherpaOnnx.OfflineModelConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OfflineRecognizerConfig
       extends [netstandard]System.ValueType
{
  .field public valuetype SherpaOnnx.FeatureConfig FeatConfig
  .field public valuetype SherpaOnnx.OfflineModelConfig ModelConfig
  .field public valuetype SherpaOnnx.OfflineLMConfig LmConfig
  .field public  marshal( lpstr) string DecodingMethod
  .field public int32 MaxActivePaths
  .field public  marshal( lpstr) string HotwordsFile
  .field public float32 HotwordsScore
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       74 (0x4a)
    .maxstack  2
    IL_0000:  ldarg.0
    IL_0001:  newobj     instance void SherpaOnnx.FeatureConfig::.ctor()
    IL_0006:  stfld      valuetype SherpaOnnx.FeatureConfig SherpaOnnx.OfflineRecognizerConfig::FeatConfig
    IL_000b:  ldarg.0
    IL_000c:  newobj     instance void SherpaOnnx.OfflineModelConfig::.ctor()
    IL_0011:  stfld      valuetype SherpaOnnx.OfflineModelConfig SherpaOnnx.OfflineRecognizerConfig::ModelConfig
    IL_0016:  ldarg.0
    IL_0017:  newobj     instance void SherpaOnnx.OfflineLMConfig::.ctor()
    IL_001c:  stfld      valuetype SherpaOnnx.OfflineLMConfig SherpaOnnx.OfflineRecognizerConfig::LmConfig
    IL_0021:  ldarg.0
    IL_0022:  ldstr      "greedy_search"
    IL_0027:  stfld      string SherpaOnnx.OfflineRecognizerConfig::DecodingMethod
    IL_002c:  ldarg.0
    IL_002d:  ldc.i4.4
    IL_002e:  stfld      int32 SherpaOnnx.OfflineRecognizerConfig::MaxActivePaths
    IL_0033:  ldarg.0
    IL_0034:  ldstr      ""
    IL_0039:  stfld      string SherpaOnnx.OfflineRecognizerConfig::HotwordsFile
    IL_003e:  ldarg.0
    IL_003f:  ldc.r4     1.5
    IL_0044:  stfld      float32 SherpaOnnx.OfflineRecognizerConfig::HotwordsScore
    IL_0049:  ret
  } // end of method OfflineRecognizerConfig::.ctor

} // end of class SherpaOnnx.OfflineRecognizerConfig

.class public auto ansi beforefieldinit SherpaOnnx.OfflineRecognizerResult
       extends [netstandard]System.Object
{
  .class sequential ansi sealed nested private beforefieldinit Impl
         extends [netstandard]System.ValueType
  {
    .field public native int Text
  } // end of class Impl

  .field private string _text
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(native int handle) cil managed
  {
    // Code size       109 (0x6d)
    .maxstack  4
    .locals init (valuetype SherpaOnnx.OfflineRecognizerResult/Impl V_0,
             int32 V_1,
             uint8[] V_2,
             uint8* V_3)
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [netstandard]System.Object::.ctor()
    IL_0006:  ldarg.1
    IL_0007:  ldtoken    SherpaOnnx.OfflineRecognizerResult/Impl
    IL_000c:  call       class [netstandard]System.Type [netstandard]System.Type::GetTypeFromHandle(valuetype [netstandard]System.RuntimeTypeHandle)
    IL_0011:  call       object [netstandard]System.Runtime.InteropServices.Marshal::PtrToStructure(native int,
                                                                                                    class [netstandard]System.Type)
    IL_0016:  unbox.any  SherpaOnnx.OfflineRecognizerResult/Impl
    IL_001b:  stloc.0
    IL_001c:  ldc.i4.0
    IL_001d:  stloc.1
    IL_001e:  ldloc.0
    IL_001f:  ldfld      native int SherpaOnnx.OfflineRecognizerResult/Impl::Text
    IL_0024:  call       void* [netstandard]System.IntPtr::op_Explicit(native int)
    IL_0029:  stloc.3
    IL_002a:  br.s       IL_0030

    IL_002c:  ldloc.3
    IL_002d:  ldc.i4.1
    IL_002e:  add
    IL_002f:  stloc.3
    IL_0030:  ldloc.3
    IL_0031:  ldind.u1
    IL_0032:  brtrue.s   IL_002c

    IL_0034:  ldloc.3
    IL_0035:  ldloc.0
    IL_0036:  ldfld      native int SherpaOnnx.OfflineRecognizerResult/Impl::Text
    IL_003b:  call       void* [netstandard]System.IntPtr::op_Explicit(native int)
    IL_0040:  sub
    IL_0041:  ldc.i4.1
    IL_0042:  div
    IL_0043:  conv.i8
    IL_0044:  conv.i4
    IL_0045:  stloc.1
    IL_0046:  ldloc.1
    IL_0047:  newarr     [netstandard]System.Byte
    IL_004c:  stloc.2
    IL_004d:  ldloc.0
    IL_004e:  ldfld      native int SherpaOnnx.OfflineRecognizerResult/Impl::Text
    IL_0053:  ldloc.2
    IL_0054:  ldc.i4.0
    IL_0055:  ldloc.1
    IL_0056:  call       void [netstandard]System.Runtime.InteropServices.Marshal::Copy(native int,
                                                                                        uint8[],
                                                                                        int32,
                                                                                        int32)
    IL_005b:  ldarg.0
    IL_005c:  call       class [netstandard]System.Text.Encoding [netstandard]System.Text.Encoding::get_UTF8()
    IL_0061:  ldloc.2
    IL_0062:  callvirt   instance string [netstandard]System.Text.Encoding::GetString(uint8[])
    IL_0067:  stfld      string SherpaOnnx.OfflineRecognizerResult::_text
    IL_006c:  ret
  } // end of method OfflineRecognizerResult::.ctor

  .method public hidebysig specialname instance string 
          get_Text() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      string SherpaOnnx.OfflineRecognizerResult::_text
    IL_0006:  ret
  } // end of method OfflineRecognizerResult::get_Text

  .property instance string Text()
  {
    .get instance string SherpaOnnx.OfflineRecognizerResult::get_Text()
  } // end of property OfflineRecognizerResult::Text
} // end of class SherpaOnnx.OfflineRecognizerResult

.class public auto ansi beforefieldinit SherpaOnnx.OfflineStream
       extends [netstandard]System.Object
       implements [netstandard]System.IDisposable
{
  .field private valuetype [netstandard]System.Runtime.InteropServices.HandleRef _handle
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(native int p) cil managed
  {
    // Code size       20 (0x14)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [netstandard]System.Object::.ctor()
    IL_0006:  ldarg.0
    IL_0007:  ldarg.0
    IL_0008:  ldarg.1
    IL_0009:  newobj     instance void [netstandard]System.Runtime.InteropServices.HandleRef::.ctor(object,
                                                                                                    native int)
    IL_000e:  stfld      valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineStream::_handle
    IL_0013:  ret
  } // end of method OfflineStream::.ctor

  .method public hidebysig instance void 
          AcceptWaveform(int32 sampleRate,
                         float32[] samples) cil managed
  {
    // Code size       17 (0x11)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance native int SherpaOnnx.OfflineStream::get_Handle()
    IL_0006:  ldarg.1
    IL_0007:  ldarg.2
    IL_0008:  ldarg.2
    IL_0009:  ldlen
    IL_000a:  conv.i4
    IL_000b:  call       void SherpaOnnx.OfflineStream::AcceptWaveform(native int,
                                                                       int32,
                                                                       float32[],
                                                                       int32)
    IL_0010:  ret
  } // end of method OfflineStream::AcceptWaveform

  .method public hidebysig specialname instance class SherpaOnnx.OfflineRecognizerResult 
          get_Result() cil managed
  {
    // Code size       30 (0x1e)
    .maxstack  2
    .locals init (class SherpaOnnx.OfflineRecognizerResult V_0)
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineStream::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  call       native int SherpaOnnx.OfflineStream::GetResult(native int)
    IL_0010:  dup
    IL_0011:  newobj     instance void SherpaOnnx.OfflineRecognizerResult::.ctor(native int)
    IL_0016:  stloc.0
    IL_0017:  call       void SherpaOnnx.OfflineStream::DestroyResult(native int)
    IL_001c:  ldloc.0
    IL_001d:  ret
  } // end of method OfflineStream::get_Result

  .method family hidebysig virtual instance void 
          Finalize() cil managed
  {
    .override [netstandard]System.Object::Finalize
    // Code size       16 (0x10)
    .maxstack  1
    .try
    {
      IL_0000:  ldarg.0
      IL_0001:  call       instance void SherpaOnnx.OfflineStream::Cleanup()
      IL_0006:  leave.s    IL_000f

    }  // end .try
    finally
    {
      IL_0008:  ldarg.0
      IL_0009:  call       instance void [netstandard]System.Object::Finalize()
      IL_000e:  endfinally
    }  // end handler
    IL_000f:  ret
  } // end of method OfflineStream::Finalize

  .method public hidebysig newslot virtual final 
          instance void  Dispose() cil managed
  {
    // Code size       13 (0xd)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void SherpaOnnx.OfflineStream::Cleanup()
    IL_0006:  ldarg.0
    IL_0007:  call       void [netstandard]System.GC::SuppressFinalize(object)
    IL_000c:  ret
  } // end of method OfflineStream::Dispose

  .method private hidebysig instance void 
          Cleanup() cil managed
  {
    // Code size       29 (0x1d)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance native int SherpaOnnx.OfflineStream::get_Handle()
    IL_0006:  call       void SherpaOnnx.OfflineStream::DestroyOfflineStream(native int)
    IL_000b:  ldarg.0
    IL_000c:  ldarg.0
    IL_000d:  ldsfld     native int [netstandard]System.IntPtr::Zero
    IL_0012:  newobj     instance void [netstandard]System.Runtime.InteropServices.HandleRef::.ctor(object,
                                                                                                    native int)
    IL_0017:  stfld      valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineStream::_handle
    IL_001c:  ret
  } // end of method OfflineStream::Cleanup

  .method public hidebysig specialname instance native int 
          get_Handle() cil managed
  {
    // Code size       12 (0xc)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineStream::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  ret
  } // end of method OfflineStream::get_Handle

  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          void  DestroyOfflineStream(native int handle) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" as "AcceptWaveformOffline" winapi) 
          void  AcceptWaveform(native int handle,
                               int32 sampleRate,
                               float32[] samples,
                               int32 n) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" as "GetOfflineStreamResult" winapi) 
          native int  GetResult(native int handle) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" as "DestroyOfflineRecognizerResult" winapi) 
          void  DestroyResult(native int handle) cil managed preservesig
  {
  }
  .property instance class SherpaOnnx.OfflineRecognizerResult
          Result()
  {
    .get instance class SherpaOnnx.OfflineRecognizerResult SherpaOnnx.OfflineStream::get_Result()
  } // end of property OfflineStream::Result
  .property instance native int Handle()
  {
    .get instance native int SherpaOnnx.OfflineStream::get_Handle()
  } // end of property OfflineStream::Handle
} // end of class SherpaOnnx.OfflineStream

.class public auto ansi beforefieldinit SherpaOnnx.OfflineRecognizer
       extends [netstandard]System.Object
       implements [netstandard]System.IDisposable
{
  .class auto ansi serializable sealed nested private beforefieldinit '<>c'
         extends [netstandard]System.Object
  {
    .custom instance void [netstandard]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static initonly class SherpaOnnx.OfflineRecognizer/'<>c' '<>9'
    .field public static class [netstandard]System.Func`2<class SherpaOnnx.OfflineStream,native int> '<>9__3_0'
    .method private hidebysig specialname rtspecialname static 
            void  .cctor() cil managed
    {
      // Code size       11 (0xb)
      .maxstack  8
      IL_0000:  newobj     instance void SherpaOnnx.OfflineRecognizer/'<>c'::.ctor()
      IL_0005:  stsfld     class SherpaOnnx.OfflineRecognizer/'<>c' SherpaOnnx.OfflineRecognizer/'<>c'::'<>9'
      IL_000a:  ret
    } // end of method '<>c'::.cctor

    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  call       instance void [netstandard]System.Object::.ctor()
      IL_0006:  ret
    } // end of method '<>c'::.ctor

    .method assembly hidebysig instance native int 
            '<Decode>b__3_0'(class SherpaOnnx.OfflineStream s) cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance native int SherpaOnnx.OfflineStream::get_Handle()
      IL_0006:  ret
    } // end of method '<>c'::'<Decode>b__3_0'

  } // end of class '<>c'

  .field private valuetype [netstandard]System.Runtime.InteropServices.HandleRef _handle
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(valuetype SherpaOnnx.OfflineRecognizerConfig config) cil managed
  {
    // Code size       28 (0x1c)
    .maxstack  3
    .locals init (native int V_0)
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [netstandard]System.Object::.ctor()
    IL_0006:  ldarga.s   config
    IL_0008:  call       native int SherpaOnnx.OfflineRecognizer::CreateOfflineRecognizer(valuetype SherpaOnnx.OfflineRecognizerConfig&)
    IL_000d:  stloc.0
    IL_000e:  ldarg.0
    IL_000f:  ldarg.0
    IL_0010:  ldloc.0
    IL_0011:  newobj     instance void [netstandard]System.Runtime.InteropServices.HandleRef::.ctor(object,
                                                                                                    native int)
    IL_0016:  stfld      valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineRecognizer::_handle
    IL_001b:  ret
  } // end of method OfflineRecognizer::.ctor

  .method public hidebysig instance class SherpaOnnx.OfflineStream 
          CreateStream() cil managed
  {
    // Code size       22 (0x16)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineRecognizer::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  call       native int SherpaOnnx.OfflineRecognizer::CreateOfflineStream(native int)
    IL_0010:  newobj     instance void SherpaOnnx.OfflineStream::.ctor(native int)
    IL_0015:  ret
  } // end of method OfflineRecognizer::CreateStream

  .method public hidebysig instance void 
          Decode(class SherpaOnnx.OfflineStream 'stream') cil managed
  {
    // Code size       23 (0x17)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineRecognizer::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  ldarg.1
    IL_000c:  callvirt   instance native int SherpaOnnx.OfflineStream::get_Handle()
    IL_0011:  call       void SherpaOnnx.OfflineRecognizer::Decode(native int,
                                                                   native int)
    IL_0016:  ret
  } // end of method OfflineRecognizer::Decode

  .method public hidebysig instance void 
          Decode(class [netstandard]System.Collections.Generic.IEnumerable`1<class SherpaOnnx.OfflineStream> streams) cil managed
  {
    // Code size       64 (0x40)
    .maxstack  3
    .locals init (native int[] V_0)
    IL_0000:  ldarg.1
    IL_0001:  ldsfld     class [netstandard]System.Func`2<class SherpaOnnx.OfflineStream,native int> SherpaOnnx.OfflineRecognizer/'<>c'::'<>9__3_0'
    IL_0006:  dup
    IL_0007:  brtrue.s   IL_0020

    IL_0009:  pop
    IL_000a:  ldsfld     class SherpaOnnx.OfflineRecognizer/'<>c' SherpaOnnx.OfflineRecognizer/'<>c'::'<>9'
    IL_000f:  ldftn      instance native int SherpaOnnx.OfflineRecognizer/'<>c'::'<Decode>b__3_0'(class SherpaOnnx.OfflineStream)
    IL_0015:  newobj     instance void class [netstandard]System.Func`2<class SherpaOnnx.OfflineStream,native int>::.ctor(object,
                                                                                                                          native int)
    IL_001a:  dup
    IL_001b:  stsfld     class [netstandard]System.Func`2<class SherpaOnnx.OfflineStream,native int> SherpaOnnx.OfflineRecognizer/'<>c'::'<>9__3_0'
    IL_0020:  call       class [netstandard]System.Collections.Generic.IEnumerable`1<!!1> [netstandard]System.Linq.Enumerable::Select<class SherpaOnnx.OfflineStream,native int>(class [netstandard]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                                                                                                 class [netstandard]System.Func`2<!!0,!!1>)
    IL_0025:  call       !!0[] [netstandard]System.Linq.Enumerable::ToArray<native int>(class [netstandard]System.Collections.Generic.IEnumerable`1<!!0>)
    IL_002a:  stloc.0
    IL_002b:  ldarg.0
    IL_002c:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineRecognizer::_handle
    IL_0031:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_0036:  ldloc.0
    IL_0037:  ldloc.0
    IL_0038:  ldlen
    IL_0039:  conv.i4
    IL_003a:  call       void SherpaOnnx.OfflineRecognizer::Decode(native int,
                                                                   native int[],
                                                                   int32)
    IL_003f:  ret
  } // end of method OfflineRecognizer::Decode

  .method public hidebysig newslot virtual final 
          instance void  Dispose() cil managed
  {
    // Code size       13 (0xd)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void SherpaOnnx.OfflineRecognizer::Cleanup()
    IL_0006:  ldarg.0
    IL_0007:  call       void [netstandard]System.GC::SuppressFinalize(object)
    IL_000c:  ret
  } // end of method OfflineRecognizer::Dispose

  .method family hidebysig virtual instance void 
          Finalize() cil managed
  {
    .override [netstandard]System.Object::Finalize
    // Code size       16 (0x10)
    .maxstack  1
    .try
    {
      IL_0000:  ldarg.0
      IL_0001:  call       instance void SherpaOnnx.OfflineRecognizer::Cleanup()
      IL_0006:  leave.s    IL_000f

    }  // end .try
    finally
    {
      IL_0008:  ldarg.0
      IL_0009:  call       instance void [netstandard]System.Object::Finalize()
      IL_000e:  endfinally
    }  // end handler
    IL_000f:  ret
  } // end of method OfflineRecognizer::Finalize

  .method private hidebysig instance void 
          Cleanup() cil managed
  {
    // Code size       34 (0x22)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineRecognizer::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  call       void SherpaOnnx.OfflineRecognizer::DestroyOfflineRecognizer(native int)
    IL_0010:  ldarg.0
    IL_0011:  ldarg.0
    IL_0012:  ldsfld     native int [netstandard]System.IntPtr::Zero
    IL_0017:  newobj     instance void [netstandard]System.Runtime.InteropServices.HandleRef::.ctor(object,
                                                                                                    native int)
    IL_001c:  stfld      valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OfflineRecognizer::_handle
    IL_0021:  ret
  } // end of method OfflineRecognizer::Cleanup

  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          native int  CreateOfflineRecognizer(valuetype SherpaOnnx.OfflineRecognizerConfig& config) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          void  DestroyOfflineRecognizer(native int handle) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          native int  CreateOfflineStream(native int handle) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" as "DecodeOfflineStream" winapi) 
          void  Decode(native int handle,
                       native int 'stream') cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" as "DecodeMultipleOfflineStreams" winapi) 
          void  Decode(native int handle,
                       native int[] streams,
                       int32 n) cil managed preservesig
  {
  }
} // end of class SherpaOnnx.OfflineRecognizer

.class private abstract auto ansi sealed beforefieldinit SherpaOnnx.Dll
       extends [netstandard]System.Object
{
  .field public static literal string Filename = "sherpa-onnx-c-api"
} // end of class SherpaOnnx.Dll

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OnlineTransducerModelConfig
       extends [netstandard]System.ValueType
{
  .field public  marshal( lpstr) string Encoder
  .field public  marshal( lpstr) string Decoder
  .field public  marshal( lpstr) string Joiner
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       34 (0x22)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldstr      ""
    IL_0006:  stfld      string SherpaOnnx.OnlineTransducerModelConfig::Encoder
    IL_000b:  ldarg.0
    IL_000c:  ldstr      ""
    IL_0011:  stfld      string SherpaOnnx.OnlineTransducerModelConfig::Decoder
    IL_0016:  ldarg.0
    IL_0017:  ldstr      ""
    IL_001c:  stfld      string SherpaOnnx.OnlineTransducerModelConfig::Joiner
    IL_0021:  ret
  } // end of method OnlineTransducerModelConfig::.ctor

} // end of class SherpaOnnx.OnlineTransducerModelConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OnlineParaformerModelConfig
       extends [netstandard]System.ValueType
{
  .field public  marshal( lpstr) string Encoder
  .field public  marshal( lpstr) string Decoder
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       23 (0x17)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldstr      ""
    IL_0006:  stfld      string SherpaOnnx.OnlineParaformerModelConfig::Encoder
    IL_000b:  ldarg.0
    IL_000c:  ldstr      ""
    IL_0011:  stfld      string SherpaOnnx.OnlineParaformerModelConfig::Decoder
    IL_0016:  ret
  } // end of method OnlineParaformerModelConfig::.ctor

} // end of class SherpaOnnx.OnlineParaformerModelConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OnlineZipformer2CtcModelConfig
       extends [netstandard]System.ValueType
{
  .field public  marshal( lpstr) string Model
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       12 (0xc)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldstr      ""
    IL_0006:  stfld      string SherpaOnnx.OnlineZipformer2CtcModelConfig::Model
    IL_000b:  ret
  } // end of method OnlineZipformer2CtcModelConfig::.ctor

} // end of class SherpaOnnx.OnlineZipformer2CtcModelConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OnlineModelConfig
       extends [netstandard]System.ValueType
{
  .field public valuetype SherpaOnnx.OnlineTransducerModelConfig Transducer
  .field public valuetype SherpaOnnx.OnlineParaformerModelConfig Paraformer
  .field public valuetype SherpaOnnx.OnlineZipformer2CtcModelConfig Zipformer2Ctc
  .field public  marshal( lpstr) string Tokens
  .field public int32 NumThreads
  .field public  marshal( lpstr) string Provider
  .field public int32 Debug
  .field public  marshal( lpstr) string ModelType
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       81 (0x51)
    .maxstack  2
    IL_0000:  ldarg.0
    IL_0001:  newobj     instance void SherpaOnnx.OnlineTransducerModelConfig::.ctor()
    IL_0006:  stfld      valuetype SherpaOnnx.OnlineTransducerModelConfig SherpaOnnx.OnlineModelConfig::Transducer
    IL_000b:  ldarg.0
    IL_000c:  newobj     instance void SherpaOnnx.OnlineParaformerModelConfig::.ctor()
    IL_0011:  stfld      valuetype SherpaOnnx.OnlineParaformerModelConfig SherpaOnnx.OnlineModelConfig::Paraformer
    IL_0016:  ldarg.0
    IL_0017:  newobj     instance void SherpaOnnx.OnlineZipformer2CtcModelConfig::.ctor()
    IL_001c:  stfld      valuetype SherpaOnnx.OnlineZipformer2CtcModelConfig SherpaOnnx.OnlineModelConfig::Zipformer2Ctc
    IL_0021:  ldarg.0
    IL_0022:  ldstr      ""
    IL_0027:  stfld      string SherpaOnnx.OnlineModelConfig::Tokens
    IL_002c:  ldarg.0
    IL_002d:  ldc.i4.1
    IL_002e:  stfld      int32 SherpaOnnx.OnlineModelConfig::NumThreads
    IL_0033:  ldarg.0
    IL_0034:  ldstr      "cpu"
    IL_0039:  stfld      string SherpaOnnx.OnlineModelConfig::Provider
    IL_003e:  ldarg.0
    IL_003f:  ldc.i4.0
    IL_0040:  stfld      int32 SherpaOnnx.OnlineModelConfig::Debug
    IL_0045:  ldarg.0
    IL_0046:  ldstr      ""
    IL_004b:  stfld      string SherpaOnnx.OnlineModelConfig::ModelType
    IL_0050:  ret
  } // end of method OnlineModelConfig::.ctor

} // end of class SherpaOnnx.OnlineModelConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.FeatureConfig
       extends [netstandard]System.ValueType
{
  .field public int32 SampleRate
  .field public int32 FeatureDim
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       20 (0x14)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldc.i4     0x3e80
    IL_0006:  stfld      int32 SherpaOnnx.FeatureConfig::SampleRate
    IL_000b:  ldarg.0
    IL_000c:  ldc.i4.s   80
    IL_000e:  stfld      int32 SherpaOnnx.FeatureConfig::FeatureDim
    IL_0013:  ret
  } // end of method FeatureConfig::.ctor

} // end of class SherpaOnnx.FeatureConfig

.class public sequential ansi sealed beforefieldinit SherpaOnnx.OnlineRecognizerConfig
       extends [netstandard]System.ValueType
{
  .field public valuetype SherpaOnnx.FeatureConfig FeatConfig
  .field public valuetype SherpaOnnx.OnlineModelConfig ModelConfig
  .field public  marshal( lpstr) string DecodingMethod
  .field public int32 MaxActivePaths
  .field public int32 EnableEndpoint
  .field public float32 Rule1MinTrailingSilence
  .field public float32 Rule2MinTrailingSilence
  .field public float32 Rule3MinUtteranceLength
  .field public  marshal( lpstr) string HotwordsFile
  .field public float32 HotwordsScore
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       103 (0x67)
    .maxstack  2
    IL_0000:  ldarg.0
    IL_0001:  newobj     instance void SherpaOnnx.FeatureConfig::.ctor()
    IL_0006:  stfld      valuetype SherpaOnnx.FeatureConfig SherpaOnnx.OnlineRecognizerConfig::FeatConfig
    IL_000b:  ldarg.0
    IL_000c:  newobj     instance void SherpaOnnx.OnlineModelConfig::.ctor()
    IL_0011:  stfld      valuetype SherpaOnnx.OnlineModelConfig SherpaOnnx.OnlineRecognizerConfig::ModelConfig
    IL_0016:  ldarg.0
    IL_0017:  ldstr      "greedy_search"
    IL_001c:  stfld      string SherpaOnnx.OnlineRecognizerConfig::DecodingMethod
    IL_0021:  ldarg.0
    IL_0022:  ldc.i4.4
    IL_0023:  stfld      int32 SherpaOnnx.OnlineRecognizerConfig::MaxActivePaths
    IL_0028:  ldarg.0
    IL_0029:  ldc.i4.0
    IL_002a:  stfld      int32 SherpaOnnx.OnlineRecognizerConfig::EnableEndpoint
    IL_002f:  ldarg.0
    IL_0030:  ldc.r4     1.2
    IL_0035:  stfld      float32 SherpaOnnx.OnlineRecognizerConfig::Rule1MinTrailingSilence
    IL_003a:  ldarg.0
    IL_003b:  ldc.r4     2.4000001
    IL_0040:  stfld      float32 SherpaOnnx.OnlineRecognizerConfig::Rule2MinTrailingSilence
    IL_0045:  ldarg.0
    IL_0046:  ldc.r4     20.
    IL_004b:  stfld      float32 SherpaOnnx.OnlineRecognizerConfig::Rule3MinUtteranceLength
    IL_0050:  ldarg.0
    IL_0051:  ldstr      ""
    IL_0056:  stfld      string SherpaOnnx.OnlineRecognizerConfig::HotwordsFile
    IL_005b:  ldarg.0
    IL_005c:  ldc.r4     2.
    IL_0061:  stfld      float32 SherpaOnnx.OnlineRecognizerConfig::HotwordsScore
    IL_0066:  ret
  } // end of method OnlineRecognizerConfig::.ctor

} // end of class SherpaOnnx.OnlineRecognizerConfig

.class public auto ansi beforefieldinit SherpaOnnx.OnlineRecognizerResult
       extends [netstandard]System.Object
{
  .class sequential ansi sealed nested private beforefieldinit Impl
         extends [netstandard]System.ValueType
  {
    .field public native int Text
  } // end of class Impl

  .field private string _text
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(native int handle) cil managed
  {
    // Code size       109 (0x6d)
    .maxstack  4
    .locals init (valuetype SherpaOnnx.OnlineRecognizerResult/Impl V_0,
             int32 V_1,
             uint8[] V_2,
             uint8* V_3)
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [netstandard]System.Object::.ctor()
    IL_0006:  ldarg.1
    IL_0007:  ldtoken    SherpaOnnx.OnlineRecognizerResult/Impl
    IL_000c:  call       class [netstandard]System.Type [netstandard]System.Type::GetTypeFromHandle(valuetype [netstandard]System.RuntimeTypeHandle)
    IL_0011:  call       object [netstandard]System.Runtime.InteropServices.Marshal::PtrToStructure(native int,
                                                                                                    class [netstandard]System.Type)
    IL_0016:  unbox.any  SherpaOnnx.OnlineRecognizerResult/Impl
    IL_001b:  stloc.0
    IL_001c:  ldc.i4.0
    IL_001d:  stloc.1
    IL_001e:  ldloc.0
    IL_001f:  ldfld      native int SherpaOnnx.OnlineRecognizerResult/Impl::Text
    IL_0024:  call       void* [netstandard]System.IntPtr::op_Explicit(native int)
    IL_0029:  stloc.3
    IL_002a:  br.s       IL_0030

    IL_002c:  ldloc.3
    IL_002d:  ldc.i4.1
    IL_002e:  add
    IL_002f:  stloc.3
    IL_0030:  ldloc.3
    IL_0031:  ldind.u1
    IL_0032:  brtrue.s   IL_002c

    IL_0034:  ldloc.3
    IL_0035:  ldloc.0
    IL_0036:  ldfld      native int SherpaOnnx.OnlineRecognizerResult/Impl::Text
    IL_003b:  call       void* [netstandard]System.IntPtr::op_Explicit(native int)
    IL_0040:  sub
    IL_0041:  ldc.i4.1
    IL_0042:  div
    IL_0043:  conv.i8
    IL_0044:  conv.i4
    IL_0045:  stloc.1
    IL_0046:  ldloc.1
    IL_0047:  newarr     [netstandard]System.Byte
    IL_004c:  stloc.2
    IL_004d:  ldloc.0
    IL_004e:  ldfld      native int SherpaOnnx.OnlineRecognizerResult/Impl::Text
    IL_0053:  ldloc.2
    IL_0054:  ldc.i4.0
    IL_0055:  ldloc.1
    IL_0056:  call       void [netstandard]System.Runtime.InteropServices.Marshal::Copy(native int,
                                                                                        uint8[],
                                                                                        int32,
                                                                                        int32)
    IL_005b:  ldarg.0
    IL_005c:  call       class [netstandard]System.Text.Encoding [netstandard]System.Text.Encoding::get_UTF8()
    IL_0061:  ldloc.2
    IL_0062:  callvirt   instance string [netstandard]System.Text.Encoding::GetString(uint8[])
    IL_0067:  stfld      string SherpaOnnx.OnlineRecognizerResult::_text
    IL_006c:  ret
  } // end of method OnlineRecognizerResult::.ctor

  .method public hidebysig specialname instance string 
          get_Text() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldfld      string SherpaOnnx.OnlineRecognizerResult::_text
    IL_0006:  ret
  } // end of method OnlineRecognizerResult::get_Text

  .property instance string Text()
  {
    .get instance string SherpaOnnx.OnlineRecognizerResult::get_Text()
  } // end of property OnlineRecognizerResult::Text
} // end of class SherpaOnnx.OnlineRecognizerResult

.class public auto ansi beforefieldinit SherpaOnnx.OnlineStream
       extends [netstandard]System.Object
       implements [netstandard]System.IDisposable
{
  .field private valuetype [netstandard]System.Runtime.InteropServices.HandleRef _handle
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(native int p) cil managed
  {
    // Code size       20 (0x14)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [netstandard]System.Object::.ctor()
    IL_0006:  ldarg.0
    IL_0007:  ldarg.0
    IL_0008:  ldarg.1
    IL_0009:  newobj     instance void [netstandard]System.Runtime.InteropServices.HandleRef::.ctor(object,
                                                                                                    native int)
    IL_000e:  stfld      valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OnlineStream::_handle
    IL_0013:  ret
  } // end of method OnlineStream::.ctor

  .method public hidebysig instance void 
          AcceptWaveform(int32 sampleRate,
                         float32[] samples) cil managed
  {
    // Code size       17 (0x11)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance native int SherpaOnnx.OnlineStream::get_Handle()
    IL_0006:  ldarg.1
    IL_0007:  ldarg.2
    IL_0008:  ldarg.2
    IL_0009:  ldlen
    IL_000a:  conv.i4
    IL_000b:  call       void SherpaOnnx.OnlineStream::AcceptWaveform(native int,
                                                                      int32,
                                                                      float32[],
                                                                      int32)
    IL_0010:  ret
  } // end of method OnlineStream::AcceptWaveform

  .method public hidebysig instance void 
          InputFinished() cil managed
  {
    // Code size       12 (0xc)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance native int SherpaOnnx.OnlineStream::get_Handle()
    IL_0006:  call       void SherpaOnnx.OnlineStream::InputFinished(native int)
    IL_000b:  ret
  } // end of method OnlineStream::InputFinished

  .method family hidebysig virtual instance void 
          Finalize() cil managed
  {
    .override [netstandard]System.Object::Finalize
    // Code size       16 (0x10)
    .maxstack  1
    .try
    {
      IL_0000:  ldarg.0
      IL_0001:  call       instance void SherpaOnnx.OnlineStream::Cleanup()
      IL_0006:  leave.s    IL_000f

    }  // end .try
    finally
    {
      IL_0008:  ldarg.0
      IL_0009:  call       instance void [netstandard]System.Object::Finalize()
      IL_000e:  endfinally
    }  // end handler
    IL_000f:  ret
  } // end of method OnlineStream::Finalize

  .method public hidebysig newslot virtual final 
          instance void  Dispose() cil managed
  {
    // Code size       13 (0xd)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void SherpaOnnx.OnlineStream::Cleanup()
    IL_0006:  ldarg.0
    IL_0007:  call       void [netstandard]System.GC::SuppressFinalize(object)
    IL_000c:  ret
  } // end of method OnlineStream::Dispose

  .method private hidebysig instance void 
          Cleanup() cil managed
  {
    // Code size       29 (0x1d)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance native int SherpaOnnx.OnlineStream::get_Handle()
    IL_0006:  call       void SherpaOnnx.OnlineStream::DestroyOnlineStream(native int)
    IL_000b:  ldarg.0
    IL_000c:  ldarg.0
    IL_000d:  ldsfld     native int [netstandard]System.IntPtr::Zero
    IL_0012:  newobj     instance void [netstandard]System.Runtime.InteropServices.HandleRef::.ctor(object,
                                                                                                    native int)
    IL_0017:  stfld      valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OnlineStream::_handle
    IL_001c:  ret
  } // end of method OnlineStream::Cleanup

  .method public hidebysig specialname instance native int 
          get_Handle() cil managed
  {
    // Code size       12 (0xc)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OnlineStream::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  ret
  } // end of method OnlineStream::get_Handle

  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          void  DestroyOnlineStream(native int handle) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          void  AcceptWaveform(native int handle,
                               int32 sampleRate,
                               float32[] samples,
                               int32 n) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          void  InputFinished(native int handle) cil managed preservesig
  {
  }
  .property instance native int Handle()
  {
    .get instance native int SherpaOnnx.OnlineStream::get_Handle()
  } // end of property OnlineStream::Handle
} // end of class SherpaOnnx.OnlineStream

.class public auto ansi beforefieldinit SherpaOnnx.OnlineRecognizer
       extends [netstandard]System.Object
       implements [netstandard]System.IDisposable
{
  .class auto ansi serializable sealed nested private beforefieldinit '<>c'
         extends [netstandard]System.Object
  {
    .custom instance void [netstandard]System.Runtime.CompilerServices.CompilerGeneratedAttribute::.ctor() = ( 01 00 00 00 ) 
    .field public static initonly class SherpaOnnx.OnlineRecognizer/'<>c' '<>9'
    .field public static class [netstandard]System.Func`2<class SherpaOnnx.OnlineStream,native int> '<>9__5_0'
    .method private hidebysig specialname rtspecialname static 
            void  .cctor() cil managed
    {
      // Code size       11 (0xb)
      .maxstack  8
      IL_0000:  newobj     instance void SherpaOnnx.OnlineRecognizer/'<>c'::.ctor()
      IL_0005:  stsfld     class SherpaOnnx.OnlineRecognizer/'<>c' SherpaOnnx.OnlineRecognizer/'<>c'::'<>9'
      IL_000a:  ret
    } // end of method '<>c'::.cctor

    .method public hidebysig specialname rtspecialname 
            instance void  .ctor() cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.0
      IL_0001:  call       instance void [netstandard]System.Object::.ctor()
      IL_0006:  ret
    } // end of method '<>c'::.ctor

    .method assembly hidebysig instance native int 
            '<Decode>b__5_0'(class SherpaOnnx.OnlineStream s) cil managed
    {
      // Code size       7 (0x7)
      .maxstack  8
      IL_0000:  ldarg.1
      IL_0001:  callvirt   instance native int SherpaOnnx.OnlineStream::get_Handle()
      IL_0006:  ret
    } // end of method '<>c'::'<Decode>b__5_0'

  } // end of class '<>c'

  .field private valuetype [netstandard]System.Runtime.InteropServices.HandleRef _handle
  .method public hidebysig specialname rtspecialname 
          instance void  .ctor(valuetype SherpaOnnx.OnlineRecognizerConfig config) cil managed
  {
    // Code size       28 (0x1c)
    .maxstack  3
    .locals init (native int V_0)
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [netstandard]System.Object::.ctor()
    IL_0006:  ldarga.s   config
    IL_0008:  call       native int SherpaOnnx.OnlineRecognizer::CreateOnlineRecognizer(valuetype SherpaOnnx.OnlineRecognizerConfig&)
    IL_000d:  stloc.0
    IL_000e:  ldarg.0
    IL_000f:  ldarg.0
    IL_0010:  ldloc.0
    IL_0011:  newobj     instance void [netstandard]System.Runtime.InteropServices.HandleRef::.ctor(object,
                                                                                                    native int)
    IL_0016:  stfld      valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OnlineRecognizer::_handle
    IL_001b:  ret
  } // end of method OnlineRecognizer::.ctor

  .method public hidebysig instance class SherpaOnnx.OnlineStream 
          CreateStream() cil managed
  {
    // Code size       22 (0x16)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OnlineRecognizer::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  call       native int SherpaOnnx.OnlineRecognizer::CreateOnlineStream(native int)
    IL_0010:  newobj     instance void SherpaOnnx.OnlineStream::.ctor(native int)
    IL_0015:  ret
  } // end of method OnlineRecognizer::CreateStream

  .method public hidebysig instance bool 
          IsReady(class SherpaOnnx.OnlineStream 'stream') cil managed
  {
    // Code size       26 (0x1a)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OnlineRecognizer::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  ldarg.1
    IL_000c:  callvirt   instance native int SherpaOnnx.OnlineStream::get_Handle()
    IL_0011:  call       int32 SherpaOnnx.OnlineRecognizer::IsReady(native int,
                                                                    native int)
    IL_0016:  ldc.i4.0
    IL_0017:  cgt.un
    IL_0019:  ret
  } // end of method OnlineRecognizer::IsReady

  .method public hidebysig instance bool 
          IsEndpoint(class SherpaOnnx.OnlineStream 'stream') cil managed
  {
    // Code size       26 (0x1a)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OnlineRecognizer::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  ldarg.1
    IL_000c:  callvirt   instance native int SherpaOnnx.OnlineStream::get_Handle()
    IL_0011:  call       int32 SherpaOnnx.OnlineRecognizer::IsEndpoint(native int,
                                                                       native int)
    IL_0016:  ldc.i4.0
    IL_0017:  cgt.un
    IL_0019:  ret
  } // end of method OnlineRecognizer::IsEndpoint

  .method public hidebysig instance void 
          Decode(class SherpaOnnx.OnlineStream 'stream') cil managed
  {
    // Code size       23 (0x17)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OnlineRecognizer::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  ldarg.1
    IL_000c:  callvirt   instance native int SherpaOnnx.OnlineStream::get_Handle()
    IL_0011:  call       void SherpaOnnx.OnlineRecognizer::Decode(native int,
                                                                  native int)
    IL_0016:  ret
  } // end of method OnlineRecognizer::Decode

  .method public hidebysig instance void 
          Decode(class [netstandard]System.Collections.Generic.IEnumerable`1<class SherpaOnnx.OnlineStream> streams) cil managed
  {
    // Code size       64 (0x40)
    .maxstack  3
    .locals init (native int[] V_0)
    IL_0000:  ldarg.1
    IL_0001:  ldsfld     class [netstandard]System.Func`2<class SherpaOnnx.OnlineStream,native int> SherpaOnnx.OnlineRecognizer/'<>c'::'<>9__5_0'
    IL_0006:  dup
    IL_0007:  brtrue.s   IL_0020

    IL_0009:  pop
    IL_000a:  ldsfld     class SherpaOnnx.OnlineRecognizer/'<>c' SherpaOnnx.OnlineRecognizer/'<>c'::'<>9'
    IL_000f:  ldftn      instance native int SherpaOnnx.OnlineRecognizer/'<>c'::'<Decode>b__5_0'(class SherpaOnnx.OnlineStream)
    IL_0015:  newobj     instance void class [netstandard]System.Func`2<class SherpaOnnx.OnlineStream,native int>::.ctor(object,
                                                                                                                         native int)
    IL_001a:  dup
    IL_001b:  stsfld     class [netstandard]System.Func`2<class SherpaOnnx.OnlineStream,native int> SherpaOnnx.OnlineRecognizer/'<>c'::'<>9__5_0'
    IL_0020:  call       class [netstandard]System.Collections.Generic.IEnumerable`1<!!1> [netstandard]System.Linq.Enumerable::Select<class SherpaOnnx.OnlineStream,native int>(class [netstandard]System.Collections.Generic.IEnumerable`1<!!0>,
                                                                                                                                                                                class [netstandard]System.Func`2<!!0,!!1>)
    IL_0025:  call       !!0[] [netstandard]System.Linq.Enumerable::ToArray<native int>(class [netstandard]System.Collections.Generic.IEnumerable`1<!!0>)
    IL_002a:  stloc.0
    IL_002b:  ldarg.0
    IL_002c:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OnlineRecognizer::_handle
    IL_0031:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_0036:  ldloc.0
    IL_0037:  ldloc.0
    IL_0038:  ldlen
    IL_0039:  conv.i4
    IL_003a:  call       void SherpaOnnx.OnlineRecognizer::Decode(native int,
                                                                  native int[],
                                                                  int32)
    IL_003f:  ret
  } // end of method OnlineRecognizer::Decode

  .method public hidebysig instance class SherpaOnnx.OnlineRecognizerResult 
          GetResult(class SherpaOnnx.OnlineStream 'stream') cil managed
  {
    // Code size       36 (0x24)
    .maxstack  2
    .locals init (class SherpaOnnx.OnlineRecognizerResult V_0)
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OnlineRecognizer::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  ldarg.1
    IL_000c:  callvirt   instance native int SherpaOnnx.OnlineStream::get_Handle()
    IL_0011:  call       native int SherpaOnnx.OnlineRecognizer::GetResult(native int,
                                                                           native int)
    IL_0016:  dup
    IL_0017:  newobj     instance void SherpaOnnx.OnlineRecognizerResult::.ctor(native int)
    IL_001c:  stloc.0
    IL_001d:  call       void SherpaOnnx.OnlineRecognizer::DestroyResult(native int)
    IL_0022:  ldloc.0
    IL_0023:  ret
  } // end of method OnlineRecognizer::GetResult

  .method public hidebysig instance void 
          Reset(class SherpaOnnx.OnlineStream 'stream') cil managed
  {
    // Code size       23 (0x17)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OnlineRecognizer::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  ldarg.1
    IL_000c:  callvirt   instance native int SherpaOnnx.OnlineStream::get_Handle()
    IL_0011:  call       void SherpaOnnx.OnlineRecognizer::Reset(native int,
                                                                 native int)
    IL_0016:  ret
  } // end of method OnlineRecognizer::Reset

  .method public hidebysig newslot virtual final 
          instance void  Dispose() cil managed
  {
    // Code size       13 (0xd)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void SherpaOnnx.OnlineRecognizer::Cleanup()
    IL_0006:  ldarg.0
    IL_0007:  call       void [netstandard]System.GC::SuppressFinalize(object)
    IL_000c:  ret
  } // end of method OnlineRecognizer::Dispose

  .method family hidebysig virtual instance void 
          Finalize() cil managed
  {
    .override [netstandard]System.Object::Finalize
    // Code size       16 (0x10)
    .maxstack  1
    .try
    {
      IL_0000:  ldarg.0
      IL_0001:  call       instance void SherpaOnnx.OnlineRecognizer::Cleanup()
      IL_0006:  leave.s    IL_000f

    }  // end .try
    finally
    {
      IL_0008:  ldarg.0
      IL_0009:  call       instance void [netstandard]System.Object::Finalize()
      IL_000e:  endfinally
    }  // end handler
    IL_000f:  ret
  } // end of method OnlineRecognizer::Finalize

  .method private hidebysig instance void 
          Cleanup() cil managed
  {
    // Code size       34 (0x22)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  ldflda     valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OnlineRecognizer::_handle
    IL_0006:  call       instance native int [netstandard]System.Runtime.InteropServices.HandleRef::get_Handle()
    IL_000b:  call       void SherpaOnnx.OnlineRecognizer::DestroyOnlineRecognizer(native int)
    IL_0010:  ldarg.0
    IL_0011:  ldarg.0
    IL_0012:  ldsfld     native int [netstandard]System.IntPtr::Zero
    IL_0017:  newobj     instance void [netstandard]System.Runtime.InteropServices.HandleRef::.ctor(object,
                                                                                                    native int)
    IL_001c:  stfld      valuetype [netstandard]System.Runtime.InteropServices.HandleRef SherpaOnnx.OnlineRecognizer::_handle
    IL_0021:  ret
  } // end of method OnlineRecognizer::Cleanup

  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          native int  CreateOnlineRecognizer(valuetype SherpaOnnx.OnlineRecognizerConfig& config) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          void  DestroyOnlineRecognizer(native int handle) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          native int  CreateOnlineStream(native int handle) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" as "IsOnlineStreamReady" winapi) 
          int32  IsReady(native int handle,
                         native int 'stream') cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" as "DecodeOnlineStream" winapi) 
          void  Decode(native int handle,
                       native int 'stream') cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" as "DecodeMultipleOnlineStreams" winapi) 
          void  Decode(native int handle,
                       native int[] streams,
                       int32 n) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" as "GetOnlineStreamResult" winapi) 
          native int  GetResult(native int handle,
                                native int 'stream') cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" as "DestroyOnlineRecognizerResult" winapi) 
          void  DestroyResult(native int result) cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          void  Reset(native int handle,
                      native int 'stream') cil managed preservesig
  {
  }
  .method private hidebysig static pinvokeimpl("sherpa-onnx-c-api" winapi) 
          int32  IsEndpoint(native int handle,
                            native int 'stream') cil managed preservesig
  {
  }
} // end of class SherpaOnnx.OnlineRecognizer


// =============================================================

// *********** DISASSEMBLY COMPLETE ***********************
// WARNING: Created Win32 resource file sherpa-onnx.res
