#pragma once


#pragma warning( disable : 4278 )
#pragma warning( disable : 4146 )

#if defined _FOR_VS2008

//The following #import imports MSO based on it's LIBID
#import "libid:1CBA492E-7263-47BB-87FE-639000619B15" version ("9.0") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids
//The following #import imports DTE based on it's LIBID
#import "libid:80cc9f66-e7d8-4ddd-85b6-d9e6cd0e93e2" version("9.0") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids

#elif defined _FOR_VS2005

//The following #import imports MSO based on it's LIBID
#import "libid:1CBA492E-7263-47BB-87FE-639000619B15" version ("8.0") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids
//The following #import imports DTE based on it's LIBID
#import "libid:80cc9f66-e7d8-4ddd-85b6-d9e6cd0e93e2" version("8.0") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids

#else

//The following #import imports MSO based on it's LIBID
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" version("2.2") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids
//The following #import imports DTE based on it's LIBID
#import "libid:80cc9f66-e7d8-4ddd-85b6-d9e6cd0e93e2" version("7.0") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids

#endif

//Addin designer
#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4" version("7.0") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids

#if defined(_FOR_VS2008)
//The following #import imports DTE80 based on it's LIBID
#import "libid:1A31287A-4D7D-413e-8E32-3B374931BD89" version("9.0") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids
#elif defined(_FOR_VS2005)
//The following #import imports DTE80 based on it's LIBID
#import "libid:1A31287A-4D7D-413e-8E32-3B374931BD89" version("8.0") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids
#else
#endif

//VC Project Model
///#import "libid:233ADBAD-405A-4249-AA0B-828093D57184" version ("7.1") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids
//VC Code Model
#if defined(_FOR_VS2008)
#import "libid:B5D4541F-A1F1-4CE0-B2E7-5DA402367104" version ("9.0") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids
#elif defined(_FOR_VS2005)
#import "libid:4660A126-0E85-4506-B325-115F0FE27B0D" version ("8.0") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids
#else
#import "libid:F6A0C8E9-68E6-4483-A89D-017AFEC485DF" version ("7.1") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids
#endif
//VC Project Engine
//#import "libid:6194E01D-71A1-419F-85E3-47BA6283DD1D" version ("7.1") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids


//res edit

#if defined(_FOR_VS2008)
#import "libid:6B96E914-4E98-44D6-BD67-173694B10F37" version("9.0") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids
#elif defined(_FOR_VS2005)
#import "libid:6B96E914-4E98-44D6-BD67-173694B10F37" version("8.0") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids
#else
#import "libid:7365C6FE-4191-476B-A3FE-1CB6A7B1C119" version("7.1") lcid("0") raw_interfaces_only no_implementation no_smart_pointers named_guids
#endif

#pragma warning( default : 4278 )
#pragma warning( default : 4146 )

