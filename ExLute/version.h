// version.h
#pragma once

// Numerical versions used by the OS for binary comparison (comma-separated)
#define VERSION_MAJOR               2
#define VERSION_MINOR               1
#define VERSION_PATCH               1
#define VERSION_BUILD               1

#define FILE_VERSION_BINARY         VERSION_MAJOR, VERSION_MINOR, VERSION_PATCH, VERSION_BUILD
#define PRODUCT_VERSION_BINARY      VERSION_MAJOR, VERSION_MINOR, VERSION_PATCH, VERSION_BUILD

// String versions displayed in File Explorer Details (quoted strings)
#define STRINGIFY(x) #x
#define TOSTRING(x) STRINGIFY(x)

#define FILE_VERSION_STRING         TOSTRING(VERSION_MAJOR) "." TOSTRING(VERSION_MINOR) "." TOSTRING(VERSION_PATCH) "." TOSTRING(VERSION_BUILD)
#define PRODUCT_VERSION_STRING      FILE_VERSION_STRING

#define COMPANY_NAME_STRING         "DragonSword Modding"
#define FILE_DESCRIPTION_STRING		"Game editor"
#define LEGAL_COPYRIGHT_STRING		"Copyright (C) 2026 Daleth"
#define PRODUCT_NAME_STRING         "Game editor"
#define INTERNAL_NAME_STRING        "ExLute.exe"
