// insert_character_gui.cpp
//
// Win32 GUI front-end for adding/removing tb_character row(s) in a
// SQLCipher-encrypted save file.
//
// The window shows three checkboxes (one per selectable CHARACTER_CID)
// plus "Add Character" and "Remove Character" buttons. At least one
// checkbox must be ticked before either button will do anything; if more
// than one is ticked, the add/remove logic runs once per ticked value,
// all inside a single transaction (all succeed or all roll back).
//
// A second tab, "Equipment", edits the main/sub stats of the accessories
// a character is currently wearing. tb_equip_mount maps CHARACTER_CID ->
// the ITEM_DBID in each ACC_* slot, and tb_equipment holds that piece's
// MAIN_STAT_CID / SUB_STAT_CID1..5. Both are offered as by-name
// dropdowns limited to the stats that piece's star grade and slot can
// actually roll -- see the "Equipment stats" section further down for the
// ID formulas behind that.
//
// Status/error output is shown via MessageBox as a SIMPLE success/fail
// message only � no file paths, directory names, or SQL details are ever
// displayed on screen. Detailed diagnostics are still collected internally
// (in g_log) in case you want to redirect them to a log file for your own
// debugging, but ShowResult() no longer surfaces them to the user.
//
// Save file discovery:
//   1. On startup the tool first tries the original convention silently:
//      the program's current working directory must be named
//      "DragonSword  Awakening" (exactly, including the double space) and
//      ".\DS\Saved\SaveGames" must exist under it.
//   2. If that auto-detect fails, the full menu is NOT drawn. Instead a
//      small picker window is shown asking the user to manually browse to
//      their SaveGames folder. The selected folder must be literally
//      named "SaveGames" (validated) -- anything else is rejected and the
//      user is asked to pick again. Closing the picker without a valid
//      pick exits the app without ever showing the main window.
//   3. Once a SaveGames folder is resolved (auto or manual), it looks for
//      "<SaveGames>\<USER_ID>" where <USER_ID> is a folder whose name is
//      made up entirely of digits. There must be exactly one such folder --
//      if none or more than one is found, it fails.
//   4. It targets the file:
//        <SaveGames>\<USER_ID>\<USER_ID>_Slot1.db
//   5. The backup button still creates a sibling "SaveGames_<timestamp>"
//      folder in the same parent directory as the resolved SaveGames
//      folder, whether that folder was auto-detected or manually picked.
//
// BUILD (Windows, MSVC, with SQLCipher, C++17 for <filesystem>):
//   - Project Properties -> C/C++ -> Language -> C++ Language Standard:
//     ISO C++17 (or later)
//   - Project Properties -> Linker -> System -> SubSystem: Windows
//     (/SUBSYSTEM:WINDOWS), since this is a GUI app (no console window).
//   - Link against sqlcipher.lib (e.g. via vcpkg: vcpkg install sqlcipher:x64-windows)
//   - Command line example with vcpkg:
//       cl /EHsc /std:c++17 insert_character_gui.cpp ^
//          /I%VCPKG_ROOT%\installed\x64-windows\include ^
//          /link /SUBSYSTEM:WINDOWS %VCPKG_ROOT%\installed\x64-windows\lib\sqlcipher.lib user32.lib
//
// NOTE: SQLCipher exposes the same API surface as sqlite3.h (it's a drop-in,
// encrypted fork of SQLite), so all sqlite3_* calls below apply transparently
// once the PRAGMA key is set as the very first operation on the connection.



#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <shobjidl.h> // IFileOpenDialog (modern "Open Folder" picker)
#include <sqlite3.h>
#include <string>
#include <vector>
#include <sstream>
#include <filesystem>
#include <cstdio>

namespace fs = std::filesystem;

static bool g_noDuplicateSubStat = true;

// Pulls in comctl32 v6 (themed controls: flat modern checkboxes, etc.)
// without needing a separate .manifest file next to the .exe.
#pragma comment(linker, \
    "\"/manifestdependency:type='win32' "                        \
    "name='Microsoft.Windows.Common-Controls' version='6.0.0.0' " \
    "processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "ole32.lib")   // COM (IFileOpenDialog), CoTaskMemFree
#pragma comment(lib, "uuid.lib")    // CLSID_FileOpenDialog / IID_IFileOpenDialog

// These two DWM attributes ship in newer SDKs; declared manually so the
// project still builds against older Windows SDKs. Calls using them fail
// harmlessly (return non-S_OK) on Windows versions that don't support them.
#ifndef DWMWA_USE_IMMERSIVE_DARK_MODE
#define DWMWA_USE_IMMERSIVE_DARK_MODE 20
#endif
#ifndef DWMWA_WINDOW_CORNER_PREFERENCE
#define DWMWA_WINDOW_CORNER_PREFERENCE 33
#endif
#ifndef DWMWCP_ROUND
#define DWMWCP_ROUND 2
#endif
extern "C" HRESULT WINAPI DwmSetWindowAttribute(HWND hwnd, DWORD dwAttribute, LPCVOID pvAttribute, DWORD cbAttribute);

// ---- Configuration ----
static const char* SQLCIPHER_KEY = "13314374259236352028";
static const char* EXPECTED_CWD_NAME = "DragonSword  Awakening"; // note: double space, as specified

// ---- Control IDs ----
#define ID_BTN_ADD      1001
#define ID_BTN_REFRESH  1003
#define ID_BTN_BACKUP   1004

// ---- SaveGames picker window (shown instead of the full menu when the
// save folder can't be auto-detected) ----
#define ID_BTN_PICKER_BROWSE 4001
#define ID_BTN_PICKER_EXIT   4002
//#define ID_CHK_10026    1003
//#define ID_CHK_10010    1004
//#define ID_CHK_10035    1005
#define ID_CHK_10001    1006
#define ID_CHK_10020    1007
#define ID_CHK_10006    1008
#define ID_CHK_10032    1009
#define ID_CHK_10026    1010
#define ID_CHK_10011    1011
#define ID_CHK_10003    1012
#define ID_CHK_10029    1013
#define ID_CHK_10016    1014
#define ID_CHK_10036    1015
#define ID_CHK_10007    1016
#define ID_CHK_10037    1017
#define ID_CHK_10002    1018
#define ID_CHK_10008    1019
#define ID_CHK_10014    1020
#define ID_CHK_10004    1021
#define ID_CHK_10005    1022
#define ID_CHK_10030    1023
#define ID_CHK_10028    1024
#define ID_CHK_10022    1025
#define ID_CHK_10035    1026
#define ID_CHK_10010    1027
#define ID_CHK_10009    1028
#define ID_CHK_10013    1029
#define ID_CHK_10034    1030
#define ID_CHK_10027    1031
//#define ID_CHK_10027    1032

#define ID_LIST_ROSTER  1100
#define ID_EDIT_TEAM_NAME 1101

// ---- Edit-character dialog control IDs ----
#define ID_EDIT_LEVEL       2001
#define ID_EDIT_AWAKEN      2002
#define ID_EDIT_SKILL_BASE  2010 // 2010..2016 = Skill 1..7
#define ID_BTN_EDIT_OK      2020
#define ID_BTN_EDIT_CANCEL  2021
#define ID_BTN_EDIT_MAX     2022
#define ID_BTN_EDIT_MIN     2023

// ---- Roster right-click context menu ----
#define ID_CTX_EDIT         3000
#define ID_CTX_DELETE       3001

// ---- Equipment tab ----
#define ID_TAB_MAIN          5000
#define ID_CMB_EQUIP_CHAR    5001
#define ID_LIST_EQUIP        5002
#define ID_BTN_EQUIP_REFRESH 5003

// ---- Edit-equipment dialog control IDs ----
#define ID_CMB_EQ_MAIN       5100
#define ID_CMB_EQ_SUB_BASE   5101 // 5101..5105 = Sub stat 1..5
#define ID_BTN_EQ_OK         5110
#define ID_BTN_EQ_CANCEL     5111

// The three selectable characters, in the order the checkboxes are shown.
struct SelectableCharacter {
    HWND hwndCheckbox = nullptr;
    int controlId;
    long long character_cid;
    const char* label;
};

static SelectableCharacter g_selectable[] = {
    {nullptr, ID_CHK_10001, 10001, "Aileen" },
    {nullptr, ID_CHK_10020, 10020, "Alex" },
    {nullptr, ID_CHK_10006, 10006, "Aria" },
    {nullptr, ID_CHK_10032, 10032, "Astria" },
    {nullptr, ID_CHK_10011, 10011, "kalsion" },
    {nullptr, ID_CHK_10003, 10003, "Castella" },
    {nullptr, ID_CHK_10029, 10029, "Cerese" },
    {nullptr, ID_CHK_10016, 10016, "charlotte" },
    {nullptr, ID_CHK_10036, 10036, "Dana" },
    {nullptr, ID_CHK_10007, 10007, "Johnny" },
    {nullptr, ID_CHK_10037, 10037, "Kalien" },
    {nullptr, ID_CHK_10002, 10002, "Lute" },
    {nullptr, ID_CHK_10008, 10008, "Onette" },
    {nullptr, ID_CHK_10014, 10014, "Othello" },
    {nullptr, ID_CHK_10004, 10004, "Reina" },
    {nullptr, ID_CHK_10005, 10005, "Roxy" },
    {nullptr, ID_CHK_10030, 10030, "Sion" },
    {nullptr, ID_CHK_10028, 10028, "Tarte" },
    {nullptr, ID_CHK_10022, 10022, "Theresia" },

    {nullptr, ID_CHK_10035, 10035, "Viola" },
    {nullptr, ID_CHK_10010, 10010, "Ysera" },
    {nullptr, ID_CHK_10009, 10009, "Ropie" },
    {nullptr, ID_CHK_10013, 10013, "Ewald" },
    {nullptr, ID_CHK_10034, 10034, "Karina" },
    {nullptr, ID_CHK_10026, 10026, "Awaken Lute" },
    //{nullptr, ID_CHK_10027, 10027, "Karma" },
    //{nullptr, ID_CHK_10027, 10027, "Karma_NoMark" }
};
static const int kNumSelectable = 25;

// ---- Modern UI theme ----
// Flat, light "card" look: soft gray-blue window background, white content
// area, blue accent primary button, outlined secondary button. Kept as a
// small palette + a couple of owner-draw helpers rather than pulling in a
// UI framework, so the rest of the file (DB logic) is untouched.
namespace ui {
    static const COLORREF kWindowBg = RGB(0xF3, 0xF5, 0xF8);
    static const COLORREF kCardBg = RGB(0xFF, 0xFF, 0xFF);
    static const COLORREF kTitleText = RGB(0x1B, 0x22, 0x2C);
    static const COLORREF kSubtitleText = RGB(0x6B, 0x74, 0x80);
    static const COLORREF kCheckboxText = RGB(0x2A, 0x30, 0x38);
    static const COLORREF kDivider = RGB(0xE3, 0xE6, 0xEB);

    static const COLORREF kAccent = RGB(0x2F, 0x6F, 0xED);
    static const COLORREF kAccentHover = RGB(0x27, 0x60, 0xD6);
    static const COLORREF kAccentPressed = RGB(0x1D, 0x4E, 0xB0);

    static const COLORREF kSecondaryBorder = RGB(0xD3, 0xD8, 0xDF);
    static const COLORREF kSecondaryHoverFill = RGB(0xF0, 0xF2, 0xF5);
    static const COLORREF kSecondaryPressedFill = RGB(0xE6, 0xE8, 0xEC);
    static const COLORREF kSecondaryText = RGB(0x2A, 0x30, 0x38);
}

// Per-button hover/pressed state, used by the owner-draw painter below.
struct ButtonVisual {
    bool hover = false;
    bool pressed = false;
    bool primary = false; // true = filled accent button, false = outlined
};
static ButtonVisual g_addBtnVisual{ false, false, true };
static ButtonVisual g_refreshBtnVisual{ false, false, false };
static ButtonVisual g_backupBtnVisual{ false, false, false };
static ButtonVisual g_equipRefreshBtnVisual{ false, false, false };

static HWND g_hwndTitle = nullptr;
static HWND g_hwndSubtitle = nullptr;
static HWND g_hwndRosterTitle = nullptr;
static HWND g_hwndRosterList = nullptr;
static HWND g_hwndTeamNameLabel = nullptr;
static HWND g_hwndTeamNameEdit = nullptr;

// ---- Tab strip ----
// The window now has two pages ("Characters" and "Equipment"). Both are
// laid out as ordinary sibling children of the main window and simply
// shown/hidden when the tab changes -- see ShowTabPage() -- rather than
// living in real child dialogs, so every control's WM_COMMAND still goes
// straight to WndProc the way the rest of this file expects.
static HWND g_hwndTabs = nullptr;
static std::vector<HWND> g_tabPageCharacters;
static std::vector<HWND> g_tabPageEquipment;
static int g_activeTab = 0;

// ---- Equipment tab controls ----
static HWND g_hwndEquipTitle = nullptr;
static HWND g_hwndEquipCharLabel = nullptr;
static HWND g_hwndEquipCharCombo = nullptr;
static HWND g_hwndEquipHint = nullptr;
static HWND g_hwndEquipList = nullptr;
static bool g_hasSkillTable = false;
// Whether tb_advent_status was found the last time the team name was
// loaded; the edit box is disabled (and saves are skipped) when false so
// we don't repeatedly try to write to a table/save file that isn't there.
static bool g_hasAdventStatusTable = false;
// Last value successfully loaded from (or saved to) tb_advent_status.NICK_NAME,
// used to skip redundant writes and to know what to fall back to on failure.
static std::string g_lastLoadedTeamName;
static HFONT g_hFontTitle = nullptr;
static HFONT g_hFontRegular = nullptr;
static HFONT g_hFontButton = nullptr;
static HFONT g_hFontSubtitle = nullptr;
static HBRUSH g_hbrWindowBg = nullptr;
static HBRUSH g_hbrCardBg = nullptr;
static UINT g_dpi = 96;

// Scales a design-time (96 DPI) pixel value to the window's actual DPI.
static int Scale(int value) {
    return MulDiv(value, static_cast<int>(g_dpi), 96);
}

// Row values taken from the reference screenshot. character_cid is filled
// in per-checkbox at operation time; everything else stays constant.
struct CharacterRow {
    long long user_dbid = 1000;
    long long character_cid = 10010;
    long long level = 1;
    long long exp = 0;
    long long ascend = 0;
    long long hp = 500;
    long long transcend = 0;
    long long transcend_total = 0;
    long long soldier_grade = 0;
    long long soldier_grade_point = 0;
    long long created_date = 0;
};

// Accumulates diagnostic log lines for a single operation. This is kept
// internal-only: nothing in here (including file paths) is ever shown to
// the user via MessageBox. If you want it, redirect g_log.str() to a file
// yourself rather than displaying it.
static std::ostringstream g_log;

static void Log(const std::string& line) {
    g_log << line << "\n";
}

static bool isAllDigits(const std::string& s) {
    if (s.empty()) return false;
    for (char c : s) {
        if (!std::isdigit(static_cast<unsigned char>(c))) return false;
    }
    return true;
}

// The resolved "SaveGames" folder everything below is rooted at. Empty
// until WinMain fills it in -- either via tryAutoDetectSaveGamesDir()
// (silent, old convention) or via the user manually browsing to it in
// RunSaveGamesPicker(). Once set, it does not change again for the
// lifetime of the process.
static fs::path g_saveGamesDir;

// True if `p` is a real, existing directory literally named "SaveGames".
// Used both to validate a manually-picked folder and, at runtime, as a
// sanity check that the resolved folder is still there.
static bool isValidSaveGamesFolder(const fs::path& p) {
    if (p.empty()) return false;
    std::error_code ec;
    if (!fs::exists(p, ec) || !fs::is_directory(p, ec)) return false;
    return p.filename().string() == "SaveGames";
}

// Old auto-detect rule, tried once (silently) at startup before falling
// back to the manual picker: the exe's current working directory must be
// named "DragonSword  Awakening" (double space) -- i.e. it's sitting in
// the game's install folder -- and ".\DS\Saved\SaveGames" must exist
// under it.
static bool checkWorkingDirectory(bool silent = false) {
    std::error_code ec;
    fs::path cwd = fs::current_path(ec);
    if (ec) {
        Log("Failed to get current directory: " + ec.message());
        return false;
    }

    std::string dirName = cwd.filename().string();
    if (dirName.empty()) {
        dirName = cwd.parent_path().filename().string();
    }

    if (dirName != EXPECTED_CWD_NAME) {
        if (!silent) {
            MessageBoxA(
                NULL,
                "The current directory is not \"DragonSword Awakening\", please run this file to the game installation folder.",
                "ERROR",
                MB_ICONERROR | MB_OK
            );
        }
        //Log("Current directory name mismatch (internal check failed).");
        return false;
    }
    return true;
}

// Tries the original convention silently. On success, fills g_saveGamesDir
// with an absolute path and returns true; on any failure, leaves
// g_saveGamesDir untouched and returns false (no MessageBox -- the caller
// falls back to the manual picker window instead of showing an error).
static bool tryAutoDetectSaveGamesDir() {
    if (!checkWorkingDirectory(/*silent=*/true)) {
        return false;
    }
    fs::path candidate = fs::path("DS") / "Saved" / "SaveGames";
    if (!isValidSaveGamesFolder(candidate)) {
        return false;
    }
    std::error_code ec;
    fs::path abs = fs::absolute(candidate, ec);
    g_saveGamesDir = !ec ? abs : candidate;
    return true;
}

// Guards every runtime operation below against g_saveGamesDir having gone
// missing after startup (folder renamed/deleted/drive unplugged, etc.).
// By the time the main window exists this should always already be valid
// -- WinMain never draws the full menu until it is -- so this is a safety
// net, not the primary discovery path anymore.
static bool ensureSaveGamesDir(bool silent = false) {
    if (isValidSaveGamesFolder(g_saveGamesDir)) {
        return true;
    }
    if (!silent) {
        MessageBoxA(
            NULL,
            "Cannot find your SaveGames folder anymore. Please restart the tool.",
            "ERROR",
            MB_ICONERROR | MB_OK
        );
    }
    return false;
}

static fs::path locateSaveFile(bool silent = false) {
    if (!isValidSaveGamesFolder(g_saveGamesDir)) {
        if (!silent) {
            MessageBoxA(
                NULL,
                "Cannot find your save file, please play the game first",
                "ERROR",
                MB_ICONERROR | MB_OK
            );
        }
        return {};
    }

    std::error_code ec;
    std::vector<fs::path> candidates;
    for (const auto& entry : fs::directory_iterator(g_saveGamesDir, ec)) {
        if (ec) break;
        if (entry.is_directory() && isAllDigits(entry.path().filename().string())) {
            candidates.push_back(entry.path());
        }
    }

    if (candidates.empty()) {
        Log("No numeric USER_ID folder found.");
        return {};
    }
    if (candidates.size() > 1) {
        Log("Multiple numeric USER_ID folders found, expected exactly one.");
        return {};
    }

    std::string userId = candidates[0].filename().string();
    fs::path dbFile = candidates[0] / (userId + "_Slot1.db");

    if (!fs::exists(dbFile, ec)) {
        if (!silent) {
            MessageBoxA(
                NULL,
                "Cannot find your save file, please play the game first",
                "ERROR",
                MB_ICONERROR | MB_OK
            );
        }
        return {};
    }

    return dbFile;
}

// Formats the current local time as DD_MM_YY_HH_MM_SS for backup folder
// names, e.g. "05_03_26_14_07_33".
static std::string FormatBackupTimestamp() {
    SYSTEMTIME st;
    GetLocalTime(&st);
    char buf[32] = {};
    sprintf_s(buf, "%02d_%02d_%02d_%02d_%02d_%02d",
        st.wDay, st.wMonth, st.wYear % 100, st.wHour, st.wMinute, st.wSecond);
    return std::string(buf);
}

// Copies the whole ".\DS\Saved\SaveGames" folder (every USER_ID subfolder
// and every .db file in it) to a sibling folder in the same "Saved"
// directory, named "SaveGames_DD_MM_YY_HH_MM_SS", so the player has a
// timestamped snapshot before making any edits. Returns true and fills
// outBackupPath with the full absolute path of the folder actually
// created on success.
static bool BackupSaveFolder(std::string& outBackupPath) {
    if (!ensureSaveGamesDir()) {
        return false;
    }

    fs::path saveGamesDir = g_saveGamesDir;
    fs::path savedDir = saveGamesDir.parent_path();

    std::error_code ec;
    std::string backupName = "SaveGames_" + FormatBackupTimestamp();
    fs::path backupDir = savedDir / backupName;

    if (fs::exists(backupDir, ec)) {
        // Same-second double click, or a leftover folder with this exact
        // name already; don't silently merge into it.
        Log("Backup destination already exists: " + backupDir.string());
        return false;
    }

    fs::copy(saveGamesDir, backupDir, fs::copy_options::recursive, ec);
    if (ec) {
        Log("Backup copy failed: " + ec.message());
        // Best-effort cleanup of a partial copy so a failed backup doesn't
        // leave a half-written folder behind.
        std::error_code cleanupEc;
        fs::remove_all(backupDir, cleanupEc);
        return false;
    }

    fs::path absBackupDir = fs::absolute(backupDir, ec);
    outBackupPath = !ec ? absBackupDir.string() : backupDir.string();
    return true;
}

static bool execSql(sqlite3* db, const std::string& sql) {
    char* errMsg = nullptr;
    int rc = sqlite3_exec(db, sql.c_str(), nullptr, nullptr, &errMsg);
    if (rc != SQLITE_OK) {
        Log(std::string("SQL error: ") + (errMsg ? errMsg : "unknown"));
        sqlite3_free(errMsg);
        return false;
    }
    return true;
}

static bool verifyKey(sqlite3* db) {
    sqlite3_stmt* stmt = nullptr;
    int rc = sqlite3_prepare_v2(db, "SELECT count(*) FROM sqlite_master;", -1, &stmt, nullptr);
    if (rc != SQLITE_OK) {
        Log(std::string("Key verification failed to prepare: ") + sqlite3_errmsg(db));
        return false;
    }
    rc = sqlite3_step(stmt);
    bool ok = (rc == SQLITE_ROW);
    if (!ok) {
        Log(std::string("Key verification failed: ") + sqlite3_errmsg(db));
    }
    sqlite3_finalize(stmt);
    return ok;
}

static bool tableExists(sqlite3* db, const char* tableName) {
    sqlite3_stmt* stmt = nullptr;
    const char* sql = "SELECT name FROM sqlite_master WHERE type='table' AND name=?;";
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        MessageBoxA(
            NULL,
            "Error DB",
            "ERROR",
            MB_ICONERROR | MB_OK
        );
        return false;
    }
    sqlite3_bind_text(stmt, 1, tableName, -1, SQLITE_STATIC);
    bool found = (sqlite3_step(stmt) == SQLITE_ROW);
    sqlite3_finalize(stmt);
    return found;
}

// Inserts one row. No MessageBox here anymore � per-row success/failure is
// reported by name via the caller's summary box (see RunOperation/ShowResult),
// since several rows can now be processed in one click.
static bool insertCharacter(sqlite3* db, const CharacterRow& row) {
    const char* sql =
        "INSERT INTO tb_character ("
        "USER_DBID, CHARACTER_CID, LEVEL, EXP, ASCEND, HP, "
        "TRANSCEND, TRANSCEND_TOTAL, SOLDIER_GRADE, SOLDIER_GRADE_POINT, CREATED_DATE"
        ") VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);";

    sqlite3_stmt* stmt = nullptr;
    int rc = sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr);
    if (rc != SQLITE_OK) {
        Log(std::string("Failed to prepare insert: ") + sqlite3_errmsg(db));
        return false;
    }

    sqlite3_bind_int64(stmt, 1, row.user_dbid);
    sqlite3_bind_int64(stmt, 2, row.character_cid);
    sqlite3_bind_int64(stmt, 3, row.level);
    sqlite3_bind_int64(stmt, 4, row.exp);
    sqlite3_bind_int64(stmt, 5, row.ascend);
    sqlite3_bind_int64(stmt, 6, row.hp);
    sqlite3_bind_int64(stmt, 7, row.transcend);
    sqlite3_bind_int64(stmt, 8, row.transcend_total);
    sqlite3_bind_int64(stmt, 9, row.soldier_grade);
    sqlite3_bind_int64(stmt, 10, row.soldier_grade_point);
    sqlite3_bind_int64(stmt, 11, row.created_date);

    rc = sqlite3_step(stmt);
    sqlite3_finalize(stmt);

    if (rc != SQLITE_DONE) {
        Log("Insert failed for CHARACTER_CID=" + std::to_string(row.character_cid) +
            " (character already exists?).");
        return false;
    }

    //Log("Inserted 1 row into tb_character (USER_DBID=" + std::to_string(row.user_dbid) +
    //    ", CHARACTER_CID=" + std::to_string(row.character_cid) + ").");
    return true;
}

// Deletes one row. No MessageBox here anymore � see insertCharacter note above.
static bool removeCharacter(sqlite3* db, const CharacterRow& row) {
    const char* sql =
        "DELETE FROM tb_character WHERE USER_DBID = ? AND CHARACTER_CID = ?;";

    sqlite3_stmt* stmt = nullptr;
    int rc = sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr);
    if (rc != SQLITE_OK) {
        Log(std::string("Failed to prepare delete: ") + sqlite3_errmsg(db));
        return false;
    }

    sqlite3_bind_int64(stmt, 1, row.user_dbid);
    sqlite3_bind_int64(stmt, 2, row.character_cid);

    rc = sqlite3_step(stmt);
    sqlite3_finalize(stmt);

    if (rc != SQLITE_DONE) {
        Log(std::string("Delete failed: ") + sqlite3_errmsg(db));
        return false;
    }

    int changes = sqlite3_changes(db);
    if (changes == 0) {
        Log("Character not found for CHARACTER_CID=" + std::to_string(row.character_cid) + ".");
        return false;
    }

    //Log("Removed " + std::to_string(changes) + " row(s) from tb_character (USER_DBID=" +
    //    std::to_string(row.user_dbid) + ", CHARACTER_CID=" +
    //    std::to_string(row.character_cid) + ").");
    return true;
}

// Result of the whole operation, one entry per ticked character, so the
// caller can show a single summary MessageBox instead of one per row.
struct OperationOutcome {
    bool overallSuccess = false;
    // label -> succeeded
    std::vector<std::pair<std::string, bool>> perCharacter;
};

// Runs the full add/remove flow for every ticked character_cid, logging
// every step internally, all inside one transaction (all rows succeed or
// the whole transaction rolls back). Returns per-character results plus
// an overall success flag.
static OperationOutcome RunOperation(bool doAdd, const std::vector<int>& selectedIndices) {
    OperationOutcome outcome;
    g_log.str("");
    g_log.clear();

    if (selectedIndices.empty()) {
        // Caller should have prevented this, but guard anyway.
        return outcome;
    }

    if (!ensureSaveGamesDir()) {
        return outcome;
    }

    fs::path dbPath = locateSaveFile();
    if (dbPath.empty()) {
        return outcome;
    }
    // Path intentionally NOT logged/displayed beyond this internal ostringstream.

    sqlite3* db = nullptr;
    int rc = sqlite3_open(dbPath.string().c_str(), &db);
    if (rc != SQLITE_OK) {
        Log(std::string("Cannot open database: ") + sqlite3_errmsg(db));
        sqlite3_close(db);
        return outcome;
    }

    std::string keyPragma = "PRAGMA key = '" + std::string(SQLCIPHER_KEY) + "';";
    if (!execSql(db, keyPragma)) {
        Log("Failed to apply SQLCipher key.");
        sqlite3_close(db);
        return outcome;
    }

    if (!verifyKey(db)) {
        Log("Database could not be decrypted with the provided key.");
        sqlite3_close(db);
        return outcome;
    }

    if (!tableExists(db, "tb_character")) {
        Log("Table 'tb_character' not found in database.");
        sqlite3_close(db);
        return outcome;
    }

    if (!execSql(db, "BEGIN TRANSACTION;")) {
        sqlite3_close(db);
        return outcome;
    }

    bool allOk = true;
    for (int idx : selectedIndices) {
        CharacterRow row; // defaults already match the screenshot values
        row.character_cid = g_selectable[idx].character_cid;

        bool ok = doAdd ? insertCharacter(db, row) : removeCharacter(db, row);
        outcome.perCharacter.emplace_back(g_selectable[idx].label, ok);
        if (!ok) {
            allOk = false;
        }
    }

    if (!allOk) {
        execSql(db, "ROLLBACK;");
        sqlite3_close(db);
        outcome.overallSuccess = false;
        return outcome;
    }

    if (!execSql(db, "COMMIT;")) {
        sqlite3_close(db);
        outcome.overallSuccess = false;
        return outcome;
    }

    sqlite3_close(db);
    outcome.overallSuccess = true;
    return outcome;
}

// Shows a plain success/failure summary, naming which of the ticked
// characters succeeded or failed � still no paths, directory names, SQL
// text, or sqlite error strings. g_log still holds the detailed trace
// internally if you want to wire it up to a log file for your own
// troubleshooting later.
static void ShowResult(HWND hwnd, bool doAdd, const OperationOutcome& outcome) {
    std::string title = doAdd ? "Add Character" : "Remove Character";

    if (outcome.perCharacter.empty()) {
        MessageBoxA(hwnd, "Operation failed.", title.c_str(), MB_OK | MB_ICONERROR);
        return;
    }

    std::ostringstream body;
    for (const auto& entry : outcome.perCharacter) {
        body << entry.first << ": " << (entry.second ? "OK" : "FAILED") << "\n";
    }
    if (!outcome.overallSuccess) {
        body << "\nNo changes were saved (all changes were rolled back).";
    }

    UINT icon = outcome.overallSuccess ? MB_ICONINFORMATION : MB_ICONERROR;
    MessageBoxA(hwnd, body.str().c_str(), title.c_str(), MB_OK | icon);
}

// Collects which checkboxes are currently ticked, as indices into
// g_selectable. Returns an empty vector if none are ticked.
static std::vector<int> getSelectedIndices() {
    std::vector<int> selected;
    for (int i = 0; i < kNumSelectable; ++i) {
        if (g_selectable[i].hwndCheckbox &&
            SendMessage(g_selectable[i].hwndCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED) {
            selected.push_back(i);
        }
    }
    return selected;
}

// ---- Level -> EXP/ASCEND mapping (from LEVEL.txt) ----
// tb_character doesn't derive EXP/ASCEND from LEVEL on its own, so when the
// user edits a character's level via the roster editor, these are looked
// up and written alongside it. Only levels 3..70 are valid edit targets �
// there's no mapping (and the game caps out at 70) outside that range.
struct LevelMapEntry {
    long long level;
    long long exp;
    long long ascend;
};

static const LevelMapEntry kLevelTable[] = {
    { 3, 500, 1 },
    { 4, 1000, 1 },
    { 5, 2000, 1 },
    { 6, 3000, 1 },
    { 7, 4000, 1 },
    { 8, 6000, 1 },
    { 9, 8000, 1 },
    { 10, 11000, 1 },
    { 11, 14000, 1 },
    { 12, 18000, 1 },
    { 13, 22500, 1 },
    { 14, 28000, 1 },
    { 15, 34000, 1 },
    { 16, 41000, 1 },
    { 17, 48500, 1 },
    { 18, 57000, 1 },
    { 19, 66500, 1 },
    { 20, 77000, 1 },
    { 21, 88500, 1 },
    { 22, 101500, 1 },
    { 23, 115500, 1 },
    { 24, 130500, 1 },
    { 25, 146500, 1 },
    { 26, 164500, 1 },
    { 27, 183500, 1 },
    { 28, 203500, 1 },
    { 29, 225500, 1 },
    { 30, 249000, 1 },
    { 31, 273500, 2 },
    { 32, 300000, 2 },
    { 33, 328500, 2 },
    { 34, 358500, 2 },
    { 35, 390000, 2 },
    { 36, 423500, 2 },
    { 37, 458500, 2 },
    { 38, 496000, 2 },
    { 39, 535000, 2 },
    { 40, 576000, 2 },
    { 41, 619000, 3 },
    { 42, 664500, 3 },
    { 43, 712000, 3 },
    { 44, 761500, 3 },
    { 45, 813500, 3 },
    { 46, 868000, 3 },
    { 47, 924500, 3 },
    { 48, 983500, 3 },
    { 49, 1045000, 3 },
    { 50, 1108500, 3 },
    { 51, 1175000, 4 },
    { 52, 1244500, 4 },
    { 53, 1316000, 4 },
    { 54, 1390500, 4 },
    { 55, 1468000, 4 },
    { 56, 1548000, 4 },
    { 57, 1630500, 4 },
    { 58, 1716500, 4 },
    { 59, 1805000, 4 },
    { 60, 1897000, 4 },
    { 61, 1991500, 5 },
    { 62, 2089500, 5 },
    { 63, 2190500, 5 },
    { 64, 2295000, 5 },
    { 65, 2402500, 5 },
    { 66, 2513500, 5 },
    { 67, 2627500, 5 },
    { 68, 2745000, 5 },
    { 69, 2866000, 5 },
    { 70, 2990500, 5 },
};
static const int kLevelTableCount = sizeof(kLevelTable) / sizeof(kLevelTable[0]);

// Looks up EXP/ASCEND for a given LEVEL. Returns false (and leaves the
// outputs untouched) for any level not in the table � i.e. anything
// outside 3..70, which the edit dialog rejects before this is even called.
static bool LookupLevelMapping(long long level, long long& outExp, long long& outAscend) {
    for (int i = 0; i < kLevelTableCount; ++i) {
        if (kLevelTable[i].level == level) {
            outExp = kLevelTable[i].exp;
            outAscend = kLevelTable[i].ascend;
            return true;
        }
    }
    return false;
}

// ---- Roster list (right-hand panel) ----
// One row of what's currently in tb_character for this save, with the
// CHARACTER_CID already resolved to a display name via g_selectable[].
// skill[i] is the SLOT_LEVEL for TYPE_VALUE == i+1 from tb_skill_growth
// (7 skill slots per character); -1 means "no data" (tb_skill_growth
// itself is missing), as opposed to 0, which means the table exists but
// that particular skill slot has no row / is unlearned.
struct RosterEntry {
    std::string name;
    long long character_cid = 0;
    long long level;
    long long transcend;
    long long skill[7] = { -1, -1, -1, -1, -1, -1, -1 };
};

// Cached copy of what's currently shown in the roster ListView, indexed
// the same way (row i <-> g_currentRoster[i]), so a right-click can look
// up which character/CID that row belongs to without re-querying the DB.
static std::vector<RosterEntry> g_currentRoster;

// Reads the 7 skill levels (TYPE_VALUE 1..7 -> SLOT_LEVEL) for one
// character out of tb_skill_growth. Caller is responsible for having
// already confirmed the table exists; outLevels is left untouched for any
// TYPE_VALUE outside 1..7 or slot that has no row.
static void loadSkillLevels(sqlite3* db, long long cid, long long(&outLevels)[7]) {
    const char* sql = "SELECT TYPE_VALUE, SLOT_LEVEL FROM tb_skill_growth WHERE CHARACTER_CID = ?;";
    sqlite3_stmt* stmt = nullptr;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        Log(std::string("Failed to prepare skill query: ") + sqlite3_errmsg(db));
        return;
    }
    sqlite3_bind_int64(stmt, 1, cid);
    while (sqlite3_step(stmt) == SQLITE_ROW) {
        long long typeValue = sqlite3_column_int64(stmt, 0);
        long long slotLevel = sqlite3_column_int64(stmt, 1);
        if (typeValue >= 1 && typeValue <= 7) {
            outLevels[typeValue - 1] = slotLevel;
        }
    }
    sqlite3_finalize(stmt);
}

static bool queryRoster(sqlite3* db, std::vector<RosterEntry>& outRoster) {
    const char* sql = "SELECT CHARACTER_CID, LEVEL, TRANSCEND FROM tb_character ORDER BY CHARACTER_CID;";
    sqlite3_stmt* stmt = nullptr;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        Log(std::string("Failed to prepare roster query: ") + sqlite3_errmsg(db));
        return false;
    }

    bool hasSkillTable = tableExists(db, "tb_skill_growth");
    g_hasSkillTable = hasSkillTable;

    while (sqlite3_step(stmt) == SQLITE_ROW) {
        long long cid = sqlite3_column_int64(stmt, 0);
        long long level = sqlite3_column_int64(stmt, 1);
        long long transcend = sqlite3_column_int64(stmt, 2);

        std::string name;
        for (int i = 0; i < kNumSelectable; ++i) {
            if (g_selectable[i].character_cid == cid) {
                name = g_selectable[i].label;
                break;
            }
        }
        if (name.empty()) {
            // Not one of the known selectable characters (e.g. a starter
            // character not in g_selectable[]) -- show a neutral label
            // instead of silently dropping the row or exposing the raw CID.
            name = "Unknown character";
        }

        RosterEntry entry;
        entry.name = name;
        entry.character_cid = cid;
        entry.level = level;
        entry.transcend = transcend;
        if (hasSkillTable) {
            // Table exists: default each slot to 0 (exists but unlearned)
            // rather than the -1 "no data" sentinel, then fill in any
            // slots that actually have a row.
            for (int i = 0; i < 7; ++i) entry.skill[i] = 0;
            loadSkillLevels(db, cid, entry.skill);
        }
        outRoster.push_back(entry);
    }
    sqlite3_finalize(stmt);
    return true;
}

// Opens the save file, applies the SQLCipher key, and reads back every row
// currently in tb_character. Deliberately silent on failure (no
// MessageBox): this runs automatically on startup and after every
// Add/Remove, and a wrong game folder / missing save is already reported
// the moment the user actually clicks Add or Remove, so popping the same
// warning again here on every refresh would just be noise.
static bool LoadRoster(std::vector<RosterEntry>& outRoster) {
    if (!ensureSaveGamesDir(/*silent=*/true)) {
        return false;
    }
    fs::path dbPath = locateSaveFile(/*silent=*/true);
    if (dbPath.empty()) {
        return false;
    }

    sqlite3* db = nullptr;
    if (sqlite3_open(dbPath.string().c_str(), &db) != SQLITE_OK) {
        sqlite3_close(db);
        return false;
    }

    std::string keyPragma = "PRAGMA key = '" + std::string(SQLCIPHER_KEY) + "';";
    if (!execSql(db, keyPragma) || !verifyKey(db) || !tableExists(db, "tb_character")) {
        sqlite3_close(db);
        return false;
    }

    bool ok = queryRoster(db, outRoster);
    sqlite3_close(db);
    return ok;
}

// Reads NICK_NAME from tb_advent_status (there's expected to be exactly
// one row; the first one found is used). Deliberately silent on failure,
// same reasoning as LoadRoster: this runs automatically on startup/refresh.
static bool LoadTeamName(std::string& outName) {
    if (!ensureSaveGamesDir(/*silent=*/true)) {
        return false;
    }
    fs::path dbPath = locateSaveFile(/*silent=*/true);
    if (dbPath.empty()) {
        return false;
    }

    sqlite3* db = nullptr;
    if (sqlite3_open(dbPath.string().c_str(), &db) != SQLITE_OK) {
        sqlite3_close(db);
        return false;
    }

    std::string keyPragma = "PRAGMA key = '" + std::string(SQLCIPHER_KEY) + "';";
    if (!execSql(db, keyPragma) || !verifyKey(db) || !tableExists(db, "tb_advent_status")) {
        sqlite3_close(db);
        return false;
    }

    bool ok = false;
    const char* sql = "SELECT NICK_NAME FROM tb_advent_status LIMIT 1;";
    sqlite3_stmt* stmt = nullptr;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) == SQLITE_OK) {
        if (sqlite3_step(stmt) == SQLITE_ROW) {
            const unsigned char* text = sqlite3_column_text(stmt, 0);
            outName = text ? reinterpret_cast<const char*>(text) : "";
            ok = true;
        }
        sqlite3_finalize(stmt);
    }
    else {
        Log(std::string("Failed to prepare team name query: ") + sqlite3_errmsg(db));
    }

    sqlite3_close(db);
    return ok;
}

// Writes a new NICK_NAME back to tb_advent_status. Since the table is
// expected to hold exactly one row, this updates unconditionally (no
// WHERE clause) rather than trying to target a specific rowid.
static bool SaveTeamName(const std::string& newName) {
    if (!ensureSaveGamesDir()) {
        return false;
    }
    fs::path dbPath = locateSaveFile();
    if (dbPath.empty()) {
        return false;
    }

    sqlite3* db = nullptr;
    if (sqlite3_open(dbPath.string().c_str(), &db) != SQLITE_OK) {
        Log(std::string("Cannot open database: ") + sqlite3_errmsg(db));
        sqlite3_close(db);
        return false;
    }

    std::string keyPragma = "PRAGMA key = '" + std::string(SQLCIPHER_KEY) + "';";
    if (!execSql(db, keyPragma) || !verifyKey(db) || !tableExists(db, "tb_advent_status")) {
        sqlite3_close(db);
        return false;
    }

    bool ok = true;
    if (!execSql(db, "BEGIN TRANSACTION;")) {
        sqlite3_close(db);
        return false;
    }

    {
        const char* sql = "UPDATE tb_advent_status SET NICK_NAME=?;";
        sqlite3_stmt* stmt = nullptr;
        if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) == SQLITE_OK) {
            sqlite3_bind_text(stmt, 1, newName.c_str(), -1, SQLITE_TRANSIENT);
            if (sqlite3_step(stmt) != SQLITE_DONE) {
                Log(std::string("Failed to update tb_advent_status: ") + sqlite3_errmsg(db));
                ok = false;
            }
            sqlite3_finalize(stmt);
        }
        else {
            Log(std::string("Failed to prepare tb_advent_status update: ") + sqlite3_errmsg(db));
            ok = false;
        }
    }

    if (ok) {
        ok = execSql(db, "COMMIT;");
    }
    else {
        execSql(db, "ROLLBACK;");
    }

    sqlite3_close(db);
    return ok;
}

// Reloads the team name from the save file into the edit box. Disables
// the control (and blanks it) when tb_advent_status isn't found, so we
// don't let the user type into a field that has nothing to save back to.
static void RefreshTeamNameEdit(HWND hwndEdit) {
    if (!hwndEdit) return;
    std::string name;
    g_hasAdventStatusTable = LoadTeamName(name);
    if (g_hasAdventStatusTable) {
        g_lastLoadedTeamName = name;
        std::wstring wname(name.begin(), name.end());
        SetWindowTextW(hwndEdit, wname.c_str());
        EnableWindow(hwndEdit, TRUE);
    }
    else {
        g_lastLoadedTeamName.clear();
        SetWindowTextW(hwndEdit, L"");
        EnableWindow(hwndEdit, FALSE);
    }
}

// Repopulates the roster ListView from the current save file. Called once
// at startup and again after every Add/Remove so the list always reflects
// what's actually on disk.
static void RefreshRosterList(HWND hwndList) {
    if (!hwndList) return;
    ListView_DeleteAllItems(hwndList);
    g_currentRoster.clear();

    std::vector<RosterEntry> roster;
    if (!LoadRoster(roster)) {
        LVITEMW item = {};
        item.mask = LVIF_TEXT;
        item.iItem = 0;
        item.iSubItem = 0;
        item.pszText = const_cast<LPWSTR>(L"No save file found");
        ListView_InsertItem(hwndList, &item);
        return;
    }

    if (roster.empty()) {
        LVITEMW item = {};
        item.mask = LVIF_TEXT;
        item.iItem = 0;
        item.iSubItem = 0;
        item.pszText = const_cast<LPWSTR>(L"Roster is empty");
        ListView_InsertItem(hwndList, &item);
        return;
    }

    g_currentRoster = roster; // cached for the right-click context menu below

    int row = 0;
    for (const auto& entry : roster) {
        std::wstring wname(entry.name.begin(), entry.name.end());
        std::wstring wlevel = std::to_wstring(entry.level);
        std::wstring wtranscend = std::to_wstring(entry.transcend);

        LVITEMW item = {};
        item.mask = LVIF_TEXT;
        item.iItem = row;
        item.iSubItem = 0;
        item.pszText = const_cast<LPWSTR>(wname.c_str());
        ListView_InsertItem(hwndList, &item);

        ListView_SetItemText(hwndList, row, 1, const_cast<LPWSTR>(wlevel.c_str()));
        ListView_SetItemText(hwndList, row, 2, const_cast<LPWSTR>(wtranscend.c_str()));

        for (int i = 0; i < 7; ++i) {
            std::wstring wskill = (entry.skill[i] < 0) ? L"-" : std::to_wstring(entry.skill[i]);
            ListView_SetItemText(hwndList, row, 3 + i, const_cast<LPWSTR>(wskill.c_str()));
        }

        ++row;
    }
}

// ---- Edit-character dialog ----
// Everything needed to write one character's edits back to the DB: LEVEL
// plus its derived EXP/ASCEND (via kLevelTable), TRANSCEND (Awaken), and
// the 7 skill slot levels.
struct CharacterEditResult {
    long long level = 0;
    long long exp = 0;
    long long ascend = 0;
    long long transcend = 0;
    long long skill[7] = { 0, 0, 0, 0, 0, 0, 0 };
};

// Per-dialog-instance state, stashed in GWLP_USERDATA so EditDialogProc
// can get back to it.
struct EditDialogState {
    HWND hwndLevel = nullptr;
    HWND hwndAwaken = nullptr;
    HWND hwndSkill[7] = {};
    bool hasSkillTable = false;
    CharacterEditResult* outResult = nullptr;
    bool confirmed = false;
};

// Parses an edit control's text as a non-negative integer. Returns false
// (leaving outValue untouched) if the text isn't a valid non-negative
// integer, so the caller can reject with a specific message instead of
// silently treating garbage input as 0.
static bool ParseNonNegativeInt(HWND hwndEdit, long long& outValue) {
    wchar_t buf[32] = {};
    GetWindowTextW(hwndEdit, buf, 32);
    std::wstring s(buf);
    // Trim whitespace.
    size_t start = s.find_first_not_of(L" \t");
    size_t end = s.find_last_not_of(L" \t");
    if (start == std::wstring::npos) return false;
    s = s.substr(start, end - start + 1);
    if (s.empty()) return false;
    for (wchar_t c : s) {
        if (c < L'0' || c > L'9') return false;
    }
    wchar_t* endPtr = nullptr;
    long long value = wcstoll(s.c_str(), &endPtr, 10);
    if (!endPtr || *endPtr != L'\0') return false;
    outValue = value;
    return true;
}

static LRESULT CALLBACK EditDialogProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    EditDialogState* state = reinterpret_cast<EditDialogState*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));

    switch (msg) {
    case WM_CREATE: {
        CREATESTRUCTW* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
        state = reinterpret_cast<EditDialogState*>(cs->lpCreateParams);
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)state);
        return 0;
    }
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)wParam;
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, ui::kCheckboxText);
        return (LRESULT)g_hbrWindowBg;
    }
    case WM_COMMAND: {
        int id = LOWORD(wParam);
        if (id == ID_BTN_EDIT_CANCEL) {
            state->confirmed = false;
            DestroyWindow(hwnd);
            return 0;
        }
        if (id == ID_BTN_EDIT_MAX) {
            SetWindowTextW(state->hwndLevel, L"70");
            SetWindowTextW(state->hwndAwaken, L"6");
            if (state->hasSkillTable) {
                for (int i = 0; i < 7; ++i) {
                    long long maxVal = (i < 4) ? 10 : 7;
                    SetWindowTextW(state->hwndSkill[i], std::to_wstring(maxVal).c_str());
                }
            }
            return 0;
        }
        if (id == ID_BTN_EDIT_MIN) {
            SetWindowTextW(state->hwndLevel, L"1");
            SetWindowTextW(state->hwndAwaken, L"0");
            if (state->hasSkillTable) {
                for (int i = 0; i < 7; ++i) {
                    SetWindowTextW(state->hwndSkill[i], L"0");
                }
            }
            return 0;
        }
        if (id == ID_BTN_EDIT_OK) {
            long long level = 0;
            if (!ParseNonNegativeInt(state->hwndLevel, level)) {
                MessageBoxW(hwnd, L"Level must be a whole number.", L"Invalid level", MB_OK | MB_ICONWARNING);
                return 0;
            }
            long long exp = 0, ascend = 0;
            if (level > 70 || !LookupLevelMapping(level, exp, ascend)) {
                MessageBoxW(hwnd, L"Level must be between 3 and 70.", L"Invalid level", MB_OK | MB_ICONWARNING);
                return 0;
            }

            long long transcend = 0;
            if (!ParseNonNegativeInt(state->hwndAwaken, transcend)) {
                MessageBoxW(hwnd, L"Awaken must be a whole non-negative number.", L"Invalid value", MB_OK | MB_ICONWARNING);
                return 0;
            }

            long long skill[7] = {};
            if (state->hasSkillTable) {
                for (int i = 0; i < 7; ++i) {
                    if (!ParseNonNegativeInt(state->hwndSkill[i], skill[i])) {
                        std::wstring msg = L"Skill " + std::to_wstring(i + 1) + L" must be a whole non-negative number.";
                        MessageBoxW(hwnd, msg.c_str(), L"Invalid value", MB_OK | MB_ICONWARNING);
                        return 0;
                    }
                    // Skill slots 1-4 (TYPE_VALUE 1..4) must be <= 10.
                    // Skill slots 5-7 (TYPE_VALUE 5..7) must be <= 7.
                    long long maxAllowed = (i < 4) ? 10 + 1 : 7 + 1;
                    if (skill[i] >= maxAllowed) {
                        std::wstring msg = L"Skill " + std::to_wstring(i + 1) + L" must be less than " +
                            std::to_wstring(maxAllowed) + L".";
                        MessageBoxW(hwnd, msg.c_str(), L"Invalid value", MB_OK | MB_ICONWARNING);
                        return 0;
                    }
                }
            }

            state->outResult->level = level;
            state->outResult->exp = exp;
            state->outResult->ascend = ascend;
            state->outResult->transcend = transcend;
            if (state->hasSkillTable) {
                for (int i = 0; i < 7; ++i) state->outResult->skill[i] = skill[i];
            }

            state->confirmed = true;
            DestroyWindow(hwnd);
            return 0;
        }
        return 0;
    }
    case WM_CLOSE:
        state->confirmed = false;
        DestroyWindow(hwnd);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// Modal-ish popup (owned by parent, parent disabled while it's open) that
// lets the user edit one character's Level, Awaken, and 7 skill slots.
// Returns true and fills outResult if the user confirmed with valid
// values; false if cancelled/closed without saving.
static bool ShowEditCharacterDialog(HWND parent, const std::string& characterName,
    long long currentLevel, long long currentTranscend,
    const long long currentSkill[7], bool hasSkillTable,
    CharacterEditResult& outResult) {
    static bool classRegistered = false;
    const wchar_t* CLASS_NAME = L"EditCharacterDialogClass";
    if (!classRegistered) {
        WNDCLASSW wc = {};
        wc.lpfnWndProc = EditDialogProc;
        wc.hInstance = GetModuleHandleW(nullptr);
        wc.lpszClassName = CLASS_NAME;
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hbrBackground = g_hbrWindowBg;
        RegisterClassW(&wc);
        classRegistered = true;
    }

    EditDialogState state;
    state.hasSkillTable = hasSkillTable;
    state.outResult = &outResult;

    const int pad = Scale(20);
    const int labelW = Scale(70);
    const int editW = Scale(110);
    const int rowH = Scale(26);
    const int rowGap = Scale(6);

    int rowCount = 2 + (hasSkillTable ? 7 : 0);
    int clientW = pad * 2 + labelW + Scale(8) + editW;
    // Extra row height reserved for the Max/Min quick-fill buttons, above
    // the OK/Cancel row.
    int clientH = pad + rowCount * (rowH + rowGap) + Scale(32 + 12) + Scale(50);

    std::wstring wtitle = L"Edit " + std::wstring(characterName.begin(), characterName.end());

    RECT wr{ 0, 0, clientW, clientH };
    AdjustWindowRectEx(&wr, WS_POPUP | WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_DLGMODALFRAME);
    int winW = wr.right - wr.left;
    int winH = wr.bottom - wr.top;

    RECT parentRect;
    GetWindowRect(parent, &parentRect);
    int posX = parentRect.left + ((parentRect.right - parentRect.left) - winW) / 2;
    int posY = parentRect.top + ((parentRect.bottom - parentRect.top) - winH) / 2;

    EnableWindow(parent, FALSE);

    HWND hwndDlg = CreateWindowExW(WS_EX_DLGMODALFRAME, CLASS_NAME, wtitle.c_str(),
        WS_POPUP | WS_CAPTION | WS_SYSMENU,
        posX, posY, winW, winH,
        parent, nullptr, GetModuleHandleW(nullptr), &state);

    if (!hwndDlg) {
        EnableWindow(parent, TRUE);
        return false;
    }

    int y = pad;
    CreateWindowW(L"STATIC", L"Level:", WS_VISIBLE | WS_CHILD, pad, y, labelW, rowH, hwndDlg, nullptr, nullptr, nullptr);
    state.hwndLevel = CreateWindowW(L"EDIT", std::to_wstring(currentLevel).c_str(),
        WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
        pad + labelW + Scale(8), y, editW, rowH, hwndDlg, (HMENU)ID_EDIT_LEVEL, nullptr, nullptr);
    y += rowH + rowGap;

    CreateWindowW(L"STATIC", L"Awaken:", WS_VISIBLE | WS_CHILD, pad, y, labelW, rowH, hwndDlg, nullptr, nullptr, nullptr);
    state.hwndAwaken = CreateWindowW(L"EDIT", std::to_wstring(currentTranscend).c_str(),
        WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
        pad + labelW + Scale(8), y, editW, rowH, hwndDlg, (HMENU)ID_EDIT_AWAKEN, nullptr, nullptr);
    y += rowH + rowGap;

    if (hasSkillTable) {
        for (int i = 0; i < 7; ++i) {
            std::wstring label = L"Skill " + std::to_wstring(i + 1) + L":";
            CreateWindowW(L"STATIC", label.c_str(), WS_VISIBLE | WS_CHILD, pad, y, labelW, rowH, hwndDlg, nullptr, nullptr, nullptr);
            long long initial = currentSkill[i] < 0 ? 0 : currentSkill[i];
            state.hwndSkill[i] = CreateWindowW(L"EDIT", std::to_wstring(initial).c_str(),
                WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
                pad + labelW + Scale(8), y, editW, rowH, hwndDlg, (HMENU)(INT_PTR)(ID_EDIT_SKILL_BASE + i), nullptr, nullptr);
            y += rowH + rowGap;
        }
    }

    const int btnW = Scale(90);
    const int btnH = Scale(32);
    int okCancelY = clientH - pad - btnH;
    int maxMinY = okCancelY - Scale(12) - btnH;

    // Max/Min quick-fill buttons, left-aligned on their own row above OK/Cancel.
    CreateWindowW(L"BUTTON", L"Max", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        pad, maxMinY, btnW, btnH, hwndDlg, (HMENU)ID_BTN_EDIT_MAX, nullptr, nullptr);
    CreateWindowW(L"BUTTON", L"Min", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        pad + btnW + Scale(10), maxMinY, btnW, btnH, hwndDlg, (HMENU)ID_BTN_EDIT_MIN, nullptr, nullptr);

    CreateWindowW(L"BUTTON", L"OK", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
        clientW - pad - btnW * 2 - Scale(10), okCancelY, btnW, btnH, hwndDlg, (HMENU)ID_BTN_EDIT_OK, nullptr, nullptr);
    CreateWindowW(L"BUTTON", L"Cancel", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        clientW - pad - btnW, okCancelY, btnW, btnH, hwndDlg, (HMENU)ID_BTN_EDIT_CANCEL, nullptr, nullptr);

    // Apply the shared UI font to every child we just created.
    EnumChildWindows(hwndDlg, [](HWND child, LPARAM lp) -> BOOL {
        SendMessage(child, WM_SETFONT, (WPARAM)lp, TRUE);
        return TRUE;
        }, (LPARAM)g_hFontRegular);

    ShowWindow(hwndDlg, SW_SHOW);
    UpdateWindow(hwndDlg);

    MSG msg;
    while (IsWindow(hwndDlg)) {
        if (!GetMessage(&msg, nullptr, 0, 0)) break;
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    EnableWindow(parent, TRUE);
    SetForegroundWindow(parent);

    return state.confirmed;
}

// Writes one character's edited Level/EXP/ASCEND/Awaken and (if the table
// exists) 7 skill-slot levels back to the save file, all inside a single
// transaction. Reuses the same locate/key pipeline as everything else, so
// a misconfigured folder or missing save is reported the normal way.
static bool SaveCharacterEdit(long long cid, const CharacterEditResult& result, bool hasSkillTable) {
    if (!ensureSaveGamesDir()) {
        return false;
    }
    fs::path dbPath = locateSaveFile();
    if (dbPath.empty()) {
        return false;
    }

    sqlite3* db = nullptr;
    if (sqlite3_open(dbPath.string().c_str(), &db) != SQLITE_OK) {
        Log(std::string("Cannot open database: ") + sqlite3_errmsg(db));
        sqlite3_close(db);
        return false;
    }

    std::string keyPragma = "PRAGMA key = '" + std::string(SQLCIPHER_KEY) + "';";
    if (!execSql(db, keyPragma) || !verifyKey(db) || !tableExists(db, "tb_character")) {
        sqlite3_close(db);
        return false;
    }

    if (!execSql(db, "BEGIN TRANSACTION;")) {
        sqlite3_close(db);
        return false;
    }

    bool ok = true;

    {
        const char* sql = "UPDATE tb_character SET LEVEL=?, EXP=?, ASCEND=?, TRANSCEND=? WHERE CHARACTER_CID=?;";
        sqlite3_stmt* stmt = nullptr;
        if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) == SQLITE_OK) {
            sqlite3_bind_int64(stmt, 1, result.level);
            sqlite3_bind_int64(stmt, 2, result.exp);
            sqlite3_bind_int64(stmt, 3, result.ascend);
            sqlite3_bind_int64(stmt, 4, result.transcend);
            sqlite3_bind_int64(stmt, 5, cid);
            if (sqlite3_step(stmt) != SQLITE_DONE) {
                Log(std::string("Failed to update tb_character: ") + sqlite3_errmsg(db));
                ok = false;
            }
            sqlite3_finalize(stmt);
        }
        else {
            Log(std::string("Failed to prepare tb_character update: ") + sqlite3_errmsg(db));
            ok = false;
        }
    }

    if (ok && hasSkillTable) {
        // Newly-added characters (via Add Character) only get a row in
        // tb_character; tb_skill_growth never gets the matching 7 rows
        // (TYPE_VALUE 1..7) for them. UPDATE-ing a CHARACTER_CID that has
        // no rows yet silently affects 0 rows and returns SQLITE_DONE, so
        // the previous code looked like it "saved" while nothing actually
        // changed. Check whether this CID has any tb_skill_growth rows
        // first; if not, INSERT all 7 slots instead of UPDATE-ing them.
        bool hasSkillRows = false;
        {
            const char* checkSql = "SELECT 1 FROM tb_skill_growth WHERE CHARACTER_CID=? LIMIT 1;";
            sqlite3_stmt* checkStmt = nullptr;
            if (sqlite3_prepare_v2(db, checkSql, -1, &checkStmt, nullptr) == SQLITE_OK) {
                sqlite3_bind_int64(checkStmt, 1, cid);
                hasSkillRows = (sqlite3_step(checkStmt) == SQLITE_ROW);
                sqlite3_finalize(checkStmt);
            }
            else {
                Log(std::string("Failed to prepare skill existence check: ") + sqlite3_errmsg(db));
                ok = false;
            }
        }

        if (ok && !hasSkillRows) {
            // No tb_skill_growth rows for this CHARACTER_CID yet (e.g. a
            // character added via Add Character): insert all 7 slots.
            const char* insertSql =
                "INSERT INTO tb_skill_growth (USER_DBID, CHARACTER_CID, TYPE_VALUE, SLOT_LEVEL) "
                "VALUES (1000, ?, ?, ?);";
            for (int i = 0; i < 7 && ok; ++i) {
                sqlite3_stmt* stmt = nullptr;
                if (sqlite3_prepare_v2(db, insertSql, -1, &stmt, nullptr) != SQLITE_OK) {
                    Log(std::string("Failed to prepare skill insert: ") + sqlite3_errmsg(db));
                    ok = false;
                    break;
                }
                sqlite3_bind_int64(stmt, 1, cid);
                sqlite3_bind_int64(stmt, 2, i + 1);
                sqlite3_bind_int64(stmt, 3, result.skill[i]);
                int rc = sqlite3_step(stmt);
                sqlite3_finalize(stmt);
                if (rc != SQLITE_DONE) {
                    Log("Failed to insert skill slot " + std::to_string(i + 1) + ": " + sqlite3_errmsg(db));
                    ok = false;
                }
            }
        }
        else if (ok) {
            const char* sql = "UPDATE tb_skill_growth SET SLOT_LEVEL=? WHERE CHARACTER_CID=? AND TYPE_VALUE=?;";
            for (int i = 0; i < 7 && ok; ++i) {
                sqlite3_stmt* stmt = nullptr;
                if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
                    Log(std::string("Failed to prepare skill update: ") + sqlite3_errmsg(db));
                    ok = false;
                    break;
                }
                sqlite3_bind_int64(stmt, 1, result.skill[i]);
                sqlite3_bind_int64(stmt, 2, cid);
                sqlite3_bind_int64(stmt, 3, i + 1);
                int rc = sqlite3_step(stmt);
                sqlite3_finalize(stmt);
                if (rc != SQLITE_DONE) {
                    Log("Failed to update skill slot " + std::to_string(i + 1) + ": " + sqlite3_errmsg(db));
                    ok = false;
                }
            }
        }
    }

    if (!ok) {
        execSql(db, "ROLLBACK;");
        sqlite3_close(db);
        return false;
    }

    if (!execSql(db, "COMMIT;")) {
        sqlite3_close(db);
        return false;
    }

    sqlite3_close(db);
    return true;
}

// Deletes every row for one CHARACTER_CID out of both tb_character and (if
// present) tb_skill_growth, in a single transaction (all-or-nothing).
// Used by the roster list's right-click "Delete" context menu.
static bool DeleteCharacterFromRoster(long long cid) {
    if (!ensureSaveGamesDir()) {
        return false;
    }
    fs::path dbPath = locateSaveFile();
    if (dbPath.empty()) {
        return false;
    }

    sqlite3* db = nullptr;
    if (sqlite3_open(dbPath.string().c_str(), &db) != SQLITE_OK) {
        Log(std::string("Cannot open database: ") + sqlite3_errmsg(db));
        sqlite3_close(db);
        return false;
    }

    std::string keyPragma = "PRAGMA key = '" + std::string(SQLCIPHER_KEY) + "';";
    if (!execSql(db, keyPragma) || !verifyKey(db) || !tableExists(db, "tb_character")) {
        sqlite3_close(db);
        return false;
    }

    bool hasSkillTable = tableExists(db, "tb_skill_growth");

    if (!execSql(db, "BEGIN TRANSACTION;")) {
        sqlite3_close(db);
        return false;
    }

    bool ok = true;

    {
        const char* sql = "DELETE FROM tb_character WHERE USER_DBID = ? AND CHARACTER_CID = ?;";
        sqlite3_stmt* stmt = nullptr;
        if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) == SQLITE_OK) {
            sqlite3_bind_int64(stmt, 1, 1000);
            sqlite3_bind_int64(stmt, 2, cid);
            if (sqlite3_step(stmt) != SQLITE_DONE) {
                Log(std::string("Failed to delete from tb_character: ") + sqlite3_errmsg(db));
                ok = false;
            }
            sqlite3_finalize(stmt);
        }
        else {
            Log(std::string("Failed to prepare tb_character delete: ") + sqlite3_errmsg(db));
            ok = false;
        }
    }

    if (ok && hasSkillTable) {
        const char* sql = "DELETE FROM tb_skill_growth WHERE USER_DBID = ? AND CHARACTER_CID = ?;";
        sqlite3_stmt* stmt = nullptr;
        if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) == SQLITE_OK) {
            sqlite3_bind_int64(stmt, 1, 1000);
            sqlite3_bind_int64(stmt, 2, cid);
            if (sqlite3_step(stmt) != SQLITE_DONE) {
                Log(std::string("Failed to delete from tb_skill_growth: ") + sqlite3_errmsg(db));
                ok = false;
            }
            sqlite3_finalize(stmt);
        }
        else {
            Log(std::string("Failed to prepare tb_skill_growth delete: ") + sqlite3_errmsg(db));
            ok = false;
        }
    }

    if (!ok) {
        execSql(db, "ROLLBACK;");
        sqlite3_close(db);
        return false;
    }

    if (!execSql(db, "COMMIT;")) {
        sqlite3_close(db);
        return false;
    }

    sqlite3_close(db);
    return true;
}

// Creates the small font set used throughout the window. Segoe UI is the
// standard Windows UI font since Vista, so this is available everywhere
// without shipping a font file.
// ---------------------------------------------------------------------
// Equipment stats
//
// Accessory stat IDs are fully derivable, so instead of embedding all 155
// main-stat and 1350 sub-stat rows of Mainstat_translate.txt /
// Substat_translated.txt this reproduces the two ID formulas those files
// follow (verified exhaustively against every row in both) plus the fixed
// stat order inside each group:
//
//   MAIN_STAT_CID  = LV * 1000  + mainSlot * 100 + statIndex
//   SUB_STAT_CIDn  = LV * 10000 + subSlot * 1000 + n * 100 + statIndex
//
// LV is the piece's star grade, n is the sub-stat slot (1..5), and
// mainSlot / subSlot identify which accessory the stat belongs to:
//
//   piece   CHEST  LEG  HEAD  FOOT  HAND
//   main      1     2     3     4     5
//   sub       0     1     2     3     4     (always mainSlot - 1)
//
// That's what makes a stat piece-specific: LV4_ACC_FOOT_2 (group 43200)
// only ever appears in SUB_STAT_CID2 of a 4-star ACC_FOOT, so the editor
// below only ever offers a slot the IDs from its own group.
//
// LV3/LV4/LV5 groups are uniform -- statIndex 0..9 for main stats and
// 0..21 for sub stats, always in the orders listed below. LV2 is the odd
// one out: its main stat is a single fixed roll per piece (statIndex is
// always 11) and its sub stats offer only three options, at statIndex 0,
// 5 and 10. LV1 accessories appear in neither data file, so a 1-star
// piece is listed but shown read-only.
// ---------------------------------------------------------------------

// statIndex 0..9 of any LV3+ main stat group.
static const char* const kMainStatNames[10] = {
    "Attack", "Max HP", "Defence", "Attack %", "Max HP %", "Defence %",
    "Crit Rate %", "Crit Damage %", "Crit Resist %", "Crit Defence %"
};

// The single LV2 main stat roll, indexed by mainSlot - 1
// (CHEST, LEG, HEAD, FOOT, HAND). Its statIndex is always 11.
static const char* const kLv2MainStatNames[5] = {
    "Max HP", "Max HP", "Defence", "Crit Rate %", "Attack"
};

// statIndex 0..21 of any LV3+ sub stat group.
static const char* const kSubStatNames[22] = {
    "Attack", "Max HP", "Defence", "Attack %", "Max HP %", "Defence %",
    "Stun Damage", "Air Damage", "Down Damage", "Bleed Damage",
    "Fire Damage", "Ice Damage", "Shock Damage", "Poison Damage",
    "Crit Resist %", "Crit Defence %", "Break Damage", "All Damage",
    "Crit Rate %", "Crit Damage %", "Defence Penetrate", "Defence Negation %"
};

// The three LV2 sub stat rolls, and the statIndex each one sits at.
static const int kLv2SubStatIndices[3] = { 0, 5, 10 };
static const char* const kLv2SubStatNames[3] = { "Attack", "Max HP", "Defence" };

// The five accessory columns of tb_equip_mount. TALISMAN_1/TALISMAN_2,
// KARMA and VEHICLE are deliberately excluded: they aren't rows of
// tb_equipment, so they have no main/sub stats to edit.
struct AccessorySlot {
    const char* column;   // tb_equip_mount column holding the ITEM_DBID
    const char* label;    // what the equipment list calls it
    int mainSlotDigit;    // the "mainSlot" digit of the formulas above
};
static const AccessorySlot kAccessorySlots[] = {
    { "ACC_HEAD",  "Head",  3 },
    { "ACC_CHEST", "Chest", 1 },
    { "ACC_LEG",   "Leg",   2 },
    { "ACC_HAND",  "Hand",  5 },
    { "ACC_FOOT",  "Foot",  4 },
};
static const int kNumAccessorySlots = 5;

// One entry of a main/sub stat dropdown: the CID that gets written to
// tb_equipment plus the readable name shown to the user.
struct StatOption {
    long long id = 0;
    std::string name;
};

// True only for star grades both data files actually cover.
static bool IsKnownEquipLevel(int level) { return level >= 2 && level <= 5; }

// Every main stat a `level`-star piece in this accessory slot can roll.
static void BuildMainStatOptions(int level, int mainSlotDigit, std::vector<StatOption>& out) {
    out.clear();
    if (!IsKnownEquipLevel(level) || mainSlotDigit < 1 || mainSlotDigit > 5) return;
    long long group = (long long)level * 1000 + (long long)mainSlotDigit * 100;
    if (level == 2) {
        out.push_back(StatOption{ group + 11, kLv2MainStatNames[mainSlotDigit - 1] });
        return;
    }
    for (int i = 0; i < 10; ++i) {
        out.push_back(StatOption{ group + i, kMainStatNames[i] });
    }
}

// Every sub stat the `subIndex`-th (1..5) sub slot of a `level`-star
// piece in this accessory slot can roll, led by an explicit empty choice
// so a slot can also be cleared back to 0.
static void BuildSubStatOptions(int level, int mainSlotDigit, int subIndex,
    std::vector<StatOption>& out) {
    out.clear();
    if (!IsKnownEquipLevel(level) || mainSlotDigit < 1 || mainSlotDigit > 5) return;
    if (subIndex < 1 || subIndex > 5) return;
    out.push_back(StatOption{ 0, "(empty)" });
    long long group = (long long)level * 10000 +
        (long long)(mainSlotDigit - 1) * 1000 +
        (long long)subIndex * 100;
    if (level == 2) {
        for (int i = 0; i < 3; ++i) {
            out.push_back(StatOption{ group + kLv2SubStatIndices[i], kLv2SubStatNames[i] });
        }
        return;
    }
    for (int i = 0; i < 22; ++i) {
        out.push_back(StatOption{ group + i, kSubStatNames[i] });
    }
}

// Reverse lookups, used to label the values already sitting in the save
// file. Both return an empty string for an ID that doesn't decode, so the
// caller can fall back to printing the raw number instead of a wrong name.
static std::string MainStatDisplayName(long long id) {
    if (id <= 0) return std::string();
    int level = (int)(id / 1000);
    int slot = (int)((id / 100) % 10);
    int index = (int)(id % 100);
    if (!IsKnownEquipLevel(level) || slot < 1 || slot > 5) return std::string();
    if (level == 2) {
        return (index == 11) ? std::string(kLv2MainStatNames[slot - 1]) : std::string();
    }
    if (index > 9) return std::string();
    return kMainStatNames[index];
}

static std::string SubStatDisplayName(long long id) {
    if (id <= 0) return std::string();
    int level = (int)(id / 10000);
    int slot = (int)((id / 1000) % 10);
    int subIndex = (int)((id / 100) % 10);
    int index = (int)(id % 100);
    if (!IsKnownEquipLevel(level) || slot > 4) return std::string();
    if (subIndex < 1 || subIndex > 5) return std::string();
    if (level == 2) {
        for (int i = 0; i < 3; ++i) {
            if (index == kLv2SubStatIndices[i]) return kLv2SubStatNames[i];
        }
        return std::string();
    }
    if (index > 21) return std::string();
    return kSubStatNames[index];
}

// "Attack %" when the ID decodes, "Unrecognised" when it doesn't, and a
// dash for an unused slot -- the raw CID itself is never shown on screen.
static std::string DescribeMainStat(long long id) {
    if (id == 0) return "-";
    std::string name = MainStatDisplayName(id);
    return name.empty() ? "Unrecognised" : name;
}

static std::string DescribeSubStat(long long id) {
    if (id == 0) return "-";
    std::string name = SubStatDisplayName(id);
    return name.empty() ? "Unrecognised" : name;
}

// Just the stat's name -- the underlying CID is never shown on screen.
static std::string DescribeStatOption(const StatOption& option) {
    return option.name;
}

// All the labels here are ASCII, so the same widening the roster list
// already does inline is enough.
static std::wstring WidenAscii(const std::string& s) {
    return std::wstring(s.begin(), s.end());
}

// One accessory currently mounted on a character, as stored in
// tb_equipment. itemDbid == 0 means the slot is empty.
struct EquipItem {
    long long itemDbid = 0;
    long long itemCid = 0;
    long long enchant = 0;
    long long mainStat = 0;
    long long sub[5] = { 0, 0, 0, 0, 0 };
    int level = 0; // star grade 2..5; 0 when it couldn't be worked out
};

// One row of tb_equip_mount, with CHARACTER_CID already resolved to a
// display name through g_selectable[] exactly the way the roster does it.
struct EquipCharacter {
    std::string name;
    long long cid = 0;
    EquipItem slots[kNumAccessorySlots];
};

// Cached copy of what the equipment tab is showing, so the dropdown, the
// list and the edit dialog can all index the same data without going back
// to the DB.
static std::vector<EquipCharacter> g_equipChars;
static int g_equipSelectedChar = -1;

// Star grade of a piece. MAIN_STAT_CID's leading digit is the grade and
// it's also what the stat-ID formulas need, so it's the primary source;
// counting non-zero SUB_STAT_CIDs (a 4-star piece leaves SUB_STAT_CID5 at
// 0, a 3-star leaves 4 and 5 at 0, and so on) is the fallback for an item
// whose main stat is missing or doesn't decode.
static int DetermineEquipLevel(const EquipItem& item) {
    int fromMain = (int)(item.mainStat / 1000);
    if (IsKnownEquipLevel(fromMain)) return fromMain;
    int filled = 0;
    for (int i = 0; i < 5; ++i) {
        if (item.sub[i] != 0) ++filled;
    }
    return filled;
}

// Reads the tb_equipment row behind one mounted ITEM_DBID. Returns false
// if there's no such row, which the caller treats as an empty slot.
static bool queryEquipItem(sqlite3* db, long long itemDbid, EquipItem& out) {
    const char* sql =
        "SELECT ITEM_CID, ENCHANT_LEVEL, MAIN_STAT_CID, "
        "SUB_STAT_CID1, SUB_STAT_CID2, SUB_STAT_CID3, SUB_STAT_CID4, SUB_STAT_CID5 "
        "FROM tb_equipment WHERE ITEM_DBID = ?;";
    sqlite3_stmt* stmt = nullptr;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        Log(std::string("Failed to prepare equipment query: ") + sqlite3_errmsg(db));
        return false;
    }
    sqlite3_bind_int64(stmt, 1, itemDbid);

    bool found = false;
    if (sqlite3_step(stmt) == SQLITE_ROW) {
        out.itemDbid = itemDbid;
        out.itemCid = sqlite3_column_int64(stmt, 0);
        out.enchant = sqlite3_column_int64(stmt, 1);
        out.mainStat = sqlite3_column_int64(stmt, 2);
        for (int i = 0; i < 5; ++i) {
            out.sub[i] = sqlite3_column_int64(stmt, 3 + i);
        }
        out.level = DetermineEquipLevel(out);
        found = true;
    }
    sqlite3_finalize(stmt);
    return found;
}

// Walks tb_equip_mount -- one row per character -- and pulls in the
// tb_equipment row behind each of the five ACC_* columns. Only what the
// character is actually wearing is ever loaded; nothing in the inventory
// is touched.
static bool queryEquipMount(sqlite3* db, std::vector<EquipCharacter>& outChars) {
    std::string sql = "SELECT CHARACTER_CID";
    for (int i = 0; i < kNumAccessorySlots; ++i) {
        sql += ", ";
        sql += kAccessorySlots[i].column;
    }
    sql += " FROM tb_equip_mount ORDER BY CHARACTER_CID;";

    sqlite3_stmt* stmt = nullptr;
    if (sqlite3_prepare_v2(db, sql.c_str(), -1, &stmt, nullptr) != SQLITE_OK) {
        Log(std::string("Failed to prepare equip mount query: ") + sqlite3_errmsg(db));
        return false;
    }

    while (sqlite3_step(stmt) == SQLITE_ROW) {
        EquipCharacter entry;
        entry.cid = sqlite3_column_int64(stmt, 0);
        for (int i = 0; i < kNumSelectable; ++i) {
            if (g_selectable[i].character_cid == entry.cid) {
                entry.name = g_selectable[i].label;
                break;
            }
        }
        if (entry.name.empty()) {
            entry.name = "Unknown character";
        }

        for (int i = 0; i < kNumAccessorySlots; ++i) {
            long long itemDbid = sqlite3_column_int64(stmt, 1 + i);
            if (itemDbid == 0) continue;
            // A mounted ID whose tb_equipment row is gone is shown as an
            // empty slot rather than as a half-populated broken entry.
            if (!queryEquipItem(db, itemDbid, entry.slots[i])) {
                entry.slots[i] = EquipItem();
            }
        }
        outChars.push_back(entry);
    }
    sqlite3_finalize(stmt);
    return true;
}

// Opens the save file and reads every equipped accessory. Silent on
// failure for the same reason LoadRoster is: this runs on startup and
// after every edit, and a missing save is already reported the moment the
// user actually tries to change something.
static bool LoadEquipment(std::vector<EquipCharacter>& outChars) {
    if (!ensureSaveGamesDir(/*silent=*/true)) {
        return false;
    }
    fs::path dbPath = locateSaveFile(/*silent=*/true);
    if (dbPath.empty()) {
        return false;
    }

    sqlite3* db = nullptr;
    if (sqlite3_open(dbPath.string().c_str(), &db) != SQLITE_OK) {
        sqlite3_close(db);
        return false;
    }

    std::string keyPragma = "PRAGMA key = '" + std::string(SQLCIPHER_KEY) + "';";
    if (!execSql(db, keyPragma) || !verifyKey(db) ||
        !tableExists(db, "tb_equip_mount") || !tableExists(db, "tb_equipment")) {
        sqlite3_close(db);
        return false;
    }

    bool ok = queryEquipMount(db, outChars);
    sqlite3_close(db);
    return ok;
}

// Writes one accessory's edited main/sub stat CIDs back. ITEM_DBID is
// tb_equipment's primary key, so this always hits exactly the piece that
// was edited and nothing else.
static bool SaveEquipmentStats(long long itemDbid, long long mainStat, const long long sub[5]) {
    if (!ensureSaveGamesDir()) {
        return false;
    }
    fs::path dbPath = locateSaveFile();
    if (dbPath.empty()) {
        return false;
    }

    sqlite3* db = nullptr;
    if (sqlite3_open(dbPath.string().c_str(), &db) != SQLITE_OK) {
        Log(std::string("Cannot open database: ") + sqlite3_errmsg(db));
        sqlite3_close(db);
        return false;
    }

    std::string keyPragma = "PRAGMA key = '" + std::string(SQLCIPHER_KEY) + "';";
    if (!execSql(db, keyPragma) || !verifyKey(db) || !tableExists(db, "tb_equipment")) {
        sqlite3_close(db);
        return false;
    }

    if (!execSql(db, "BEGIN TRANSACTION;")) {
        sqlite3_close(db);
        return false;
    }

    bool ok = true;
    {
        const char* sql =
            "UPDATE tb_equipment SET MAIN_STAT_CID=?, "
            "SUB_STAT_CID1=?, SUB_STAT_CID2=?, SUB_STAT_CID3=?, "
            "SUB_STAT_CID4=?, SUB_STAT_CID5=? "
            "WHERE ITEM_DBID=?;";
        sqlite3_stmt* stmt = nullptr;
        if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) == SQLITE_OK) {
            sqlite3_bind_int64(stmt, 1, mainStat);
            for (int i = 0; i < 5; ++i) {
                sqlite3_bind_int64(stmt, 2 + i, sub[i]);
            }
            sqlite3_bind_int64(stmt, 7, itemDbid);
            if (sqlite3_step(stmt) != SQLITE_DONE) {
                Log(std::string("Failed to update tb_equipment: ") + sqlite3_errmsg(db));
                ok = false;
            }
            sqlite3_finalize(stmt);
        }
        else {
            Log(std::string("Failed to prepare tb_equipment update: ") + sqlite3_errmsg(db));
            ok = false;
        }
    }

    if (ok) {
        ok = execSql(db, "COMMIT;");
    }
    else {
        execSql(db, "ROLLBACK;");
    }

    sqlite3_close(db);
    return ok;
}

// ---- Edit-equipment dialog ----
// Per-dialog-instance state, stashed in GWLP_USERDATA the same way
// EditDialogProc does it.
struct EquipEditState {
    HWND hwndMain = nullptr;
    HWND hwndSub[5] = {};
    int subCount = 0;
    std::vector<StatOption> mainOptions;
    std::vector<StatOption> subOptions[5];
    long long* outMain = nullptr;
    long long* outSub = nullptr; // points at the caller's long long[5]
    bool confirmed = false;
};

static LRESULT CALLBACK EquipEditDialogProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    EquipEditState* state = reinterpret_cast<EquipEditState*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));

    switch (msg) {
    case WM_CREATE: {
        CREATESTRUCTW* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
        state = reinterpret_cast<EquipEditState*>(cs->lpCreateParams);
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)state);
        return 0;
    }
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)wParam;
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, ui::kCheckboxText);
        return (LRESULT)g_hbrWindowBg;
    }
    case WM_COMMAND: {
        int id = LOWORD(wParam);
        if (id == ID_BTN_EQ_CANCEL) {
            state->confirmed = false;
            DestroyWindow(hwnd);
            return 0;
        }
        if (id == ID_BTN_EQ_OK) {
            int sel = (int)SendMessageW(state->hwndMain, CB_GETCURSEL, 0, 0);
            if (sel < 0 || sel >= (int)state->mainOptions.size()) {
                MessageBoxW(hwnd, L"Please pick a main stat.", L"Main stat",
                    MB_OK | MB_ICONWARNING);
                return 0;
            }
            long long picked[5] = { 0, 0, 0, 0, 0 };
            std::string pickedName[5];
            for (int i = 0; i < state->subCount; ++i) {
                int s = (int)SendMessageW(state->hwndSub[i], CB_GETCURSEL, 0, 0);
                if (s < 0 || s >= (int)state->subOptions[i].size()) {
                    std::wstring warn = L"Please pick a value for sub stat " +
                        std::to_wstring(i + 1) + L".";
                    MessageBoxW(hwnd, warn.c_str(), L"Sub stat", MB_OK | MB_ICONWARNING);
                    return 0;
                }
                picked[i] = state->subOptions[i][s].id;
                pickedName[i] = state->subOptions[i][s].name;
            }

            // Every non-empty sub stat must be a distinct stat -- e.g. Crit
            // Damage % can't be rolled into two slots at once. Compare by
            // name rather than ID, since the same stat's CID differs across
            // slots (the slot number is baked into the group offset).
            for (int i = 0; i < state->subCount; ++i) {
                if (picked[i] == 0) continue;
                for (int j = i + 1; j < state->subCount; ++j) {
                    if (picked[j] != 0 && pickedName[i] == pickedName[j] && g_noDuplicateSubStat) {
                        std::wstring warn = L"Sub stat " + std::to_wstring(i + 1) +
                            L" and sub stat " + std::to_wstring(j + 1) +
                            L" are both \"" + WidenAscii(pickedName[i]) +
                            L"\". Each sub stat must be different.";
                        MessageBoxW(hwnd, warn.c_str(), L"Duplicate sub stat",
                            MB_OK | MB_ICONWARNING);
                        return 0;
                    }
                }
            }

            *state->outMain = state->mainOptions[sel].id;
            for (int i = 0; i < 5; ++i) {
                state->outSub[i] = picked[i];
            }
            state->confirmed = true;
            DestroyWindow(hwnd);
            return 0;
        }
        return 0;
    }
    case WM_CLOSE:
        state->confirmed = false;
        DestroyWindow(hwnd);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// Fills one CBS_DROPDOWNLIST with "<stat name>  (<CID>)" entries and
// selects whichever one matches `current`. A value that isn't a legal
// roll for this slot is kept as an extra first entry rather than being
// quietly replaced, so merely opening the dialog can never rewrite it.
static void PopulateStatCombo(HWND combo, std::vector<StatOption>& options, long long current) {
    int selected = -1;
    for (size_t i = 0; i < options.size(); ++i) {
        if (options[i].id == current) {
            selected = (int)i;
            break;
        }
    }
    if (selected < 0) {
        StatOption kept;
        kept.id = current;
        kept.name = "Unrecognised";
        options.insert(options.begin(), kept);
        selected = 0;
    }

    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
    for (const StatOption& option : options) {
        std::wstring text = WidenAscii(DescribeStatOption(option));
        SendMessageW(combo, CB_ADDSTRING, 0, (LPARAM)text.c_str());
    }
    SendMessageW(combo, CB_SETCURSEL, selected, 0);
}

// Modal-ish popup (same pattern as ShowEditCharacterDialog: owned by the
// parent, parent disabled while it's up) with one dropdown for the main
// stat and one per sub stat slot the piece actually has -- 5 dropdowns on
// a 5-star, 4 on a 4-star, and so on. Each dropdown lists only the stats
// that this exact piece / grade / slot combination is allowed to roll,
// by name. Returns true and fills outMain/outSub when confirmed.
static bool ShowEditEquipmentDialog(HWND parent, const std::string& characterName,
    const AccessorySlot& slot, const EquipItem& item,
    long long& outMain, long long(&outSub)[5]) {
    if (!IsKnownEquipLevel(item.level)) {
        MessageBoxA(parent,
            "This accessory's star grade isn't covered by the stat data, "
            "so its stats can't be edited.",
            "Unsupported equipment", MB_OK | MB_ICONWARNING);
        return false;
    }

    static bool classRegistered = false;
    const wchar_t* CLASS_NAME = L"EditEquipmentDialogClass";
    if (!classRegistered) {
        WNDCLASSW wc = {};
        wc.lpfnWndProc = EquipEditDialogProc;
        wc.hInstance = GetModuleHandleW(nullptr);
        wc.lpszClassName = CLASS_NAME;
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hbrBackground = g_hbrWindowBg;
        RegisterClassW(&wc);
        classRegistered = true;
    }

    EquipEditState state;
    // Star grade == number of sub stat slots: a 4-star leaves
    // SUB_STAT_CID5 at 0, so only four dropdowns are offered.
    state.subCount = item.level;
    state.outMain = &outMain;
    state.outSub = outSub;
    BuildMainStatOptions(item.level, slot.mainSlotDigit, state.mainOptions);
    for (int i = 0; i < state.subCount; ++i) {
        BuildSubStatOptions(item.level, slot.mainSlotDigit, i + 1, state.subOptions[i]);
    }

    const int pad = Scale(20);
    const int labelW = Scale(90);
    const int comboW = Scale(230);
    const int rowH = Scale(26);
    const int rowGap = Scale(8);
    const int dropH = Scale(260); // extra height reserved for the drop-down list

    const int rowCount = 1 + state.subCount;
    const int clientW = pad * 2 + labelW + Scale(8) + comboW;
    const int headerH = Scale(22);
    const int clientH = pad + headerH + Scale(10) + rowCount * (rowH + rowGap) +
        Scale(14) + Scale(32) + pad;

    std::wstring wtitle = L"Edit " + WidenAscii(slot.label) + L" - " + WidenAscii(characterName);

    RECT wr{ 0, 0, clientW, clientH };
    AdjustWindowRectEx(&wr, WS_POPUP | WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_DLGMODALFRAME);
    int winW = wr.right - wr.left;
    int winH = wr.bottom - wr.top;

    RECT parentRect;
    GetWindowRect(parent, &parentRect);
    int posX = parentRect.left + ((parentRect.right - parentRect.left) - winW) / 2;
    int posY = parentRect.top + ((parentRect.bottom - parentRect.top) - winH) / 2;

    EnableWindow(parent, FALSE);

    HWND hwndDlg = CreateWindowExW(WS_EX_DLGMODALFRAME, CLASS_NAME, wtitle.c_str(),
        WS_POPUP | WS_CAPTION | WS_SYSMENU,
        posX, posY, winW, winH,
        parent, nullptr, GetModuleHandleW(nullptr), &state);

    if (!hwndDlg) {
        EnableWindow(parent, TRUE);
        return false;
    }

    int y = pad;
    std::string headerText = std::to_string(item.level) + "-star " + slot.label;
    std::wstring wheader = WidenAscii(headerText);
    CreateWindowW(L"STATIC", wheader.c_str(), WS_VISIBLE | WS_CHILD,
        pad, y, clientW - pad * 2, headerH, hwndDlg, nullptr, nullptr, nullptr);
    y += headerH + Scale(10);

    const int comboX = pad + labelW + Scale(8);

    CreateWindowW(L"STATIC", L"Main stat:", WS_VISIBLE | WS_CHILD,
        pad, y + Scale(4), labelW, Scale(20), hwndDlg, nullptr, nullptr, nullptr);
    state.hwndMain = CreateWindowW(L"COMBOBOX", L"",
        WS_VISIBLE | WS_CHILD | WS_VSCROLL | CBS_DROPDOWNLIST,
        comboX, y, comboW, rowH + dropH,
        hwndDlg, (HMENU)ID_CMB_EQ_MAIN, nullptr, nullptr);
    PopulateStatCombo(state.hwndMain, state.mainOptions, item.mainStat);
    y += rowH + rowGap;

    for (int i = 0; i < state.subCount; ++i) {
        std::wstring label = L"Sub stat " + std::to_wstring(i + 1) + L":";
        CreateWindowW(L"STATIC", label.c_str(), WS_VISIBLE | WS_CHILD,
            pad, y + Scale(4), labelW, Scale(20), hwndDlg, nullptr, nullptr, nullptr);
        state.hwndSub[i] = CreateWindowW(L"COMBOBOX", L"",
            WS_VISIBLE | WS_CHILD | WS_VSCROLL | CBS_DROPDOWNLIST,
            comboX, y, comboW, rowH + dropH,
            hwndDlg, (HMENU)(INT_PTR)(ID_CMB_EQ_SUB_BASE + i), nullptr, nullptr);
        PopulateStatCombo(state.hwndSub[i], state.subOptions[i], item.sub[i]);
        y += rowH + rowGap;
    }

    const int btnW = Scale(90);
    const int btnH = Scale(32);
    const int btnY = clientH - pad - btnH;
    CreateWindowW(L"BUTTON", L"Save", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
        clientW - pad - btnW * 2 - Scale(10), btnY, btnW, btnH,
        hwndDlg, (HMENU)ID_BTN_EQ_OK, nullptr, nullptr);
    CreateWindowW(L"BUTTON", L"Cancel", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        clientW - pad - btnW, btnY, btnW, btnH,
        hwndDlg, (HMENU)ID_BTN_EQ_CANCEL, nullptr, nullptr);

    EnumChildWindows(hwndDlg, [](HWND child, LPARAM lp) -> BOOL {
        SendMessage(child, WM_SETFONT, (WPARAM)lp, TRUE);
        return TRUE;
        }, (LPARAM)g_hFontRegular);

    ShowWindow(hwndDlg, SW_SHOW);
    UpdateWindow(hwndDlg);

    MSG msg;
    while (IsWindow(hwndDlg)) {
        if (!GetMessage(&msg, nullptr, 0, 0)) break;
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    EnableWindow(parent, TRUE);
    SetForegroundWindow(parent);

    return state.confirmed;
}

// Fills the slot list for whichever character the dropdown is on. All
// five accessory slots are always listed, empty ones included, so it's
// obvious at a glance what isn't equipped.
static void RefreshEquipmentList() {
    if (!g_hwndEquipList) return;
    ListView_DeleteAllItems(g_hwndEquipList);

    if (g_equipSelectedChar < 0 || g_equipSelectedChar >= (int)g_equipChars.size()) {
        LVITEMW item = {};
        item.mask = LVIF_TEXT;
        item.iItem = 0;
        item.iSubItem = 0;
        item.pszText = const_cast<LPWSTR>(L"No equipped characters found");
        ListView_InsertItem(g_hwndEquipList, &item);
        return;
    }

    const EquipCharacter& character = g_equipChars[g_equipSelectedChar];
    for (int row = 0; row < kNumAccessorySlots; ++row) {
        const EquipItem& piece = character.slots[row];
        std::wstring wslot = WidenAscii(kAccessorySlots[row].label);

        LVITEMW item = {};
        item.mask = LVIF_TEXT;
        item.iItem = row;
        item.iSubItem = 0;
        item.pszText = const_cast<LPWSTR>(wslot.c_str());
        ListView_InsertItem(g_hwndEquipList, &item);

        if (piece.itemDbid == 0) {
            ListView_SetItemText(g_hwndEquipList, row, 1, const_cast<LPWSTR>(L"-"));
            ListView_SetItemText(g_hwndEquipList, row, 2, const_cast<LPWSTR>(L"(nothing equipped)"));
            for (int i = 0; i < 5; ++i) {
                ListView_SetItemText(g_hwndEquipList, row, 3 + i, const_cast<LPWSTR>(L"-"));
            }
            continue;
        }

        std::wstring wstar = (piece.level > 0) ? std::to_wstring(piece.level) : std::wstring(L"?");
        ListView_SetItemText(g_hwndEquipList, row, 1, const_cast<LPWSTR>(wstar.c_str()));

        std::wstring wmain = WidenAscii(DescribeMainStat(piece.mainStat));
        ListView_SetItemText(g_hwndEquipList, row, 2, const_cast<LPWSTR>(wmain.c_str()));

        for (int i = 0; i < 5; ++i) {
            std::wstring wsub = WidenAscii(DescribeSubStat(piece.sub[i]));
            ListView_SetItemText(g_hwndEquipList, row, 3 + i, const_cast<LPWSTR>(wsub.c_str()));
        }
    }
}

// Reloads the whole equipment tab from the save file. The character
// currently picked in the dropdown is kept selected when it's still
// there, so refreshing after an edit doesn't jump back to the top.
static void RefreshEquipmentTab() {
    if (!g_hwndEquipCharCombo || !g_hwndEquipList) return;

    long long keepCid = 0;
    if (g_equipSelectedChar >= 0 && g_equipSelectedChar < (int)g_equipChars.size()) {
        keepCid = g_equipChars[g_equipSelectedChar].cid;
    }

    g_equipChars.clear();
    g_equipSelectedChar = -1;
    SendMessageW(g_hwndEquipCharCombo, CB_RESETCONTENT, 0, 0);

    std::vector<EquipCharacter> characters;
    if (!LoadEquipment(characters) || characters.empty()) {
        EnableWindow(g_hwndEquipCharCombo, FALSE);
        RefreshEquipmentList();
        return;
    }

    g_equipChars = characters;
    EnableWindow(g_hwndEquipCharCombo, TRUE);

    int select = 0;
    for (size_t i = 0; i < g_equipChars.size(); ++i) {
        std::wstring wname = WidenAscii(g_equipChars[i].name);
        SendMessageW(g_hwndEquipCharCombo, CB_ADDSTRING, 0, (LPARAM)wname.c_str());
        if (g_equipChars[i].cid == keepCid) {
            select = (int)i;
        }
    }
    SendMessageW(g_hwndEquipCharCombo, CB_SETCURSEL, select, 0);
    g_equipSelectedChar = select;

    RefreshEquipmentList();
}

// Opens the stat editor for one row of the equipment list and, when the
// user confirms, writes the new CIDs back and reloads the tab.
static void EditEquipmentRow(HWND parent, int row) {
    if (g_equipSelectedChar < 0 || g_equipSelectedChar >= (int)g_equipChars.size()) return;
    if (row < 0 || row >= kNumAccessorySlots) return;

    const AccessorySlot& slot = kAccessorySlots[row]; // static table, safe to alias
    // These two are copies rather than references: the dialog runs its own
    // message loop, which can outlive a refresh that reallocates
    // g_equipChars underneath us.
    EquipItem piece = g_equipChars[g_equipSelectedChar].slots[row];
    std::string characterName = g_equipChars[g_equipSelectedChar].name;

    if (piece.itemDbid == 0) {
        std::string msg = characterName + " has nothing equipped in the " +
            slot.label + " slot.";
        MessageBoxA(parent, msg.c_str(), "Nothing to edit", MB_OK | MB_ICONINFORMATION);
        return;
    }

    long long newMain = piece.mainStat;
    long long newSub[5] = { 0, 0, 0, 0, 0 };
    for (int i = 0; i < 5; ++i) {
        newSub[i] = piece.sub[i];
    }

    if (!ShowEditEquipmentDialog(parent, characterName, slot, piece, newMain, newSub)) {
        return;
    }

    bool saved = SaveEquipmentStats(piece.itemDbid, newMain, newSub);
    std::string caption = characterName + " - " + slot.label;
    MessageBoxA(parent,
        saved ? "Equipment stats updated." : "Failed to save equipment stats.",
        caption.c_str(),
        MB_OK | (saved ? MB_ICONINFORMATION : MB_ICONERROR));
    RefreshEquipmentTab();
}

// Switches which page's controls are visible. Both pages are plain
// siblings in the main window, so this is just a bulk show/hide.
static void ShowTabPage(int index) {
    g_activeTab = index;
    for (HWND child : g_tabPageCharacters) {
        ShowWindow(child, index == 0 ? SW_SHOW : SW_HIDE);
    }
    for (HWND child : g_tabPageEquipment) {
        ShowWindow(child, index == 1 ? SW_SHOW : SW_HIDE);
    }
}

static void CreateModernFonts() {
    LOGFONTW lf = {};
    wcscpy_s(lf.lfFaceName, L"Segoe UI");
    lf.lfWeight = FW_NORMAL;
    lf.lfQuality = CLEARTYPE_QUALITY;
    lf.lfCharSet = DEFAULT_CHARSET;

    lf.lfHeight = -Scale(14);
    lf.lfWeight = FW_SEMIBOLD;
    g_hFontTitle = CreateFontIndirectW(&lf);

    lf.lfHeight = -Scale(10);
    lf.lfWeight = FW_NORMAL;
    g_hFontSubtitle = CreateFontIndirectW(&lf);

    lf.lfHeight = -Scale(11);
    lf.lfWeight = FW_NORMAL;
    g_hFontRegular = CreateFontIndirectW(&lf);

    lf.lfHeight = -Scale(11);
    lf.lfWeight = FW_SEMIBOLD;
    g_hFontButton = CreateFontIndirectW(&lf);
}

// Subclasses the two owner-draw buttons purely to track hover/pressed
// state (needed for the "lift on hover" look); actual click handling
// still goes through the normal BN_CLICKED / WM_COMMAND path.
static LRESULT CALLBACK ButtonSubclassProc(HWND hwnd, UINT msg, WPARAM wParam,
    LPARAM lParam, UINT_PTR uIdSubclass,
    DWORD_PTR dwRefData) {
    ButtonVisual* vis = reinterpret_cast<ButtonVisual*>(dwRefData);
    switch (msg) {
    case WM_MOUSEMOVE:
        if (!vis->hover) {
            vis->hover = true;
            TRACKMOUSEEVENT tme{ sizeof(tme), TME_LEAVE, hwnd, 0 };
            TrackMouseEvent(&tme);
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        break;
    case WM_MOUSELEAVE:
        vis->hover = false;
        vis->pressed = false;
        InvalidateRect(hwnd, nullptr, FALSE);
        break;
    case WM_LBUTTONDOWN:
        vis->pressed = true;
        InvalidateRect(hwnd, nullptr, FALSE);
        break;
    case WM_LBUTTONUP:
        vis->pressed = false;
        InvalidateRect(hwnd, nullptr, FALSE);
        break;
    case WM_NCDESTROY:
        RemoveWindowSubclass(hwnd, ButtonSubclassProc, uIdSubclass);
        break;
    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

// Owner-draw painter for the Add/Remove buttons: flat rounded-rect,
// filled accent color for the primary action, outlined for the secondary
// one, with hover/pressed shading. No GDI+ dependency needed.
static void DrawModernButton(LPDRAWITEMSTRUCT dis, const ButtonVisual& vis) {
    HDC hdc = dis->hDC;
    RECT rc = dis->rcItem;

    COLORREF fill, border, textColor;
    if (vis.primary) {
        fill = vis.pressed ? ui::kAccentPressed : (vis.hover ? ui::kAccentHover : ui::kAccent);
        border = fill;
        textColor = RGB(0xFF, 0xFF, 0xFF);
    }
    else {
        fill = vis.pressed ? ui::kSecondaryPressedFill
            : (vis.hover ? ui::kSecondaryHoverFill : ui::kCardBg);
        border = ui::kSecondaryBorder;
        textColor = ui::kSecondaryText;
    }

    HBRUSH hBrush = CreateSolidBrush(fill);
    HPEN hPen = CreatePen(PS_SOLID, 1, border);
    HGDIOBJ oldBrush = SelectObject(hdc, hBrush);
    HGDIOBJ oldPen = SelectObject(hdc, hPen);
    int radius = Scale(8);
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, radius, radius);
    SelectObject(hdc, oldBrush);
    SelectObject(hdc, oldPen);
    DeleteObject(hBrush);
    DeleteObject(hPen);

    wchar_t text[64] = {};
    GetWindowTextW(dis->hwndItem, text, 64);

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);
    HGDIOBJ oldFont = SelectObject(hdc, g_hFontButton);
    DrawTextW(hdc, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, oldFont);
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        g_dpi = GetDpiForWindow(hwnd);
        CreateModernFonts();
        g_hbrWindowBg = CreateSolidBrush(ui::kWindowBg);
        g_hbrCardBg = CreateSolidBrush(ui::kCardBg);

        const int pad = Scale(24);
        const int panelGap = Scale(28);

        // Left panel: title/subtitle, a 2-column checkbox grid (25
        // characters no longer fit nicely in a single column), then the
        // Add/Remove buttons spanning the full grid width.
        const int checkboxColW = Scale(150);
        const int checkboxColGap = Scale(16);
        const int numCheckCols = 2;
        const int rowsPerCol = (kNumSelectable + numCheckCols - 1) / numCheckCols;
        const int leftContentW = checkboxColW * numCheckCols + checkboxColGap * (numCheckCols - 1);

        // Right-panel geometry, hoisted up here because the tab strip (and
        // the Backup Save button parked at its right-hand end) need the
        // window's final content width before anything else is laid out.
        const int rightX = pad + leftContentW + panelGap;
        const int colCharacterW = Scale(120);
        const int colLevelW = Scale(50);
        const int colAwakenW = Scale(60);
        const int colSkillW = Scale(42);
        const int rightW = colCharacterW + colLevelW + colAwakenW + colSkillW * 7;
        const int windowContentW = rightX + rightW + pad;

        // Tab strip across the top: "Characters" (everything this tool did
        // before) and "Equipment" (the worn-accessory stat editor). It's
        // sized to the header row only -- both pages are ordinary sibling
        // children laid out underneath and toggled by ShowTabPage().
        const int backupBtnW = Scale(150);
        const int backupBtnH = Scale(32);
        const int tabTop = Scale(10);
        const int tabH = Scale(30);
        g_hwndTabs = CreateWindowExW(0, WC_TABCONTROLW, L"",
            WS_VISIBLE | WS_CHILD | TCS_TABS | TCS_FIXEDWIDTH,
            pad, tabTop, windowContentW - pad * 2 - backupBtnW - Scale(12), tabH,
            hwnd, (HMENU)ID_TAB_MAIN, nullptr, nullptr);
        SendMessage(g_hwndTabs, WM_SETFONT, (WPARAM)g_hFontRegular, TRUE);
        SendMessage(g_hwndTabs, TCM_SETITEMSIZE, 0, MAKELPARAM(Scale(130), Scale(24)));
        {
            TCITEMW tab = {};
            tab.mask = TCIF_TEXT;
            tab.pszText = const_cast<LPWSTR>(L"Characters");
            TabCtrl_InsertItem(g_hwndTabs, 0, &tab);
            tab.pszText = const_cast<LPWSTR>(L"Equipment");
            TabCtrl_InsertItem(g_hwndTabs, 1, &tab);
        }

        // Everything below the tab strip shifts down by this much; the
        // original layout constants are otherwise left exactly as they were.
        const int topOff = tabTop + tabH + Scale(8);

        // Backup Save covers the whole save file rather than one page, so
        // it sits on the tab row itself and stays visible on both tabs.
        HWND hwndBackup = CreateWindowW(L"BUTTON", L"Backup Save",
            WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,
            windowContentW - pad - backupBtnW, tabTop, backupBtnW, backupBtnH,
            hwnd, (HMENU)ID_BTN_BACKUP, nullptr, nullptr);
        SetWindowSubclass(hwndBackup, ButtonSubclassProc, 4, (DWORD_PTR)&g_backupBtnVisual);

        int y = topOff + Scale(20);

        g_hwndTitle = CreateWindowW(L"STATIC", L"DragonSword Awakening",
            WS_VISIBLE | WS_CHILD,
            pad, y, leftContentW, Scale(24),
            hwnd, nullptr, nullptr, nullptr);
        SendMessage(g_hwndTitle, WM_SETFONT, (WPARAM)g_hFontTitle, TRUE);
        y += Scale(26);

        g_hwndSubtitle = CreateWindowW(L"STATIC", L"Tick character(s), then add or remove",
            WS_VISIBLE | WS_CHILD,
            pad, y, leftContentW, Scale(18),
            hwnd, nullptr, nullptr, nullptr);
        SendMessage(g_hwndSubtitle, WM_SETFONT, (WPARAM)g_hFontSubtitle, TRUE);
        y += Scale(18) + Scale(16); // leaves room for the divider drawn in WM_PAINT

        const int checkboxTop = y;
        for (int i = 0; i < kNumSelectable; ++i) {
            int col = i / rowsPerCol;
            int row = i % rowsPerCol;
            int cx = pad + col * (checkboxColW + checkboxColGap);
            int cy = checkboxTop + row * Scale(26);

            std::wstring wlabel(g_selectable[i].label, g_selectable[i].label + strlen(g_selectable[i].label));
            g_selectable[i].hwndCheckbox = CreateWindowW(
                L"BUTTON", wlabel.c_str(),
                WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
                cx, cy, checkboxColW, Scale(22),
                hwnd, (HMENU)(INT_PTR)g_selectable[i].controlId, nullptr, nullptr);
            SendMessage(g_selectable[i].hwndCheckbox, WM_SETFONT, (WPARAM)g_hFontRegular, TRUE);
        }
        y = checkboxTop + rowsPerCol * Scale(26) + Scale(14);

        HWND hwndAdd = CreateWindowW(L"BUTTON", L"Add Character",
            WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,
            pad, y, leftContentW, Scale(40),
            hwnd, (HMENU)ID_BTN_ADD, nullptr, nullptr);
        SetWindowSubclass(hwndAdd, ButtonSubclassProc, 1, (DWORD_PTR)&g_addBtnVisual);
        y += Scale(40) + Scale(12);

        HWND hwndRefresh = CreateWindowW(L"BUTTON", L"Refresh Roster",
            WS_VISIBLE | WS_CHILD | BS_OWNERDRAW,
            pad, y, leftContentW, Scale(40),
            hwnd, (HMENU)ID_BTN_REFRESH, nullptr, nullptr);
        SetWindowSubclass(hwndRefresh, ButtonSubclassProc, 3, (DWORD_PTR)&g_refreshBtnVisual);
        y += Scale(40) + Scale(20);

        const int leftContentBottom = y;

        // Right panel: read-only roster pulled straight from tb_character
        // (+ tb_skill_growth for the 7 skill slots), refreshed on open and
        // after every Add/Remove. Its title spans the full panel width.
        g_hwndRosterTitle = CreateWindowW(L"STATIC", L"Current Roster",
            WS_VISIBLE | WS_CHILD,
            rightX, topOff + Scale(20), rightW, Scale(24),
            hwnd, nullptr, nullptr, nullptr);
        SendMessage(g_hwndRosterTitle, WM_SETFONT, (WPARAM)g_hFontTitle, TRUE);

        // "Team name:" label + editable box, sitting in the gap between
        // the "Current Roster" title and the roster list below it. Backed
        // by tb_advent_status.NICK_NAME (see LoadTeamName/SaveTeamName);
        // the box is populated by RefreshTeamNameEdit() below and saved
        // back when it loses focus (see the EN_KILLFOCUS handling in
        // WM_COMMAND).
        const int teamNameLabelW = Scale(90);
        const int teamNameRowY = topOff + Scale(48);
        const int teamNameRowH = Scale(24);
        g_hwndTeamNameLabel = CreateWindowW(L"STATIC", L"Team name:",
            WS_VISIBLE | WS_CHILD,
            rightX, teamNameRowY, teamNameLabelW, teamNameRowH,
            hwnd, nullptr, nullptr, nullptr);
        g_hwndTeamNameEdit = CreateWindowW(L"EDIT", L"",
            WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
            rightX + teamNameLabelW + Scale(8), teamNameRowY,
            rightW - teamNameLabelW - Scale(8), teamNameRowH,
            hwnd, (HMENU)ID_EDIT_TEAM_NAME, nullptr, nullptr);
        SendMessage(g_hwndTeamNameLabel, WM_SETFONT, (WPARAM)g_hFontRegular, TRUE);
        SendMessage(g_hwndTeamNameEdit, WM_SETFONT, (WPARAM)g_hFontRegular, TRUE);

        const int rosterListH = leftContentBottom - checkboxTop;
        g_hwndRosterList = CreateWindowExW(0, WC_LISTVIEWW, L"",
            WS_VISIBLE | WS_CHILD | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
            rightX, checkboxTop, rightW, rosterListH,
            hwnd, (HMENU)ID_LIST_ROSTER, nullptr, nullptr);
        ListView_SetExtendedListViewStyle(g_hwndRosterList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
        SendMessage(g_hwndRosterList, WM_SETFONT, (WPARAM)g_hFontRegular, TRUE);

        LVCOLUMNW col = {};
        col.mask = LVCF_TEXT | LVCF_WIDTH;
        col.cx = colCharacterW;
        col.pszText = const_cast<LPWSTR>(L"Character");
        ListView_InsertColumn(g_hwndRosterList, 0, &col);
        col.cx = colLevelW;
        col.pszText = const_cast<LPWSTR>(L"Level");
        ListView_InsertColumn(g_hwndRosterList, 1, &col);
        col.cx = colAwakenW;
        col.pszText = const_cast<LPWSTR>(L"Awaken");
        ListView_InsertColumn(g_hwndRosterList, 2, &col);
        for (int i = 0; i < 7; ++i) {
            std::wstring header = L"Skill " + std::to_wstring(i + 1);
            col.cx = colSkillW;
            col.pszText = const_cast<LPWSTR>(header.c_str());
            ListView_InsertColumn(g_hwndRosterList, 3 + i, &col);
        }

        RefreshRosterList(g_hwndRosterList);
        RefreshTeamNameEdit(g_hwndTeamNameEdit);

        g_tabPageCharacters.push_back(g_hwndTitle);
        g_tabPageCharacters.push_back(g_hwndSubtitle);
        for (int i = 0; i < kNumSelectable; ++i) {
            g_tabPageCharacters.push_back(g_selectable[i].hwndCheckbox);
        }
        g_tabPageCharacters.push_back(hwndAdd);
        g_tabPageCharacters.push_back(hwndRefresh);
        g_tabPageCharacters.push_back(g_hwndRosterTitle);
        g_tabPageCharacters.push_back(g_hwndTeamNameLabel);
        g_tabPageCharacters.push_back(g_hwndTeamNameEdit);
        g_tabPageCharacters.push_back(g_hwndRosterList);

        // ---- Equipment page ----
        // One full-width column: a character picker fed from
        // tb_equip_mount, then a row per accessory slot showing the stats
        // currently on the piece worn there. Created hidden (no WS_VISIBLE)
        // because the Characters tab is the one selected on open.
        const int equipTitleY = topOff + Scale(20);
        g_hwndEquipTitle = CreateWindowW(L"STATIC", L"Equipment Stats",
            WS_CHILD,
            pad, equipTitleY, Scale(300), Scale(24),
            hwnd, nullptr, nullptr, nullptr);
        SendMessage(g_hwndEquipTitle, WM_SETFONT, (WPARAM)g_hFontTitle, TRUE);

        const int equipRowY = equipTitleY + Scale(34);
        const int equipLabelW = Scale(80);
        g_hwndEquipCharLabel = CreateWindowW(L"STATIC", L"Character:",
            WS_CHILD,
            pad, equipRowY + Scale(4), equipLabelW, Scale(20),
            hwnd, nullptr, nullptr, nullptr);
        g_hwndEquipCharCombo = CreateWindowW(L"COMBOBOX", L"",
            WS_CHILD | WS_VSCROLL | CBS_DROPDOWNLIST,
            pad + equipLabelW + Scale(8), equipRowY, Scale(220), Scale(26) + Scale(240),
            hwnd, (HMENU)ID_CMB_EQUIP_CHAR, nullptr, nullptr);
        SendMessage(g_hwndEquipCharLabel, WM_SETFONT, (WPARAM)g_hFontRegular, TRUE);
        SendMessage(g_hwndEquipCharCombo, WM_SETFONT, (WPARAM)g_hFontRegular, TRUE);

        const int equipRefreshW = Scale(170);
        HWND hwndEquipRefresh = CreateWindowW(L"BUTTON", L"Refresh Equipment",
            WS_CHILD | BS_OWNERDRAW,
            windowContentW - pad - equipRefreshW, equipRowY - Scale(3), equipRefreshW, Scale(32),
            hwnd, (HMENU)ID_BTN_EQUIP_REFRESH, nullptr, nullptr);
        SetWindowSubclass(hwndEquipRefresh, ButtonSubclassProc, 5, (DWORD_PTR)&g_equipRefreshBtnVisual);

        const int equipHintY = equipRowY + Scale(40);
        g_hwndEquipHint = CreateWindowW(L"STATIC",
            L"Only what this character is currently wearing is listed. "
            L"Double-click (or right-click) a row to change that piece's stats.",
            WS_CHILD,
            pad, equipHintY, windowContentW - pad * 2, Scale(18),
            hwnd, nullptr, nullptr, nullptr);
        SendMessage(g_hwndEquipHint, WM_SETFONT, (WPARAM)g_hFontSubtitle, TRUE);

        const int equipListY = equipHintY + Scale(26);
        const int equipListW = windowContentW - pad * 2;
        g_hwndEquipList = CreateWindowExW(0, WC_LISTVIEWW, L"",
            WS_CHILD | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
            pad, equipListY, equipListW, leftContentBottom - Scale(20) - equipListY,
            hwnd, (HMENU)ID_LIST_EQUIP, nullptr, nullptr);
        ListView_SetExtendedListViewStyle(g_hwndEquipList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
        SendMessage(g_hwndEquipList, WM_SETFONT, (WPARAM)g_hFontRegular, TRUE);

        {
            const int colSlotW = Scale(70);
            const int colStarW = Scale(44);
            const int colMainW = Scale(134);
            const int colSubW = (equipListW - colSlotW - colStarW - colMainW) / 5;

            LVCOLUMNW ecol = {};
            ecol.mask = LVCF_TEXT | LVCF_WIDTH;
            ecol.cx = colSlotW;
            ecol.pszText = const_cast<LPWSTR>(L"Slot");
            ListView_InsertColumn(g_hwndEquipList, 0, &ecol);
            ecol.cx = colStarW;
            ecol.pszText = const_cast<LPWSTR>(L"Star");
            ListView_InsertColumn(g_hwndEquipList, 1, &ecol);
            ecol.cx = colMainW;
            ecol.pszText = const_cast<LPWSTR>(L"Main Stat");
            ListView_InsertColumn(g_hwndEquipList, 2, &ecol);
            for (int i = 0; i < 5; ++i) {
                std::wstring header = L"Sub " + std::to_wstring(i + 1);
                ecol.cx = colSubW;
                ecol.pszText = const_cast<LPWSTR>(header.c_str());
                ListView_InsertColumn(g_hwndEquipList, 3 + i, &ecol);
            }
        }

        g_tabPageEquipment.push_back(g_hwndEquipTitle);
        g_tabPageEquipment.push_back(g_hwndEquipCharLabel);
        g_tabPageEquipment.push_back(g_hwndEquipCharCombo);
        g_tabPageEquipment.push_back(hwndEquipRefresh);
        g_tabPageEquipment.push_back(g_hwndEquipHint);
        g_tabPageEquipment.push_back(g_hwndEquipList);

        RefreshEquipmentTab();
        ShowTabPage(0);

        // Resize the window to exactly fit the content we just laid out.
        RECT client{ 0, 0, windowContentW, leftContentBottom };
        AdjustWindowRectEx(&client, (DWORD)GetWindowLongPtr(hwnd, GWL_STYLE),
            FALSE, (DWORD)GetWindowLongPtr(hwnd, GWL_EXSTYLE));
        SetWindowPos(hwnd, nullptr, 0, 0, client.right - client.left, client.bottom - client.top,
            SWP_NOMOVE | SWP_NOZORDER);

        return 0;
    }
    case WM_CTLCOLORSTATIC: {
        HDC hdcStatic = (HDC)wParam;
        HWND hwndCtl = (HWND)lParam;
        SetBkMode(hdcStatic, TRANSPARENT);
        if (hwndCtl == g_hwndTitle || hwndCtl == g_hwndRosterTitle ||
            hwndCtl == g_hwndEquipTitle) {
            SetTextColor(hdcStatic, ui::kTitleText);
        }
        else if (hwndCtl == g_hwndSubtitle || hwndCtl == g_hwndEquipHint) {
            SetTextColor(hdcStatic, ui::kSubtitleText);
        }
        else {
            SetTextColor(hdcStatic, ui::kCheckboxText);
        }
        return (LRESULT)g_hbrWindowBg;
    }
    case WM_CTLCOLORBTN: {
        // Only reached for the (non owner-draw) checkboxes; the two
        // buttons are BS_OWNERDRAW and painted entirely in WM_DRAWITEM.
        HDC hdcBtn = (HDC)wParam;
        SetBkMode(hdcBtn, TRANSPARENT);
        SetTextColor(hdcBtn, ui::kCheckboxText);
        return (LRESULT)g_hbrWindowBg;
    }
    case WM_DRAWITEM: {
        LPDRAWITEMSTRUCT dis = (LPDRAWITEMSTRUCT)lParam;
        if (dis->CtlID == ID_BTN_ADD) {
            DrawModernButton(dis, g_addBtnVisual);
            return TRUE;
        }
        if (dis->CtlID == ID_BTN_REFRESH) {
            DrawModernButton(dis, g_refreshBtnVisual);
            return TRUE;
        }
        if (dis->CtlID == ID_BTN_BACKUP) {
            DrawModernButton(dis, g_backupBtnVisual);
            return TRUE;
        }
        if (dis->CtlID == ID_BTN_EQUIP_REFRESH) {
            DrawModernButton(dis, g_equipRefreshBtnVisual);
            return TRUE;
        }
        break;
    }
    case WM_ERASEBKGND: {
        HDC hdc = (HDC)wParam;
        RECT rc;
        GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_hbrWindowBg);
        return 1;
    }
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT rc;
        GetClientRect(hwnd, &rc);
        // Thin divider under the title/subtitle block. Only the Characters
        // page has that block, so the Equipment page doesn't get a stray
        // line across it. The leading Scale(10)+Scale(30)+Scale(8) is the
        // same tab-strip offset WM_CREATE applies to the page content.
        if (g_activeTab == 0) {
            RECT divider{ Scale(24),
                           Scale(10) + Scale(30) + Scale(8) +
                           Scale(20) + Scale(26) + Scale(18) + Scale(8),
                           rc.right - Scale(24), 0 };
            divider.bottom = divider.top + 1;
            HBRUSH hbrDivider = CreateSolidBrush(ui::kDivider);
            FillRect(hdc, &divider, hbrDivider);
            DeleteObject(hbrDivider);
        }
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_NOTIFY: {
        LPNMHDR nmhdr = reinterpret_cast<LPNMHDR>(lParam);
        if (nmhdr->idFrom == ID_TAB_MAIN && nmhdr->code == TCN_SELCHANGE) {
            ShowTabPage(TabCtrl_GetCurSel(g_hwndTabs));
            // The two pages overlap, and this window has no
            // WS_CLIPCHILDREN, so repaint the background AND every child
            // that stayed visible -- otherwise the page that was just
            // hidden leaves pixels behind.
            RedrawWindow(hwnd, nullptr, nullptr,
                RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
            return 0;
        }
        if (nmhdr->idFrom == ID_LIST_EQUIP &&
            (nmhdr->code == NM_DBLCLK || nmhdr->code == NM_RCLICK)) {
            // Both gestures open the same stat editor for the clicked
            // accessory slot; there's only one action, so no context menu.
            LPNMITEMACTIVATE nmia = reinterpret_cast<LPNMITEMACTIVATE>(lParam);
            EditEquipmentRow(hwnd, nmia->iItem);
            return 0;
        }
        if (nmhdr->idFrom == ID_LIST_ROSTER && nmhdr->code == NM_RCLICK) {
            // Right-click on a roster row: resolve which row was clicked,
            // then show a two-item "Edit Character" / "Delete" context menu.
            LPNMITEMACTIVATE nmia = reinterpret_cast<LPNMITEMACTIVATE>(lParam);
            int row = nmia->iItem;
            if (row >= 0 && row < (int)g_currentRoster.size()) {
                RosterEntry sel = g_currentRoster[row]; // copy: menu/dialog loop can outlive a refresh

                POINT pt = nmia->ptAction; // client coords within the list view
                ClientToScreen(g_hwndRosterList, &pt);

                HMENU hMenu = CreatePopupMenu();
                AppendMenuW(hMenu, MF_STRING, ID_CTX_EDIT, L"Edit Character");
                AppendMenuW(hMenu, MF_STRING, ID_CTX_DELETE, L"Delete");
                SetForegroundWindow(hwnd);
                UINT clicked = TrackPopupMenu(hMenu,
                    TPM_RETURNCMD | TPM_RIGHTBUTTON | TPM_NONOTIFY,
                    pt.x, pt.y, 0, hwnd, nullptr);
                DestroyMenu(hMenu);

                if (clicked == ID_CTX_EDIT) {
                    CharacterEditResult result;
                    if (ShowEditCharacterDialog(hwnd, sel.name, sel.level, sel.transcend, sel.skill, g_hasSkillTable, result)) {
                        bool saved = SaveCharacterEdit(sel.character_cid, result, g_hasSkillTable);
                        MessageBoxA(hwnd,
                            saved ? "Character updated." : "Failed to save changes.",
                            sel.name.c_str(),
                            MB_OK | (saved ? MB_ICONINFORMATION : MB_ICONERROR));
                        RefreshRosterList(g_hwndRosterList);
                    }
                }
                else if (clicked == ID_CTX_DELETE) {
                    std::wstring wname(sel.name.begin(), sel.name.end());
                    std::wstring confirmMsg = L"Delete " + wname +
                        L"?\n\nThis action cannot be undone.";
                    if (MessageBoxW(hwnd, confirmMsg.c_str(), L"Confirm Delete",
                        MB_YESNO | MB_ICONWARNING) == IDYES) {
                        bool deleted = DeleteCharacterFromRoster(sel.character_cid);
                        MessageBoxA(hwnd,
                            deleted ? "Character deleted." : "Failed to delete character.",
                            sel.name.c_str(),
                            MB_OK | (deleted ? MB_ICONINFORMATION : MB_ICONERROR));
                        RefreshRosterList(g_hwndRosterList);
                    }
                }
            }
        }
        return 0;
    }
    case WM_COMMAND: {
        int id = LOWORD(wParam);
        if (id == ID_BTN_ADD) {
            std::vector<int> selected = getSelectedIndices();
            if (selected.empty()) {
                MessageBoxA(
                    hwnd,
                    "Please tick at least one character before continuing.",
                    "No character selected",
                    MB_ICONWARNING | MB_OK
                );
                return 0;
            }

            OperationOutcome outcome = RunOperation(/*doAdd=*/true, selected);
            ShowResult(hwnd, /*doAdd=*/true, outcome);
            RefreshRosterList(g_hwndRosterList);
        }
        else if (id == ID_BTN_REFRESH) {
            RefreshRosterList(g_hwndRosterList);
            RefreshTeamNameEdit(g_hwndTeamNameEdit);
        }
        else if (id == ID_CMB_EQUIP_CHAR && HIWORD(wParam) == CBN_SELCHANGE) {
            g_equipSelectedChar = (int)SendMessageW(g_hwndEquipCharCombo, CB_GETCURSEL, 0, 0);
            RefreshEquipmentList();
        }
        else if (id == ID_BTN_EQUIP_REFRESH) {
            RefreshEquipmentTab();
        }
        else if (id == ID_BTN_BACKUP) {
            std::string backupPath;
            bool ok = BackupSaveFolder(backupPath);
            if (ok) {
                std::string msg = "Save folder backed up successfully to:\n\n" + backupPath;
                MessageBoxA(hwnd, msg.c_str(), "Backup Save", MB_OK | MB_ICONINFORMATION);
            }
            else {
                MessageBoxA(hwnd, "Failed to back up the save folder.", "Backup Save",
                    MB_OK | MB_ICONERROR);
            }
        }
        else if (id == ID_EDIT_TEAM_NAME && HIWORD(wParam) == EN_KILLFOCUS) {
            // Auto-save the team name when the box loses focus, but only
            // if there's actually a table to write to and the value
            // changed -- avoids nagging the user with a failure box every
            // time they merely tab through the field.
            if (g_hasAdventStatusTable) {
                wchar_t buf[256] = {};
                GetWindowTextW(g_hwndTeamNameEdit, buf, 256);
                std::wstring wnew(buf);
                std::string newName(wnew.begin(), wnew.end());
                if (newName != g_lastLoadedTeamName) {
                    if (SaveTeamName(newName)) {
                        g_lastLoadedTeamName = newName;
                    }
                    else {
                        MessageBoxA(hwnd, "Failed to save team name.", "Team name",
                            MB_OK | MB_ICONERROR);
                        // Revert the box to the last known-good value.
                        std::wstring wrevert(g_lastLoadedTeamName.begin(), g_lastLoadedTeamName.end());
                        SetWindowTextW(g_hwndTeamNameEdit, wrevert.c_str());
                    }
                }
            }
        }
        return 0;
    }
    case WM_DESTROY:
        if (g_hFontTitle) DeleteObject(g_hFontTitle);
        if (g_hFontSubtitle) DeleteObject(g_hFontSubtitle);
        if (g_hFontRegular) DeleteObject(g_hFontRegular);
        if (g_hFontButton) DeleteObject(g_hFontButton);
        if (g_hbrWindowBg) DeleteObject(g_hbrWindowBg);
        if (g_hbrCardBg) DeleteObject(g_hbrCardBg);
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// ---------------------------------------------------------------------
// SaveGames picker window.
//
// Shown INSTEAD of the full menu whenever tryAutoDetectSaveGamesDir()
// fails at startup. Nothing about the roster/checkbox UI is created or
// drawn yet -- just a small explanation and a "Browse..." button. The
// selected folder is validated (must exist and be literally named
// "SaveGames"); an invalid pick shows a warning and lets the user try
// again. Closing this window (or its Exit button) quits the app without
// ever creating the main window.
// ---------------------------------------------------------------------

// Best-effort visual polish shared by both this window and the main one;
// harmless no-ops on Windows versions that don't support these DWM
// attributes.
static void ApplyModernTitleBarStyle(HWND hwnd) {
    BOOL useDarkTitleBar = FALSE;
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
        0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        DWORD value = 1, size = sizeof(value);
        if (RegQueryValueExW(hKey, L"AppsUseLightTheme", nullptr, nullptr,
            (LPBYTE)&value, &size) == ERROR_SUCCESS) {
            useDarkTitleBar = (value == 0);
        }
        RegCloseKey(hKey);
    }
    DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &useDarkTitleBar, sizeof(useDarkTitleBar));
    DWORD cornerPref = DWMWCP_ROUND;
    DwmSetWindowAttribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, &cornerPref, sizeof(cornerPref));
}

// Opens the classic "Browse For Folder" dialog. Returns true and fills
// outPath if the user picked a folder; false on cancel/error. Does not
// validate the name itself -- the caller does that so it can show a
// tailored message.
// Opens the modern "Open Folder" picker (IFileOpenDialog with
// FOS_PICKFOLDERS) -- the same folder-browsing dialog Explorer/Visual
// Studio use, with the quick-access sidebar, breadcrumb path bar, and
// search box -- instead of the old tree-view "Browse For Folder" dialog.
// Returns true and fills outPath if the user picked a folder; false on
// cancel/error. Requires COM to already be initialized on this thread
// (see ComInitGuard in WinMain).
static bool BrowseForFolder(HWND hwndParent, const std::wstring& title, fs::path& outPath) {
    IFileOpenDialog* pDialog = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&pDialog));
    if (FAILED(hr) || !pDialog) {
        return false;
    }

    DWORD options = 0;
    pDialog->GetOptions(&options);
    pDialog->SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST);
    pDialog->SetTitle(title.c_str());

    bool result = false;
    hr = pDialog->Show(hwndParent);
    if (SUCCEEDED(hr)) {
        IShellItem* pItem = nullptr;
        hr = pDialog->GetResult(&pItem);
        if (SUCCEEDED(hr) && pItem) {
            PWSTR pszPath = nullptr;
            hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszPath);
            if (SUCCEEDED(hr) && pszPath) {
                outPath = fs::path(pszPath);
                CoTaskMemFree(pszPath);
                result = true;
            }
            pItem->Release();
        }
    }
    // hr == HRESULT from Show() being a "cancelled" error (e.g.
    // HRESULT_FROM_WIN32(ERROR_CANCELLED)) just falls through with
    // result left false -- that's a normal, silent cancel.

    pDialog->Release();
    return result;
}

static HFONT g_hFontPickerTitle = nullptr;
static HFONT g_hFontPickerBody = nullptr;
// Set true only once the user has picked a folder that passes
// isValidSaveGamesFolder(); read by RunSaveGamesPicker() after its
// message loop ends to decide whether to proceed to the main window.
static bool g_pickerAccepted = false;

static LRESULT CALLBACK PickerWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        g_dpi = GetDpiForWindow(hwnd);

        g_hFontPickerTitle = CreateFontW(-Scale(16), 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
            DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
        g_hFontPickerBody = CreateFontW(-Scale(11), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
            DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");

        const int marginX = Scale(24);
        const int width = Scale(400) - marginX * 2;
        int y = Scale(24);

        HWND hTitle = CreateWindowW(L"STATIC", L"SaveGames folder not found",
            WS_VISIBLE | WS_CHILD, marginX, y, width, Scale(24),
            hwnd, nullptr, nullptr, nullptr);
        SendMessage(hTitle, WM_SETFONT, (WPARAM)g_hFontPickerTitle, TRUE);
        y += Scale(34);

        HWND hBody = CreateWindowW(L"STATIC",
            L"The tool couldn't automatically find your game's save data.\r\n\r\n"
            L"Please browse to it manually. The folder you pick MUST be "
            L"named exactly \"SaveGames\" (for example ...\\DS\\Saved\\SaveGames).",
            WS_VISIBLE | WS_CHILD, marginX, y, width, Scale(92),
            hwnd, nullptr, nullptr, nullptr);
        SendMessage(hBody, WM_SETFONT, (WPARAM)g_hFontPickerBody, TRUE);
        y += Scale(104);

        HWND hwndBrowse = CreateWindowW(L"BUTTON", L"Browse for SaveGames folder...",
            WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON, marginX, y, width, Scale(34),
            hwnd, (HMENU)(INT_PTR)ID_BTN_PICKER_BROWSE, nullptr, nullptr);
        SendMessage(hwndBrowse, WM_SETFONT, (WPARAM)g_hFontPickerBody, TRUE);
        y += Scale(44);

        HWND hwndExit = CreateWindowW(L"BUTTON", L"Exit",
            WS_VISIBLE | WS_CHILD, marginX, y, width, Scale(28),
            hwnd, (HMENU)(INT_PTR)ID_BTN_PICKER_EXIT, nullptr, nullptr);
        SendMessage(hwndExit, WM_SETFONT, (WPARAM)g_hFontPickerBody, TRUE);

        return 0;
    }
    case WM_COMMAND: {
        int id = LOWORD(wParam);
        if (id == ID_BTN_PICKER_BROWSE) {
            fs::path picked;
            if (BrowseForFolder(hwnd, L"Select your SaveGames folder", picked)) {
                if (isValidSaveGamesFolder(picked)) {
                    std::error_code ec;
                    fs::path abs = fs::absolute(picked, ec);
                    g_saveGamesDir = !ec ? abs : picked;
                    g_pickerAccepted = true;
                    DestroyWindow(hwnd);
                }
                else {
                    MessageBoxW(hwnd,
                        L"That folder isn't valid. Please select a folder named exactly \"SaveGames\".",
                        L"Invalid folder", MB_OK | MB_ICONWARNING);
                }
            }
            // Dialog cancelled: just leave the picker window open.
        }
        else if (id == ID_BTN_PICKER_EXIT) {
            DestroyWindow(hwnd);
        }
        return 0;
    }
    case WM_DESTROY:
        if (g_hFontPickerTitle) { DeleteObject(g_hFontPickerTitle); g_hFontPickerTitle = nullptr; }
        if (g_hFontPickerBody) { DeleteObject(g_hFontPickerBody); g_hFontPickerBody = nullptr; }
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// Creates and shows the picker window, then pumps its own message loop
// until that window is destroyed. Returns true iff the user picked a
// valid SaveGames folder (g_saveGamesDir has been set to it); false means
// the caller should exit the app without ever creating the main window.
static bool RunSaveGamesPicker(HINSTANCE hInstance, int nCmdShow) {
    g_pickerAccepted = false;

    const wchar_t* CLASS_NAME = L"SaveGamesPickerWindowClass";

    WNDCLASSW wc = {};
    wc.lpfnWndProc = PickerWndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    RegisterClassW(&wc);

    HWND hwnd = CreateWindowExW(
        0, CLASS_NAME, L"DragonSword Awakening - Character Editor",
        WS_OVERLAPPEDWINDOW & ~WS_THICKFRAME & ~WS_MAXIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, 440, 320,
        nullptr, nullptr, hInstance, nullptr);

    if (!hwnd) {
        return false;
    }

    ApplyModernTitleBarStyle(hwnd);

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return g_pickerAccepted;
}

// RAII COM initialization for this thread -- required by IFileOpenDialog
// (BrowseForFolder, used by the SaveGames picker window). Lives for the
// whole of WinMain so CoUninitialize() runs automatically no matter which
// return path is taken.
struct ComInitGuard {
    HRESULT hr;
    ComInitGuard() : hr(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE)) {}
    ~ComInitGuard() { if (SUCCEEDED(hr)) CoUninitialize(); }
};

int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int nCmdShow) {
    ComInitGuard comGuard;

    // Per-monitor DPI awareness so the layout (scaled via Scale()) looks
    // crisp and correctly sized on high-DPI displays. Falls back quietly
    // on older Windows versions where the newer context isn't available.
    HMODULE hUser32 = GetModuleHandleW(L"user32.dll");
    if (hUser32) {
        typedef BOOL(WINAPI* SetCtxFn)(DPI_AWARENESS_CONTEXT);
        auto setCtx = reinterpret_cast<SetCtxFn>(GetProcAddress(hUser32, "SetProcessDpiAwarenessContext"));
        if (setCtx) {
            setCtx(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
        }
        else {
            SetProcessDPIAware();
        }
    }

    INITCOMMONCONTROLSEX icc{ sizeof(icc),
        ICC_STANDARD_CLASSES | ICC_LISTVIEW_CLASSES | ICC_TAB_CLASSES };
    InitCommonControlsEx(&icc);

    // Resolve the SaveGames folder before drawing anything else. Try the
    // old silent auto-detect first; if that fails, show ONLY the small
    // picker window (no roster/checkbox menu at all) until the user
    // supplies a valid folder, or bail out entirely if they close it.
    if (!tryAutoDetectSaveGamesDir()) {
        if (!RunSaveGamesPicker(hInstance, nCmdShow)) {
            return 0;
        }
    }

    const wchar_t* CLASS_NAME = L"InsertCharacterWindowClass";

    WNDCLASSW wc = {};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = CreateSolidBrush(ui::kWindowBg);
    RegisterClassW(&wc);

    // Initial size is a placeholder; WM_CREATE resizes the window to
    // exactly fit the laid-out controls once it knows the real DPI.
    HWND hwnd = CreateWindowExW(
        0, CLASS_NAME, L"DragonSword Awakening - Editor - v2 - By Daleth",
        WS_OVERLAPPEDWINDOW & ~WS_THICKFRAME & ~WS_MAXIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, 340, 420,
        nullptr, nullptr, hInstance, nullptr);

    if (!hwnd) {
        return 0;
    }

    // Best-effort visual polish; harmless no-ops on Windows versions that
    // don't support these DWM attributes.
    ApplyModernTitleBarStyle(hwnd);

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}