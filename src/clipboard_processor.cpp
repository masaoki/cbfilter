/**
 * @file clipboard_processor.cpp
 * @brief Implementation of clipboard access functions
 */

#include "clipboard_processor.h"

#include <cstring>
#include <stdexcept>

namespace {
/**
 * @brief RAII wrapper for clipboard access
 * Automatically closes clipboard when destroyed
 */
struct ClipboardGuard {
    ClipboardGuard(HWND hwnd) { open = OpenClipboard(hwnd); }
    ~ClipboardGuard() { if (open) CloseClipboard(); }
    bool open{};
};

/**
 * @brief Convert Windows clipboard format to ClipboardType enum
 * @param fmt Windows clipboard format identifier
 * @return Corresponding ClipboardType
 */
ClipboardType FormatToType(UINT fmt) {
    if (fmt == CF_UNICODETEXT) return ClipboardType::Text;
    if (fmt == CF_BITMAP || fmt == CF_DIB || fmt == CF_DIBV5) return ClipboardType::Bitmap;
    return ClipboardType::None;
}
} // namespace

ClipboardType DetectClipboard() {
    ClipboardGuard guard(nullptr);
    if (!guard.open) return ClipboardType::None;
    // Check for text content first
    if (IsClipboardFormatAvailable(CF_UNICODETEXT)) return ClipboardType::Text;
    // Check for bitmap formats (CF_BITMAP, CF_DIB, CF_DIBV5)
    if (IsClipboardFormatAvailable(CF_BITMAP) || IsClipboardFormatAvailable(CF_DIB) || IsClipboardFormatAvailable(CF_DIBV5))
        return ClipboardType::Bitmap;
    return ClipboardType::None;
}

std::wstring GetClipboardText() {
    ClipboardGuard guard(nullptr);
    if (!guard.open) return L"";
    HANDLE h = GetClipboardData(CF_UNICODETEXT);
    if (!h) return L"";
    auto* data = static_cast<wchar_t*>(GlobalLock(h));
    if (!data) return L"";
    std::wstring text(data);
    GlobalUnlock(h);
    return text;
}

HBITMAP GetClipboardBitmap() {
    ClipboardGuard guard(nullptr);
    if (!guard.open) return nullptr;
    HBITMAP bmp = static_cast<HBITMAP>(GetClipboardData(CF_BITMAP));
    if (!bmp) return nullptr;
    // Duplicate the bitmap so we own it (clipboard bitmap is read-only)
    BITMAP bm{};
    if (!GetObjectW(bmp, sizeof(BITMAP), &bm)) return nullptr;
    HDC hdc = GetDC(nullptr);
    HDC mem = CreateCompatibleDC(hdc);
    HBITMAP copy = CreateCompatibleBitmap(hdc, bm.bmWidth, bm.bmHeight);
    HGDIOBJ old = SelectObject(mem, copy);
    HDC src = CreateCompatibleDC(hdc);
    HGDIOBJ oldSrc = SelectObject(src, bmp);
    BitBlt(mem, 0, 0, bm.bmWidth, bm.bmHeight, src, 0, 0, SRCCOPY);
    SelectObject(src, oldSrc);
    SelectObject(mem, old);
    DeleteDC(src);
    DeleteDC(mem);
    ReleaseDC(nullptr, hdc);
    return copy;
}

void SetClipboardText(const std::wstring& text) {
    ClipboardGuard guard(nullptr);
    if (!guard.open) throw std::runtime_error("OpenClipboard failed");
    EmptyClipboard();
    // Allocate global memory for text (including null terminator)
    size_t bytes = (text.size() + 1) * sizeof(wchar_t);
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (!hMem) throw std::runtime_error("GlobalAlloc failed");
    void* dst = GlobalLock(hMem);
    memcpy(dst, text.c_str(), bytes);
    GlobalUnlock(hMem);
    SetClipboardData(CF_UNICODETEXT, hMem);
}

void SetClipboardBitmap(HBITMAP bmp) {
    ClipboardGuard guard(nullptr);
    if (!guard.open) throw std::runtime_error("OpenClipboard failed");
    EmptyClipboard();
    
    // Get bitmap info
    BITMAP bm{};
    if (!GetObjectW(bmp, sizeof(BITMAP), &bm)) {
        throw std::runtime_error("GetObject failed for bitmap");
    }
    
    // Create a device-dependent bitmap (DDB) copy for clipboard
    // This ensures compatibility with Windows clipboard
    HDC hdc = GetDC(nullptr);
    HDC memDC = CreateCompatibleDC(hdc);
    
    // Create a compatible bitmap
    HBITMAP hClipBmp = CreateCompatibleBitmap(hdc, bm.bmWidth, bm.bmHeight);
    if (!hClipBmp) {
        DeleteDC(memDC);
        ReleaseDC(nullptr, hdc);
        throw std::runtime_error("CreateCompatibleBitmap failed");
    }
    
    HGDIOBJ oldBmp = SelectObject(memDC, hClipBmp);
    
    // Create source DC and copy the bitmap
    HDC srcDC = CreateCompatibleDC(hdc);
    HGDIOBJ oldSrcBmp = SelectObject(srcDC, bmp);
    
    // Copy bitmap bits
    if (!BitBlt(memDC, 0, 0, bm.bmWidth, bm.bmHeight, srcDC, 0, 0, SRCCOPY)) {
        SelectObject(srcDC, oldSrcBmp);
        DeleteDC(srcDC);
        SelectObject(memDC, oldBmp);
        DeleteObject(hClipBmp);
        DeleteDC(memDC);
        ReleaseDC(nullptr, hdc);
        throw std::runtime_error("BitBlt failed");
    }
    
    SelectObject(srcDC, oldSrcBmp);
    DeleteDC(srcDC);
    SelectObject(memDC, oldBmp);
    DeleteDC(memDC);
    ReleaseDC(nullptr, hdc);
    
    // Set the copied bitmap to clipboard (CF_BITMAP requires DDB)
    HANDLE hResult = SetClipboardData(CF_BITMAP, hClipBmp);
    if (!hResult) {
        DeleteObject(hClipBmp);
        DWORD err = GetLastError();
        throw std::runtime_error("SetClipboardData(CF_BITMAP) failed: " + std::to_string(err));
    }
    
    // Also create DIB format for better compatibility
    hdc = GetDC(nullptr);
    memDC = CreateCompatibleDC(hdc);
    oldBmp = SelectObject(memDC, bmp);
    
    BITMAPINFO bi{};
    bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bi.bmiHeader.biWidth = bm.bmWidth;
    bi.bmiHeader.biHeight = bm.bmHeight;
    bi.bmiHeader.biPlanes = 1;
    bi.bmiHeader.biBitCount = 24; // Use 24-bit RGB for better compatibility
    bi.bmiHeader.biCompression = BI_RGB;
    
    // Calculate DIB size
    int rowSize = ((bm.bmWidth * 24 + 31) / 32) * 4; // Align to 4 bytes
    DWORD dibSize = sizeof(BITMAPINFOHEADER) + rowSize * abs(bm.bmHeight);
    
    HGLOBAL hDib = GlobalAlloc(GMEM_MOVEABLE, dibSize);
    if (hDib) {
        BITMAPINFO* pbi = static_cast<BITMAPINFO*>(GlobalLock(hDib));
        if (pbi) {
            pbi->bmiHeader = bi.bmiHeader;
            BYTE* pBits = reinterpret_cast<BYTE*>(pbi) + sizeof(BITMAPINFOHEADER);
            if (GetDIBits(memDC, bmp, 0, abs(bm.bmHeight), pBits, pbi, DIB_RGB_COLORS)) {
                GlobalUnlock(hDib);
                SetClipboardData(CF_DIB, hDib);
            } else {
                GlobalUnlock(hDib);
                GlobalFree(hDib);
            }
        } else {
            GlobalFree(hDib);
        }
    }
    
    SelectObject(memDC, oldBmp);
    DeleteDC(memDC);
    ReleaseDC(nullptr, hdc);
    
    // Note: CF_BITMAP ownership is transferred to clipboard, so we don't delete hClipBmp
    // Original bmp is still owned by caller
}

void SendCtrlV() {
    // Simulate Ctrl+V keypress sequence
    INPUT inputs[4]{};
    inputs[0].type = INPUT_KEYBOARD; inputs[0].ki.wVk = VK_CONTROL;  // Ctrl down
    inputs[1].type = INPUT_KEYBOARD; inputs[1].ki.wVk = 'V';          // V down
    inputs[2].type = INPUT_KEYBOARD; inputs[2].ki.wVk = 'V'; inputs[2].ki.dwFlags = KEYEVENTF_KEYUP;  // V up
    inputs[3].type = INPUT_KEYBOARD; inputs[3].ki.wVk = VK_CONTROL; inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;  // Ctrl up
    SendInput(4, inputs, sizeof(INPUT));
}
