/**
 * @file clipboard_processor.h
 * @brief Clipboard access and manipulation functions
 * 
 * Provides functions to read from and write to the Windows clipboard,
 * detect clipboard content type, and simulate paste operations.
 */

#pragma once

#include <string>
#include <windows.h>

/**
 * @enum ClipboardType
 * @brief Type of content currently in the clipboard
 */
enum class ClipboardType { None, Text, Bitmap };

/**
 * @brief Detect the type of content in the clipboard
 * @return ClipboardType indicating text, bitmap, or none
 */
ClipboardType DetectClipboard();

/**
 * @brief Get text content from the clipboard
 * @return Clipboard text as wide string, or empty string if unavailable
 */
std::wstring GetClipboardText();

/**
 * @brief Get bitmap from the clipboard
 * @return Bitmap handle (caller must call DeleteObject), or nullptr if unavailable
 */
HBITMAP GetClipboardBitmap();

/**
 * @brief Process text content (legacy function, not currently used)
 * @param in Input text
 * @return Processed text (currently just trims whitespace)
 */
std::wstring ProcessText(const std::wstring& in);

/**
 * @brief Process bitmap content (legacy function, not currently used)
 * @param bmp Input bitmap
 * @return Processed bitmap (currently returns original unchanged)
 */
HBITMAP ProcessBitmap(HBITMAP bmp);

/**
 * @brief Set text content to the clipboard
 * @param text Text to set
 * @throws std::runtime_error if clipboard operations fail
 */
void SetClipboardText(const std::wstring& text);

/**
 * @brief Set bitmap to the clipboard
 * @param bmp Bitmap handle (ownership transferred to clipboard)
 * @throws std::runtime_error if clipboard operations fail
 */
void SetClipboardBitmap(HBITMAP bmp);

/**
 * @brief Simulate Ctrl+V keypress to paste clipboard content
 */
void SendCtrlV();
