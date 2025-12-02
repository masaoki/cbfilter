# cbfilter

A Windows clipboard transformation utility that uses AI models to process clipboard content. Transform text or images with your definition through various AI-powered filters with a simple hotkey.

## Features

- **Text Processing**: Translate, summarize, or transform text using AI language models
- **Image Processing**: Analyze images using vision models or generate images from text
- **Multiple AI Models**: Configure multiple model endpoints (OpenAI, local servers, etc.)
- **Custom Filters**: Create custom transformation filters with your own prompts
- **System Tray Integration**: Runs in the background with system tray icon
- **Hotkey Support**: Press `Win+Alt+V` to show filter menu and process clipboard
- **Persistent Configuration**: Settings saved to `%APPDATA%\cbfilter\config.json` file

## Supported Transformations

The application supports four types of transformations:

1. **Text → Text**: Text completion, translation, summarization, etc.
2. **Text → Image**: Generate images from text prompts
3. **Image → Text**: Analyze images and extract text descriptions
4. **Image → Image**: Transform images using AI models

## Requirements

- Windows 10 or later
- C++20 compatible compiler (MSVC, MinGW, etc.)
- Windows SDK

## Building

Simply run:
```bash
build.bat
```

Debug logs are enabled with a build option.

```bash
build.bat debug
```

## Configuration

On first run, the setup dialg appears. You can edit these files directly or use the settings dialog.
The configuration is saved as `%APPDATA%cbfilter\config.json`. The setting dialog is shown only when the config file does not exist.

### Language Settings

The following languages are supported.

- English
- Japanese
- German
- Spanish
- French
- Italian
- Korean
- Dutch
- Portuguise
- Russian
- Thai
- Vietnamese
- Chinese

You can change the language on the setup dialog, but the default configuration depends on the `defconf.json` file.
The installer installs the default configuration as `defconf.json` for the installation language.

### Hotkey Configuration

The hotkey configuration can be changed with the setup dialog.
It is also saved in `%APPDATA%cbfilter\config.json`.

### Model Configuration

Each model requires:
- **Name**: Display name for the model
- **Server URL**: API endpoint (e.g., `https://api.openai.com/v1`)
- **Model Name**: Model identifier (e.g., `gpt-5.1`, `gpt-image-1`)
- **API Key**: Authentication key for the API

### Filter Configuration

Each filter requires:
- **Title**: Display name shown in the filter menu
- **Input Type**: Text or Image
- **Output Type**: Text or Image
- **Model**: Which AI model to use
- **Prompt**: Instruction text sent to the AI model

## Usage

1. **Start the application**: Run `cbfilter.exe`. It will appear in the system tray.

2. **Configure models and filters**: 
   - Right-click the system tray icon and select "Settings"
   - Add or edit models and filters as needed

3. **Use filters**:
   - Copy text or image to clipboard
   - Press `Win+Alt+V` to show the filter menu
   - Select a filter from the menu
   - The clipboard will be replaced with the transformed content
   - The result will be automatically pasted (Ctrl+V simulated)

4. **System tray menu**:
   - Right-click: Show settings menu
   - Double-click: Open settings window
   - Left-click: Show filter menu (if clipboard has compatible content)

## Logging

The debug log `cbfilter.log` is created only with the debug build of `cbfilter.exe`.

## API Compatibility

This application is compatible with OpenAI API format, or Google Gemini API format.
The API structure is inside template `apidefs/<Provider_Name>.json` file.
Currently OpenAI, Gemini, and OpenRouter configurations are included.
OpenAI configuration will fit most of OpenAI compatible APIs like LiteLLM Proxy, Requesty, etc.
Please let us know if you create any useful API provider definition template.

## License

This project is provided as MIT License.
