# Kanna for VBA

[![Release](https://img.shields.io/github/v/release/alineapps/kanna-vba)](https://github.com/alineapps/kanna-vba/releases/latest)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

**A lightweight code indenter for the VBA editor (VBE).** Re-indent your Excel / Access VBA code with a single keystroke.

*Kanna* (鉋) is the Japanese woodworking plane — the tool that shaves a rough surface flat and smooth.

[日本語 README](README.md)

| Before | After (`Ctrl+Shift+F`) |
|---|---|
| ![Before](docs/images/before.png) | ![After](docs/images/after.png) |

> **Note:** The user interface (menus, dialogs and messages) is currently **Japanese only**. The formatter itself works with any VBA code. English UI is planned.

## Features

- **A free alternative to Smart Indenter**: focused on the indentation features most people need.
- **Fast and small**: a single ~60 KB DLL. It doesn't slow down Office startup the way Rubberduck can.
- **Works inside the VBE**: run it from the menu bar, the code window's context menu, or a keyboard shortcut.
- **Easy install**: just run `install.bat`. No admin rights needed, and you don't have to enable "Trust access to the VBA project object model".
- **32-bit and 64-bit Office**: works in Excel, Access, Word and any other host with a VBE.
- **Never changes the line count**: only leading whitespace changes, so breakpoints stay where they are.

## Download and install

1. Download `KannaForVBA-x.y.z.zip` from [Releases](https://github.com/alineapps/kanna-vba/releases/latest).
2. Close all Office applications.
3. Right-click the zip → **Properties** → check **Unblock** at the bottom of the General tab → **OK**.
4. Extract the zip and double-click `install.bat`.
5. Start Excel or Access and press `Alt+F11`. A new **整形(J)** ("Format") menu appears in the VBE menu bar.

> If you skip step 3, Windows may show a "Windows protected your PC" warning. Click **More info** → **Run anyway** to continue.

To uninstall, run `uninstall.bat`.

**Requirements**: Windows 10 / 11, .NET Framework 4.8 (built into Windows), Office 2010 or later (including Microsoft 365)

## Usage

| Action | Menu item (Japanese) | Default shortcut |
|---|---|---|
| Format the procedure at the cursor (or the selected lines) | 整形 → プロシージャ/選択範囲を整形 | `Ctrl+Shift+F` |
| Format the current module | 整形 → モジュールを整形 | (assignable) |
| Format every module in the active project | 整形 → プロジェクト全体を整形 | |
| Undo the last formatting | 整形 → 直前の整形を元に戻す | |
| Settings (with live preview) | 整形 → 設定 | |

The same commands are also in the code window's right-click menu.

Note: the VBE's own `Ctrl+Z` can't undo changes made by add-ins. Use "Undo the last formatting" instead.

## What it formats

- Indents `Sub` / `Function` / `Property`, `If` / `ElseIf` / `Else`, `Select Case` / `Case`, `For` / `Next`, `Do` / `Loop`, `While` / `Wend`, `With`, `Type`, `Enum` and `#If` blocks.
- Single-line `If x Then y` statements are not treated as blocks.
- Line labels and line numbers stay in column 1.
- Continuation lines (` _`) keep their position relative to the first line.
- Keywords inside strings and comments are ignored.

You can change the indent width, use tabs, choose whether `Case` is indented, set shortcuts and more in the settings dialog.
The full specification is in [docs/仕様書.md](docs/仕様書.md) (Japanese).

## Building from source

Requirements: .NET SDK 10 or later, and the Office interop assemblies (installed into the GAC with Office).

```powershell
.\build.ps1      # test → Release build → dist\KannaForVBA-x.y.z.zip
```

If the interop assemblies aren't in the GAC, put `Microsoft.Vbe.Interop.dll` and `office.dll` in a folder and build with
`dotnet build src\Kanna.AddIn -c Release -p:OfficeInteropDir=C:\path\to\dlls`.

Bug reports and feature requests are welcome in [Issues](https://github.com/alineapps/kanna-vba/issues). If something gets indented incorrectly, please include a code sample.

## License

[MIT License](LICENSE)
