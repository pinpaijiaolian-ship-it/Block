# Block

This repository includes a PowerShell script that can pin any window on Windows 11 (or other modern Windows versions) so it always stays on top. This is handy for keeping a Sticky Notes window visible while typing in another application.

## Requirements

- Windows 10/11 with PowerShell 5.1 or PowerShell 7+
- Permission to run scripts. You may need to run `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` once in an elevated PowerShell session.

## Usage

1. Open PowerShell.
2. Navigate to the folder containing this repository.
3. Run the script with part of the window title you want to pin:
   ```powershell
   .\scripts\Set-WindowTopMost.ps1 -WindowTitle "Sticky Notes"
   ```
   The command searches for windows whose titles contain the provided text, then marks the first match as always-on-top.

### Additional options

- If multiple windows match, choose a different one with `-Index`:
  ```powershell
  .\scripts\Set-WindowTopMost.ps1 -WindowTitle "Notepad" -Index 2
  ```
- To remove the always-on-top flag from a window:
  ```powershell
  .\scripts\Set-WindowTopMost.ps1 -WindowTitle "Sticky Notes" -Disable
  ```

The script reports the exact window it changed or throws a detailed error if nothing matched so you know what to adjust.
