 Script Name : acquire-ollama-port.ps1
 Version     : 1.0
 Author      : Tristen Sinanju - Pangolin Atelier
 Date        : 2025-06-13

Purpose:
--------
Automate the management of the Ollama server port on Windows:

- Check if Ollama is running and responding
- Verify that the required model is available
- If not, gracefully free the TCP port (default 11434)
- Stop Windows IP Helper service (iphlpsvc) if needed
- Start Ollama serve
- Restart IP Helper service after Ollama is running

Provides a Unix-friendly CLI interface that works well in PowerShell:

    FreeOllama --help
    FreeOllama -p 11435 -r "mistral-7b" -v

Usage:
------

    FreeOllama [--port <port>] [--required-model <model>] [--verbose] [--help]
    FreeOllama -p <port> -r <model> -v -h

Examples:
---------

    FreeOllama
    FreeOllama -p 11435 -r "llama3.2-vision"
    FreeOllama -r "mistral-7b" -v

Notes:
------

- Uses $args[] parsing to support CLI-style arguments (--flag or -f).
- No PowerShell param() type enforcement is used to avoid binding errors.
- Makes it easy for Unix-native users to control Ollama in PowerShell.

Change Log:
-----------

v1.0 (2025-06-13)
    - Initial version with CLI-style argument parsing
    - Supports short and long flags
    - Added verbose logging and help output
    - Added full section comments and thought process documentation
