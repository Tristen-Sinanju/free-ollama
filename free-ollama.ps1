<#
===========================================================================
 Script Name : acquire-ollama-port.ps1
 Version     : 1.0
 Author      : Tristen Sinanju - Pangolin Atelier
 Date        : 2025-06-13
===========================================================================

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

===========================================================================
#>



function FreeOllama {
    param()

    #---------------------------------------------------
    # Section: Banner
    #---------------------------------------------------
    Write-Host "FreeOllama v1.0 🚀" -ForegroundColor Cyan

    #---------------------------------------------------
    # Section: Defaults
    #---------------------------------------------------
    # These are the default values that can be overridden via CLI args.
    $Port = 11434
    $WaitTimeoutSeconds = 10
    $OllamaStartupDelay = 5
    $RequiredModel = "llama3.2-vision"
    $Help = $false
    $Verbose = $false

    #---------------------------------------------------
    # Section: Argument Parsing
    #---------------------------------------------------
    <#
    We use param() with no parameters to avoid PowerShell trying to bind CLI flags.
    All args come in via $args[], so we can support:

        FreeOllama --help
        FreeOllama -p 11435 -r "mistral-7b" -v

    This mimics Unix CLI behavior.

    We define defaults at the top of the function, and allow user to override via CLI args.
    #>

    for ($i = 0; $i -lt $args.Count; $i++) {
        switch ($args[$i]) {
            "--help" { $Help = $true }
            "-h"     { $Help = $true }

            "--required-model" { if ($i + 1 -lt $args.Count) { $RequiredModel = $args[$i + 1]; $i++ } }
            "-r"               { if ($i + 1 -lt $args.Count) { $RequiredModel = $args[$i + 1]; $i++ } }

            "--port"  { if ($i + 1 -lt $args.Count) { $Port = [int]$args[$i + 1]; $i++ } }
            "-p"      { if ($i + 1 -lt $args.Count) { $Port = [int]$args[$i + 1]; $i++ } }

            "--wait-timeout-seconds" { if ($i + 1 -lt $args.Count) { $WaitTimeoutSeconds = [int]$args[$i + 1]; $i++ } }
            "--ollama-startup-delay" { if ($i + 1 -lt $args.Count) { $OllamaStartupDelay = [int]$args[$i + 1]; $i++ } }

            "--verbose" { $Verbose = $true }
            "-v"        { $Verbose = $true }
        }
    }

    #---------------------------------------------------
    # Section: Verbose logging helper
    #---------------------------------------------------
    function Write-VerboseLog($message) {
        if ($Verbose) {
            $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Write-Host "[$timestamp] $message" -ForegroundColor DarkGray
        }
    }

    #---------------------------------------------------
    # Section: Help
    #---------------------------------------------------
    if ($Help) {
        Write-Host @"
Usage: FreeOllama [--port <port>] [--required-model <model>] [--verbose] [--help]
   or: FreeOllama -p <port> -r <model> -v -h

Defaults:
    --port                 11434
    --wait-timeout-seconds 10
    --ollama-startup-delay 5
    --required-model       llama3.2-vision

Description:
    Checks if Ollama is running and serving the required model.
    If the model is not available, or Ollama is unresponsive, it will:
        - Stop IP Helper (iphlpsvc)
        - Free the port
        - Start Ollama serve
        - Restart IP Helper
"@ -ForegroundColor Cyan
        return
    }

    #---------------------------------------------------
    # Section: Check if model is already available
    #---------------------------------------------------
    Write-Host "`n🚀 Checking if Ollama is already running and serving model '$RequiredModel' on port $Port..." -ForegroundColor Cyan
    Write-VerboseLog "Running: ollama list"

    $ollamaListResult = & ollama list 2>&1

    if ($ollamaListResult -and $ollamaListResult -notmatch "Error") {
        if ($ollamaListResult -match $RequiredModel) {
            Write-Host "✅ Ollama is running and model '$RequiredModel' is available! No need to free port $Port." -ForegroundColor Green
            return
        } else {
            Write-Host "⚠️ Ollama is running but model '$RequiredModel' is NOT available. Proceeding to restart Ollama..." -ForegroundColor Yellow
        }
    } else {
        Write-Host "⚠️ Ollama is not responding on port $Port. Proceeding to free port..." -ForegroundColor Yellow
    }

    #---------------------------------------------------
    # Section: Check if port is used by Ollama
    #---------------------------------------------------
    Write-VerboseLog "Checking NetTCPConnection on port $Port"
    $portConn = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue
    if ($portConn) {
        $proc = Get-Process -Id $portConn.OwningProcess -ErrorAction SilentlyContinue
        if ($proc -and $proc.ProcessName -eq "ollama") {
            Write-Host "⚠️ Ollama process is using port $Port (PID $($proc.Id)). Stopping it..." -ForegroundColor Yellow
            Stop-Process -Id $proc.Id -Force
            Start-Sleep -Seconds 1
        } elseif ($proc) {
            Write-Host "❌ Port $Port is in use by another process: $($proc.ProcessName) (PID $($proc.Id)). Aborting." -ForegroundColor Red
            return
        }
    }

    #---------------------------------------------------
    # Section: Stop IP Helper service
    #---------------------------------------------------
    Write-Host "🔧 Stopping IP Helper service (iphlpsvc)..." -ForegroundColor Yellow
    Write-VerboseLog "Running: Stop-Service iphlpsvc"
    Stop-Service iphlpsvc -Force

    #---------------------------------------------------
    # Section: Wait for port to be free
    #---------------------------------------------------
    $elapsed = 0
    while ($elapsed -lt $WaitTimeoutSeconds) {
        Write-VerboseLog "Checking if port $Port is free (elapsed $elapsed seconds)"
        $portInUse = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue
        if (-not $portInUse) {
            Write-Host "✅ Port $Port is now free." -ForegroundColor Green
            break
        } else {
            Write-Host "⏳ Waiting for port $Port to be free... ($elapsed/$WaitTimeoutSeconds seconds)"
            Start-Sleep -Seconds 1
            $elapsed++
        }
    }

    if ($elapsed -ge $WaitTimeoutSeconds) {
        Write-Host "❌ Timeout: Port $Port is still in use. Aborting." -ForegroundColor Red
        return
    }

    #---------------------------------------------------
    # Section: Start Ollama serve
    #---------------------------------------------------
    Write-Host "🚀 Starting ollama serve..." -ForegroundColor Cyan
    Write-VerboseLog "Running: Start-Process powershell -NoExit -Command 'ollama serve'"
    Start-Process powershell -ArgumentList "-NoExit", "-Command", "ollama serve"

    #---------------------------------------------------
    # Section: Wait for Ollama to start
    #---------------------------------------------------
    Write-Host "⏳ Waiting $OllamaStartupDelay seconds to let Ollama start..."
    Start-Sleep -Seconds $OllamaStartupDelay

    #---------------------------------------------------
    # Section: Restart IP Helper service
    #---------------------------------------------------
    Write-Host "🔄 Restarting IP Helper service (iphlpsvc)..." -ForegroundColor Yellow
    Write-VerboseLog "Running: Start-Service iphlpsvc"
    Start-Service iphlpsvc

    #---------------------------------------------------
    # Section: Conclusion
    #---------------------------------------------------
    Write-Host "`n🎉 Done! Ollama should now be running on port $Port. IP Helper restarted." -ForegroundColor Green
}
