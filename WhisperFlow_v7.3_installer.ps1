# ============================================================
# WhisperFlow v7.3 - Instalador Automático Aprimorado (Final)
# Autor: Paulo Estimado  
# Data: Outubro 2025
# ============================================================

# Configurações iniciais
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONUTF8 = "1"
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$InformationPreference = "Continue"

# ============================================================
# Configurações
# ============================================================
$CONFIG = @{
    Version           = "7.3"
    BasePath          = "C:\Temp\WhisperFlow"
    LogDir            = "$env:LOCALAPPDATA\WhisperLocal"
    PythonVersion     = "3.13.0"
    PythonUrl         = "https://www.python.org/ftp/python/3.13.0/python-3.13.0-amd64.exe"
    RequiredSpace     = 2 * 1024 * 1024 * 1024  # 2GB
    Hotkey            = "ctrl+alt+j"
    LogFile           = "$env:LOCALAPPDATA\WhisperLocal\installer.log"
    PythonInstallPath = "C:\Program Files\Python313"
}

# ============================================================
# Funções Auxiliares
# ============================================================
function Write-Log {
    param([string]$Message, [string]$Type = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Type] $Message"
    if (-not (Test-Path $CONFIG.LogFile -PathType Leaf)) {
        New-Item -Path (Split-Path $CONFIG.LogFile) -ItemType Directory -Force | Out-Null
    }
    Add-Content -Path $CONFIG.LogFile -Value $logEntry -Encoding UTF8
    Write-Output $logEntry
}

function Write-Status {
    param([string]$Message, [string]$Type = "Info")
    $colors = @{ Info = "Cyan"; Success = "Green"; Warning = "Yellow"; Error = "Red" }
    $icons  = @{ Info = "ℹ️"; Success = "✅"; Warning = "⚠️"; Error = "❌" }
    $color = $colors[$Type]; $icon = $icons[$Type]
    Write-Host "$icon $Message" -ForegroundColor $color
    Write-Log -Message $Message -Type $Type.ToUpper()
}

function Test-DiskSpace {
    $drive = if (Test-Path $CONFIG.BasePath) { (Get-Item $CONFIG.BasePath).PSDrive.Name } else { "C" }
    $freeSpace = (Get-PSDrive $drive).Free
    $reqGB = [math]::Round($CONFIG.RequiredSpace / 1GB, 2)
    $availGB = [math]::Round($freeSpace / 1GB, 2)
    if ($freeSpace -lt $CONFIG.RequiredSpace) {
        throw "Espaço insuficiente. Necessário: $reqGB GB, Disponível: $availGB GB"
    }
    Write-Status "Espaço em disco suficiente ($availGB GB disponíveis no drive $drive)" "Success"
}

function Test-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Status "Executando com privilégios de Administrador" "Success"
    } else {
        Write-Status "Executando sem privilégios de Administrador (algumas etapas podem falhar)" "Warning"
    }
}

function Test-ExistingInstallation {
    if (Test-Path $CONFIG.BasePath) {
        $choice = Read-Host "Instalação existente detectada em $($CONFIG.BasePath). Deseja remover e reinstalar? (S/N)"
        if ($choice -match '^[sS]$') {
            Write-Status "Removendo instalação anterior..." "Info"
            Remove-Item -Path $CONFIG.BasePath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Status "Instalação anterior removida com sucesso" "Success"
        } else {
            Write-Status "Instalação cancelada pelo usuário" "Warning"
            exit 0
        }
    }
}

function Install-Python {
    Write-Status "Verificando instalação do Python..." "Info"
    try {
        $pyv = (python --version 2>$null) -replace "Python ", ""
        if ($pyv -and [version]$pyv -ge [version]$CONFIG.PythonVersion) {
            Write-Status "Python $pyv já instalado e compatível" "Success"
            return
        }
    } catch {}
    Write-Status "Python $($CONFIG.PythonVersion) não encontrado. Instalando..." "Warning"
    $installer = Join-Path $CONFIG.BasePath "python-installer.exe"
    New-Item -ItemType Directory -Force -Path $CONFIG.BasePath | Out-Null
    Invoke-WebRequest -Uri $CONFIG.PythonUrl -OutFile $installer -UseBasicParsing
    Write-Status "Instalando Python..." "Info"
    $args = @("/quiet", "InstallAllUsers=1", "PrependPath=1", "Include_pip=1", "Include_test=0", "Include_tcltk=0")
    Start-Process -FilePath $installer -ArgumentList $args -Wait
    [Environment]::SetEnvironmentVariable("Path", "$env:Path;$($CONFIG.PythonInstallPath);$($CONFIG.PythonInstallPath)\Scripts", "Process")
    Remove-Item $installer -Force -ErrorAction SilentlyContinue
    Write-Status "Python instalado e configurado com sucesso" "Success"
}

function New-VirtualEnvironment {
    Write-Progress -Activity "Instalação do WhisperFlow" -Status "Criando ambiente virtual..." -PercentComplete 30
    Write-Status "Criando ambiente virtual Python..." "Info"
    Push-Location $CONFIG.BasePath
    python -m venv venv
    & .\venv\Scripts\python.exe -m pip install --upgrade pip --quiet
    Pop-Location
    Write-Status "Ambiente virtual criado com sucesso" "Success"
}

function Install-Dependencies {
    Write-Progress -Activity "Instalação do WhisperFlow" -Status "Instalando dependências..." -PercentComplete 50
    Write-Status "Instalando dependências Python..." "Info"
    $pip = "$($CONFIG.BasePath)\venv\Scripts\pip.exe"
    & $pip install --upgrade pip setuptools wheel --quiet
    Write-Status "Instalando PyTorch com suporte CUDA (ou fallback CPU)..." "Info"
    try {
        & $pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu121 --quiet
    } catch {
        Write-Status "Falha CUDA, instalando versão CPU-only..." "Warning"
        & $pip install torch torchvision torchaudio --quiet
    }
    $packages = @("keyboard", "sounddevice", "numpy", "pyperclip", "soundfile", "faster-whisper", "pyaudio")
    foreach ($pkg in $packages) {
        Write-Status "  Instalando $pkg..." "Info"
        try { & $pip install $pkg --quiet --disable-pip-version-check }
        catch { Write-Status "  Aviso: falha ao instalar $pkg" "Warning" }
    }
    Write-Status "Dependências instaladas com sucesso" "Success"
}

function New-ApplicationFiles {
    Write-Progress -Activity "Instalação do WhisperFlow" -Status "Criando arquivos da aplicação..." -PercentComplete 70
    Write-Status "Gerando arquivos Python..." "Info"
    $script:mainPath = Join-Path $CONFIG.BasePath "whisper_flow_v$($CONFIG.Version).pyw"
    @'
#!/usr/bin/env python3.13
"""
WhisperFlow v7.3 - Ditado por voz com Whisper (GPU/CPU)
Autor: Paulo Estimado (Melhorado por Grok)
"""
import os, io, time, queue, threading, tempfile, numpy as np
import sounddevice as sd, soundfile as sf, keyboard, pyperclip, winsound
from faster_whisper import WhisperModel
from datetime import datetime
from pathlib import Path
import logging

CONFIG = {"hotkey": "ctrl+alt+j", "model": "medium", "device": "cuda", "compute_type": "float16"}

LOG_DIR = Path(os.getenv("LOCALAPPDATA")) / "WhisperLocal"
LOG_DIR.mkdir(parents=True, exist_ok=True)
EVENTS_LOG = LOG_DIR / "events.log"
logging.basicConfig(filename=EVENTS_LOG, encoding="utf-8", level=logging.INFO)
def log(msg): open(EVENTS_LOG,"a",encoding="utf-8").write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")
def beep(f=1000,d=120): 
    try: winsound.Beep(f,d)
    except: pass

class Recorder:
    def __init__(self): self.q=queue.Queue(); self.frames=[]; self.rec=False
    def _cb(self,indata,_,__,___): 
        if self.rec: self.q.put(indata.copy())
    def start(self):
        self.frames=[]; self.rec=True
        self.stream=sd.InputStream(samplerate=16000,channels=1,callback=self._cb)
        self.stream.start(); threading.Thread(target=self._collect,daemon=True).start(); beep(900,100); log("🎙️ Gravando...")
    def _collect(self):
        while self.rec:
            try:self.frames.append(self.q.get(timeout=0.2))
            except queue.Empty: pass
    def stop(self):
        self.rec=False
        try:self.stream.stop();self.stream.close()
        except: pass
        if not self.frames: beep(400,200); log("⚠️ Nenhum áudio"); return b""
        audio=np.concatenate(self.frames,axis=0)
        with io.BytesIO() as b:
            with sf.SoundFile(b,"w",samplerate=16000,channels=1,subtype="PCM_16") as f:f.write(audio)
            return b.getvalue()

class Transcriber:
    def __init__(self):
        try:
            self.model=WhisperModel(CONFIG["model"],device=CONFIG["device"],compute_type=CONFIG["compute_type"])
            log("✅ Modelo carregado (GPU)")
        except Exception as e:
            log(f"⚠️ GPU falhou: {e}, usando CPU")
            self.model=WhisperModel(CONFIG["model"],device="cpu",compute_type="int8")
            log("✅ Modelo CPU ativo")
    def transcribe(self,audio):
        if not audio: return ""
        tmp=tempfile.NamedTemporaryFile(suffix=".wav",delete=False)
        tmp.write(audio); tmp.close()
        try:
            seg,_=self.model.transcribe(tmp.name,language="pt",vad_filter=True,beam_size=5)
            txt=" ".join([s.text.strip() for s in seg]); beep(1100,100); log(f"📝 Texto: {txt[:60]}...")
            return txt
        except Exception as e:
            log(f"❌ Erro: {e}"); beep(300,400); return ""
        finally:
            os.remove(tmp.name)

def paste(t):
    if not t: return
    try: keyboard.write(t)
    except: pyperclip.copy(t); keyboard.send("ctrl+v")

def main():
    rec=Recorder(); asr=Transcriber()
    keyboard.add_hotkey(CONFIG["hotkey"], lambda: run(rec,asr))
    print(f"WhisperFlow v7.3 ativo. Use {CONFIG['hotkey'].upper()} para ditar.")
    log("🚀 WhisperFlow iniciado")
    keyboard.wait()

def run(rec,asr):
    rec.start()
    while keyboard.is_pressed(CONFIG["hotkey"]): time.sleep(0.05)
    audio=rec.stop()
    if audio: paste(asr.transcribe(audio))

if __name__=="__main__": main()
'@ | Out-File -FilePath $mainPath -Encoding UTF8
    Write-Status "Arquivos gerados com sucesso" "Success"
}

function New-DesktopShortcut {
    Write-Progress -Activity "Instalação do WhisperFlow" -Status "Criando atalho..." -PercentComplete 90
    Write-Status "Criando atalho na área de trabalho..." "Info"
    $desktop = [Environment]::GetFolderPath("Desktop")
    $shortcut = Join-Path $desktop "WhisperFlow.lnk"
    $wshell = New-Object -ComObject WScript.Shell
    $lnk = $wshell.CreateShortcut($shortcut)
    $lnk.TargetPath = "$($CONFIG.BasePath)\venv\Scripts\pythonw.exe"
    $lnk.Arguments = "`"$script:mainPath`""
    $lnk.WorkingDirectory = $CONFIG.BasePath
    $lnk.IconLocation = "shell32.dll,220"
    $lnk.Description = "WhisperFlow v$($CONFIG.Version) - Ditado por voz com IA"
    $lnk.Save()
    Write-Status "Atalho criado: WhisperFlow.lnk" "Success"
}

function Test-GPUAvailability {
    Write-Progress -Activity "Instalação do WhisperFlow" -Status "Testando GPU..." -PercentComplete 95
    $py = "$($CONFIG.BasePath)\venv\Scripts\python.exe"
    $result = & $py -c "from faster_whisper import WhisperModel; 
try: WhisperModel('tiny', device='cuda', compute_type='float16'); print('GPU_OK')
except Exception as e: print('GPU_FAIL:', e)" 2>&1
    if ($result -match "GPU_OK") {
        Write-Status "GPU CUDA detectada e funcional" "Success"
    } else {
        Write-Status "GPU indisponível, usará CPU. ($result)" "Warning"
    }
}

# ============================================================
# Execução Principal
# ============================================================
try {
    Clear-Host
    Write-Host ("="*65), ("WhisperFlow v7.3 - Instalador Automático"), ("="*65) -ForegroundColor Cyan
    Test-Administrator
    Test-ExistingInstallation
    Test-DiskSpace
    New-Item -ItemType Directory -Force -Path $CONFIG.BasePath, $CONFIG.LogDir | Out-Null
    Install-Python
    New-VirtualEnvironment
    Install-Dependencies
    New-ApplicationFiles
    New-DesktopShortcut
    Test-GPUAvailability
    Write-Progress -Activity "Instalação do WhisperFlow" -Completed
    Write-Host "`n✅ Instalação concluída com sucesso! 🎉`n" -ForegroundColor Green
    Write-Host "👉 Abra o atalho 'WhisperFlow' na área de trabalho e use $($CONFIG.Hotkey.ToUpper()) para ditar." -ForegroundColor Cyan
    Write-Host "📄 Logs: $($CONFIG.LogDir)`n" -ForegroundColor Yellow
} catch {
    Write-Status "Erro: $($_.Exception.Message)" "Error"
    Write-Host "Verifique o log: $($CONFIG.LogFile)" -ForegroundColor Red
    exit 1
}
