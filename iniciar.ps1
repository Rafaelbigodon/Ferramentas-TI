# Configurações
$BaseUrl = "https://raw.githubusercontent.com/Rafaelbigodon/Ferramentas-TI/main"
$TempFolder = Join-Path $env:TEMP "IdeiaInformatica"

# 1. Criar pasta temporária limpa
if (Test-Path $TempFolder) { Remove-Item $TempFolder -Recurse -Force }
New-Item -ItemType Directory -Path $TempFolder | Out-Null

# 2. Lista de arquivos que o seu script original precisa para funcionar
$Arquivos = @(
    "Ferramentas.ps1",
    "matrix.png",
    "Robo.png",
    "theme_switch_sound.wav",
    "robo_sound.wav",
    "Death_note.wav",
    "fuleiro_sound.wav",
    "death_note.png",
    "special_opening_sound.wav",
    "special_opening_image.png",
    "icon_mute.png",
    "icon_unmute.png"
)

Write-Host "Baixando ferramentas IDEIA INFORMATICA..." -ForegroundColor Cyan

# 3. Baixar tudo para a pasta temporária
foreach ($Arquivo in $Arquivos) {
    $Url = "$BaseUrl/$Arquivo"
    $Destino = Join-Path $TempFolder $Arquivo
    try {
        Invoke-WebRequest -Uri $Url -OutFile $Destino -ErrorAction SilentlyContinue
    } catch {
        Write-Host "Aviso: Nao foi possivel baixar $Arquivo" -ForegroundColor Yellow
    }
}

# 4. Ir para a pasta e executar o SEU script original
Set-Location $TempFolder
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Ferramentas.ps1"