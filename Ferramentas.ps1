# #############################################################################
# # FERRAMENTA DE SISTEMA UNIFICADA (IDEIA INFORMATICA)
# # Versao Grafica (PowerShell + WinForms) - V110 (Log Visível + Ajuste Altura)
# #############################################################################

# --- 1. Verificacao de Administrador ---
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"{0}`"" -f $MyInvocation.MyCommand.Path
    Start-Process powershell -ArgumentList $arguments -Verb RunAs -WindowStyle Normal
    exit
}

# --- 2. Carregar as bibliotecas graficas (WinForms) ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
Import-Module BitsTransfer -ErrorAction SilentlyContinue

# --- V109: Variável Global para o LogBox ---
$script:logBox = $null

# --- ### V109: FUNÇÃO DE LOG CENTRALIZADA ### ---
# Esta função substitui todos os Write-Host e Write-Warning
function Add-Log {
    param (
        [string]$Message,
        [string]$Type = "Info" # Opções: Info, Success, Warning, Error
    )

    $timestamp = (Get-Date -Format "HH:mm:ss")
    $logEntry = "[$timestamp] $Message"

    # 1. Escrever no console (para debug)
    switch ($Type) {
        "Success" { Write-Host $logEntry -ForegroundColor Green }
        "Warning" { Write-Host $logEntry -ForegroundColor Yellow }
        "Error"   { Write-Host $logEntry -ForegroundColor Red }
        default   { Write-Host $logEntry }
    }

    # 2. Adicionar à GUI (RichTextBox)
    if ($script:logBox -ne $null) {
        # Definir cor
        $color = $script:logBox.ForeColor # Cor padrão (depende do tema)
        switch ($Type) {
            "Success" { $color = [System.Drawing.Color]::LawnGreen }
            "Warning" { $color = [System.Drawing.Color]::Yellow }
            "Error"   { $color = [System.Drawing.Color]::Red }
        }

        # Usar Invoke para segurança de thread
        try {
            $script:logBox.Invoke([Action]{
                $script:logBox.SelectionStart = $script:logBox.TextLength
                $script:logBox.SelectionLength = 0
                $script:logBox.SelectionColor = $color
                $script:logBox.AppendText("$logEntry`n") # Adiciona nova linha
                $script:logBox.ScrollToCaret()
            })
        } catch {
            Write-Warning "Falha ao invocar atualização do LogBox: $($_.Exception.Message)"
        }
    }
}


# --- ### Funcoes para Sons dos Botoes (V105: Alterado) ### ---
function Play-HoverSound {
}
function Play-ClickSound {
}

# --- ### V105: Funcao para Reiniciar a Musica de Fundo ### ---
function Resume-BackgroundMusic {
    # Esta função reinicia o loop da música se ela não estiver mutada.
    # Usada após um som (ex: troca de tema) interromper o loop.
    if ($backgroundMusicPlayer -ne $null -and $script:isMuted -eq $false) {
        try {
            $backgroundMusicPlayer.PlayLooping()
        } catch {
            Add-Log "Erro ao reiniciar a música de fundo." -Type "Warning"
        }
    }
}

# --- ### Pre-carregar Sons e Imagens ### ---
$PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

# ### V103: Musica de Fundo (Loop) ###
# !!! IMPORTANTE: System.Media.SoundPlayer NÃO tem controle de volume! !!!
# Para ajustar o volume da música de fundo, você deve editar o arquivo
# 'background_music.wav' em um editor de áudio (ex: Audacity) e salvá-lo
# com um volume mais baixo ANTES de executar este script.
$backgroundMusicPath = "$PSScriptRoot\background_music.wav" # <-- FORNECER .WAV (Com volume ajustado)
$backgroundMusicPlayer = $null
#if (Test-Path $backgroundMusicPath) {
#   try {
#        $backgroundMusicPlayer = New-Object System.Media.SoundPlayer $backgroundMusicPath
#        $backgroundMusicPlayer.Load()
#    } catch {
#        Add-Log "ERRO ao carregar 'background_music.wav': $($_.Exception.Message)" -Type "Error"
#       $backgroundMusicPlayer = $null
#    }
# } else {
#   Add-Log "AVISO: 'background_music.wav' nao encontrado." -Type "Warning"
# }
# ### FIM V103 ###

# V97: Som de Troca de Tema (Unificado)
$themeSwitchSoundPath = "$PSScriptRoot\theme_switch_sound.wav" # <-- FORNECER .WAV
$themeSwitchSoundPlayer = $null
if (Test-Path $themeSwitchSoundPath) {
    try {
        $themeSwitchSoundPlayer = New-Object System.Media.SoundPlayer $themeSwitchSoundPath
        $themeSwitchSoundPlayer.Load()
    } catch {
        Add-Log "ERRO ao carregar 'theme_switch_sound.wav': $($_.Exception.Message)" -Type "Error"
        $themeSwitchSoundPlayer = $null
    }
} else {
    Add-Log "AVISO: 'theme_switch_sound.wav' nao encontrado." -Type "Warning"
}

# Som do Robo
$robotSoundPath = "$PSScriptRoot\robo_sound.wav" # <-- FORNECER .WAV
$robotSoundPlayer = $null
if (Test-Path $robotSoundPath) {
    try {
        $robotSoundPlayer = New-Object System.Media.SoundPlayer $robotSoundPath
        $robotSoundPlayer.Load()
    } catch {
        Add-Log "ERRO ao carregar 'robo_sound.wav': $($_.Exception.Message)" -Type "Error"
        $robotSoundPlayer = $null
    }
} else {
    Add-Log "AVISO: 'robo_sound.wav' nao encontrado." -Type "Warning"
}

# Som Death Note (Easter Egg)
$deathNoteSoundPath = "$PSScriptRoot\Death_note.wav" # <-- FORNECER .WAV
$deathNoteSoundPlayer = $null
if (Test-Path $deathNoteSoundPath) {
    try {
        $deathNoteSoundPlayer = New-Object System.Media.SoundPlayer $deathNoteSoundPath
        $deathNoteSoundPlayer.Load()
    } catch {
        Add-Log "ERRO ao carregar 'Death_note.wav': $($_.Exception.Message)" -Type "Error"
        $deathNoteSoundPlayer = $null
    }
} else {
    Add-Log "AVISO: 'Death_note.wav' nao encontrado para Easter Egg." -Type "Warning"
}

# V101: Som Fuleiro (Easter Egg)
$fuleiroSoundPath = "$PSScriptRoot\fuleiro_sound.wav" # <-- FORNECER .WAV
$fuleiroSoundPlayer = $null
if (Test-Path $fuleiroSoundPath) {
    try {
        $fuleiroSoundPlayer = New-Object System.Media.SoundPlayer $fuleiroSoundPath
        $fuleiroSoundPlayer.Load()
    } catch {
        Add-Log "ERRO ao carregar 'fuleiro_sound.wav': $($_.Exception.Message)" -Type "Error"
        $fuleiroSoundPlayer = $null
    }
} else {
    Add-Log "AVISO: 'fuleiro_sound.wav' nao encontrado para Easter Egg." -Type "Warning"
}

# Imagem Death Note (Easter Egg)
$deathNoteImagePath = "$PSScriptRoot\death_note.png" # <-- FORNECER .PNG
$deathNoteBGImage = $null
if (Test-Path $deathNoteImagePath) {
    try {
        $deathNoteBGImage = [System.Drawing.Image]::FromFile($deathNoteImagePath)
    } catch {
        Add-Log "ERRO ao carregar 'death_note.png': $($_.Exception.Message)" -Type "Error"
        $deathNoteBGImage = $null
    }
} else {
    Add-Log "Imagem 'death_note.png' não encontrada para Easter Egg." -Type "Warning"
}

# V98: Imagem e Som de Abertura Rara (Easter Egg)
$specialOpeningSoundPath = "$PSScriptRoot\special_opening_sound.wav" # <-- FORNECER .WAV
$specialOpeningSoundPlayer = $null
if (Test-Path $specialOpeningSoundPath) {
    try {
        $specialOpeningSoundPlayer = New-Object System.Media.SoundPlayer $specialOpeningSoundPath
        $specialOpeningSoundPlayer.Load()
    } catch {
        Add-Log "ERRO ao carregar 'special_opening_sound.wav': $($_.Exception.Message)" -Type "Error"
        $specialOpeningSoundPlayer = $null
    }
} else {
    Add-Log "AVISO: 'special_opening_sound.wav' nao encontrado para Easter Egg." -Type "Warning"
}

$specialOpeningImagePath = "$PSScriptRoot\special_opening_image.png" # <-- FORNECER .PNG
$specialOpeningBGImage = $null
if (Test-Path $specialOpeningImagePath) {
    try {
        $specialOpeningBGImage = [System.Drawing.Image]::FromFile($specialOpeningImagePath)
    } catch {
        Add-Log "ERRO ao carregar 'special_opening_image.png': $($_.Exception.Message)" -Type "Error"
        $specialOpeningBGImage = $null
    }
} else {
    Add-Log "Imagem 'special_opening_image.png' não encontrada para Easter Egg." -Type "Warning"
}

# ### V103: Imagens do Botão de Mudo ###
$muteImagePath = "$PSScriptRoot\icon_mute.png" # <-- FORNECER .PNG
$muteImage = $null
if (Test-Path $muteImagePath) {
    try { $muteImage = [System.Drawing.Image]::FromFile($muteImagePath) } catch { Add-Log "ERRO ao carregar 'icon_mute.png'" -Type "Error" }
} else { Add-Log "AVISO: 'icon_mute.png' nao encontrado." -Type "Warning" }

$unmuteImagePath = "$PSScriptRoot\icon_unmute.png" # <-- FORNECER .PNG
$unmuteImage = $null
if (Test-Path $unmuteImagePath) {
    try { $unmuteImage = [System.Drawing.Image]::FromFile($unmuteImagePath) } catch { Add-Log "ERRO ao carregar 'icon_unmute.png'" -Type "Error" }
} else { Add-Log "AVISO: 'icon_unmute.png' nao encontrado." -Type "Warning" }
# ### FIM V103 ###


# --- ### Variaveis Globais / Script Level ### ---
$script:creditClicks = 0 # Contador para Easter Egg dos Créditos
$script:isClosing = $false # Flag para controlar animação de fecho
$script:currentTheme = 'dark' # V90: Controlar tema atual
$script:playSpecialOpening = (Get-Random -Minimum 1 -Maximum 101) -eq 1 # V100: 1 em 100 de chance
$script:isMuted = $false # V103: Controla a musica de fundo


# --- Som de Abertura / Saudação (V109: Logs alterados) ---
if ($script:playSpecialOpening -eq $false) {
    # Abertura Normal (99% das vezes)
    Add-Log "============================================================"
    Add-Log "Bem-vindo ao Assistente de ferramentas IDEIA INFORMATICA!"
    Add-Log "============================================================"
    Start-Sleep -Seconds 1
} else {
    # Abertura Rara (1% das vezes)
    Add-Log "============================================================"
    Add-Log "!!! ABERTURA RARA ATIVADA !!!" -Type "Warning"
    Add-Log "============================================================"
}

# #############################################################################
# # FUNÇÕES DOS BOTÕES (1 a 14) 
# #############################################################################

# #############################################################################
# # (BOTAO 1) ACAO: Mudar Nome do Computador E Nome Completo do Usuario (V108)
# #############################################################################
function Start-RenameComputer {
    # --- Parte 1: Renomear Computador ---
    try {
        $currentComputerName = $env:COMPUTERNAME
        Add-Log "Etapa 1: Renomear Computador. Nome atual: $currentComputerName"
        [System.Windows.Forms.Application]::DoEvents() # V110: Força atualização do log
        
        $newComputerName = [Microsoft.VisualBasic.Interaction]::InputBox(
            "ETAPA 1/2: Por favor, insira o novo nome para este COMPUTADOR:",
            "Mudar Nome do Computador",
            $currentComputerName
        )

        # Verifica se o utilizador clicou OK e se o nome é valido
        if ($newComputerName -and $newComputerName -ne $currentComputerName -and $newComputerName -notmatch " ") {
            Add-Log "Tentando renomear o computador de '$currentComputerName' para '$newComputerName'..."
            [System.Windows.Forms.Application]::DoEvents() # V110
            
            Rename-Computer -NewName $newComputerName -Force -ErrorAction Stop
            Add-Log "SUCESSO: Nome do COMPUTADOR alterado para '$newComputerName'. Reinicie para aplicar." -Type "Success"
            [System.Windows.Forms.MessageBox]::Show(
                ("O nome do COMPUTADOR foi alterado para '$newComputerName'.`n`nÉ necessário REINICIAR o computador para que esta alteração tenha efeito."),
                "Nome do Computador Alterado",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } elseif ($newComputerName -and $newComputerName -eq $currentComputerName) {
            Add-Log "Nome do computador não alterado (igual ao atual)." -Type "Info"
        } elseif ($newComputerName -and $newComputerName -match " ") {
            Add-Log "Nome do computador não pode conter espaços." -Type "Warning"
             [System.Windows.Forms.MessageBox]::Show("O nome do computador não pode conter espaços.", "Erro: Nome Inválido", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        } elseif (-not $newComputerName) {
            Add-Log "Ação de renomear cancelada pelo utilizador." -Type "Warning"
            Return
        }

    } catch {
        Add-Log "ERRO ao renomear o COMPUTADOR: $($_.Exception.Message)" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show(
            ("Ocorreu um erro ao tentar renomear o COMPUTADOR:`n" + $_.Exception.Message),
            "Erro ao Renomear Computador",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }

    # --- Parte 2: Renomear Nome Completo do Usuário Atual ---
    Add-Log "Etapa 2: Iniciando alteração do Nome Completo do usuário..."
    [System.Windows.Forms.Application]::DoEvents() # V110
    
    try {
        # Obter o nome de usuário atual (sem domínio/nome do pc)
        $currentUserNameOnly = (Get-WmiObject -Class Win32_ComputerSystem).UserName.Split('\')[-1]
        
        # V108: Usando Get-CimInstance (compatível com Win7+) em vez de Get-LocalUser
        $currentUserObject = Get-CimInstance -ClassName Win32_UserAccount -Filter "LocalAccount = True AND Name = '$currentUserNameOnly'" -ErrorAction Stop
        $currentFullName = $currentUserObject.FullName
        Add-Log "Usuário atual: '$currentUserNameOnly'. Nome completo atual: '$currentFullName'"

        # Perguntar o novo Nome Completo
        $newUserFullName = [Microsoft.VisualBasic.Interaction]::InputBox(
            ("ETAPA 2/2: Insira o novo NOME COMPLETO para o usuário atual '$currentUserNameOnly':`n" +
             "(Este é o nome exibido na tela de login/configurações)"),
            "Mudar Nome Completo do Usuário",
            $currentFullName # Preenche com o nome atual
        )

        # Verifica se clicou OK e se o nome mudou
        if ($newUserFullName -ne $null -and $newUserFullName -ne $currentFullName) {
            Add-Log "Tentando alterar o Nome Completo de '$currentUserNameOnly' para '$newUserFullName'..."
            [System.Windows.Forms.Application]::DoEvents() # V110
            
            # V108: Usando 'net user' (compatível com Win7+) em vez de Set-LocalUser
            net user $currentUserNameOnly /fullname:"$newUserFullName"
            
            Add-Log "SUCESSO: Nome Completo do usuário '$currentUserNameOnly' alterado para '$newUserFullName'." -Type "Success"
            [System.Windows.Forms.MessageBox]::Show(
                "O Nome Completo do usuário foi alterado com sucesso para '$newUserFullName'.`n`nA alteração pode ser visível após reiniciar a sessão ou o computador.",
                "Nome Completo Alterado",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } elseif ($newUserFullName -ne $null -and $newUserFullName -eq $currentFullName) {
             Add-Log "Nome Completo do usuário não alterado (igual ao atual)." -Type "Info"
        } else { # $newUserFullName é $null (usuário clicou Cancelar na segunda caixa)
             Add-Log "Alteração do Nome Completo cancelada pelo utilizador." -Type "Warning"
        }

    } catch {
        Add-Log "ERRO ao alterar o Nome Completo do usuário: $($_.Exception.Message)" -Type "Error"
        # V108: Mensagem de erro atualizada
        if ($_.Exception.ToString() -like "*Get-CimInstance*") {
             [System.Windows.Forms.MessageBox]::Show(
                ("Não foi possível encontrar o usuário '$currentUserNameOnly' para alterar o Nome Completo.`n" +
                 "Esta função pode não funcionar corretamente para contas de domínio ou contas da Microsoft."),
                "Erro ao Alterar Nome Completo",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
             )
        } else {
             [System.Windows.Forms.MessageBox]::Show(
                ("Ocorreu um erro ao tentar alterar o Nome Completo do usuário:`n" + $_.Exception.Message),
                "Erro ao Alterar Nome Completo",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
             )
        }
    }
}


# #############################################################################
# # (BOTAO 2) ACAO: Mostrar Icones Desktop
# #############################################################################
function Start-ShowDesktopIcons {
    try {
        Add-Log "Mostrando icones classicos do Desktop..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        $RegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
        if (-not (Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
        Set-ItemProperty -Path $RegPath -Name '{20D04FE0-3AEA-1069-A2D8-08002B30309D}' -Value 0 -Type DWord -ErrorAction Stop # Computador
        Set-ItemProperty -Path $RegPath -Name '{59031a47-3f72-44a7-89c5-5595fe6b30ee}' -Value 0 -Type DWord -ErrorAction Stop # Ficheiros do Utilizador
        Set-ItemProperty -Path $RegPath -Name '{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}' -Value 0 -Type DWord -ErrorAction Stop # Rede
        Add-Log "SUCESSO: Icones do Desktop definidos para serem mostrados." -Type "Success"
        
        Add-Log "Reiniciando explorer.exe..." -Type "Info"
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue; Start-Sleep -Seconds 1; Start-Process explorer -ErrorAction SilentlyContinue
        [System.Windows.Forms.MessageBox]::Show(
            "Os icones 'Computador', 'Ficheiros do Utilizador' e 'Rede' foram definidos para aparecer na Area de Trabalho.`n`nPode ser necessario atualizar a Area de Trabalho (F5) ou reiniciar o Explorer.",
            "Icones do Desktop",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } catch {
        Add-Log "ERRO ao definir icones do Desktop: $($_.Exception.Message)" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Erro nos Icones do Desktop", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# #############################################################################
# # (BOTAO 3) ACAO: Definir Energia
# #############################################################################
function Start-SetPowerNever {
    try {
        Add-Log "Definindo plano de energia para Alto Desempenho..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        $HighPerfGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
        powercfg /setactive $HighPerfGuid
        Add-Log "Plano de energia ativo definido para Alto Desempenho." -Type "Info"

        Add-Log "Definindo configuracoes de energia para 'Nunca' (Tomada e Bateria)..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        # Na Tomada (AC)
        powercfg /change monitor-timeout-ac 0
        powercfg /change standby-timeout-ac 0
        # Na Bateria (DC)
        powercfg /change monitor-timeout-dc 0
        powercfg /change standby-timeout-dc 0
        Add-Log "SUCESSO: Configuracoes de energia aplicadas." -Type "Success"

        [System.Windows.Forms.MessageBox]::Show(
            "O plano de energia foi definido para 'Alto Desempenho'.`n`nAs configuracoes de energia (Monitor e Suspensao) foram definidas para 'Nunca' tanto na TOMADA como na BATERIA.",
            "Configuracoes de Energia",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } catch {
        Add-Log "ERRO ao definir configuracoes de energia: $($_.Exception.Message)" -Type "Error"
        if ($_.Exception.Message -like '*powercfg /setactive*') {
             [System.Windows.Forms.MessageBox]::Show("Nao foi possivel definir o plano para Alto Desempenho. O plano pode nao existir nesta versao do Windows.", "Erro de Plano de Energia", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        } else {
             [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Erro nas Configuracoes de Energia", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}
# #############################################################################
# # (BOTAO 4) ACAO: Desativar Inicializacao Rapida
# #############################################################################
function Start-DisableFastStartup {
    try {
        Add-Log "Desativando a Inicializacao Rapida..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        $RegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
        if(Test-Path $RegPath) {
             Set-ItemProperty -Path $RegPath -Name "HiberbootEnabled" -Value 0 -Type DWord -ErrorAction Stop
             Add-Log "SUCESSO: Inicializacao Rapida desativada." -Type "Success"
             [System.Windows.Forms.MessageBox]::Show(
                 "A Inicializacao Rapida foi desativada com sucesso.`n`nE necessario REINICIAR o computador para que a alteracao tenha efeito.",
                 "Inicializacao Rapida",
                 [System.Windows.Forms.MessageBoxButtons]::OK,
                 [System.Windows.Forms.MessageBoxIcon]::Information
             )
        } else {
            Add-Log "AVISO: Caminho do Registo para Inicializacao Rapida nao encontrado: $RegPath" -Type "Warning"
             [System.Windows.Forms.MessageBox]::Show(
                 "Nao foi possivel encontrar a chave do registo para desativar a Inicializacao Rapida.",
                 "Erro",
                 [System.Windows.Forms.MessageBoxButtons]::OK,
                 [System.Windows.Forms.MessageBoxIcon]::Warning
             )
        }
    } catch {
        Add-Log "ERRO ao desativar a Inicializacao Rapida: $($_.Exception.Message)" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Erro na Inicializacao Rapida", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# #############################################################################
# # (BOTAO 5) ACAO: Diagnostico de Drivers (dxdiag, devmgmt)
# #############################################################################
function Start-DriverDiagnostics {
    try {
        Add-Log "Abrindo Ferramentas de Diagnostico de Drivers..."
        [System.Windows.Forms.Application]::DoEvents() # V110

        Add-Log "Abrindo Gestor de Dispositivos (devmgmt.msc)..."
        Start-Process "devmgmt.msc"

        Add-Log "Abrindo Ferramenta de Diagnostico do DirectX (dxdiag.exe)..."
        Start-Process "dxdiag.exe"

        Add-Log "Tentando abrir a Camera..."
        try {
            Start-Process "microsoft.windows.camera:" -ErrorAction Stop
        } catch {
            Add-Log "AVISO: Nao foi possivel abrir a app 'Camera'. Pode nao estar instalada." -Type "Warning"
        }

        [System.Windows.Forms.MessageBox]::Show(
            ("As ferramentas de diagnostico foram abertas:`n`n" +
             "1. GESTOR DE DISPOSITIVOS: Procure por icones '!' amarelos (problemas de driver).`n" +
             "2. DXDIAG: Verifique as abas 'Ecra' e 'Som' para problemas reportados.`n" +
             "3. CAMERA: Verifique se a imagem aparece corretamente."),
            "Diagnostico de Drivers",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

    } catch {
        Add-Log "ERRO ao abrir ferramentas de diagnostico: $($_.Exception.Message)" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Erro no Diagnostico", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# #############################################################################
# # (BOTAO 6) ACAO: Configurar Impressora (V108: Compatibilidade Total)
# #############################################################################
function Start-PrinterScript {
    try {
        Add-Log "Iniciando Assistente de Impressora..."
        [System.Windows.Forms.Application]::DoEvents() # V110

        # --- Etapa 1: Criar Utilizador (REVISADO V108) ---
        Add-Log "[ETAPA 1/4] Verificando/Criando utilizador 'Impressora'..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        $Username = "Impressora"
        # V108: Senha mais complexa para evitar falha de política de segurança (ex: Win10/11 Pro)
        $PasswordString = "Impressora@123" 
        
        # 1. Verificar se o usuário já existe (método compatível com Win7+)
        $userExists = $false
        $netUserTest = net user $Username 2>&1
        if ($LASTEXITCODE -eq 0) {
            $userExists = $true
        }

        if ($userExists) {
            Add-Log "AVISO: Utilizador '$Username' já existe. Pulando criação." -Type "Warning"
        } else {
            # 2. Se não existe, tentar criar
            Add-Log "Utilizador '$Username' não encontrado. Tentando criar..."
            [System.Windows.Forms.Application]::DoEvents() # V110
            
            try {
                # V108: Usando 'net user' para máxima compatibilidade (Win10/11 Home/Pro, Win7)
                net user $Username $PasswordString /add /comment:"Utilizador para acesso a impressoras compartilhadas" /expires:never /passwordchg:no
                
                if ($LASTEXITCODE -ne 0) {
                    # Se net user falhar (ex: $LASTEXITCODE é 2245 para política de senha)
                    throw "Falha ao executar 'net user'. Código de saída: $LASTEXITCODE. Verifique a política de senhas."
                }
                
                Add-Log "SUCESSO: Utilizador '$Username' criado com a senha '$PasswordString'." -Type "Success"
                [System.Windows.Forms.MessageBox]::Show(
                    "O usuário 'Impressora' foi criado com sucesso.`n`nSenha: $PasswordString`n`nAnote esta senha, ela pode ser necessária para o compartilhamento.",
                    "Usuário Criado",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )

            } catch {
                # Se 'net user' falhar (ex: política de senha)
                Add-Log "ERRO CRÍTICO ao criar usuário: $($_.Exception.Message)" -Type "Error"
                [System.Windows.Forms.MessageBox]::Show(
                    ("Não foi possível criar o usuário 'Impressora'.`n`nERRO: " + $_.Exception.Message + "`n`nVerifique se a política de complexidade de senha do Windows está bloqueando a criação. O script tentou usar a senha '$PasswordString'."),
                    "Erro na Criação do Usuário",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return # V108: Parar a função aqui, não adianta continuar
            }
        }

        # --- Etapa 2: Aplicar Correcao RPC ---
        Add-Log "[ETAPA 2/4] Aplicando correcao de compartilhamento (RPC)..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        try {
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Print" -Name "RpcAuthnLevelPrivacyEnabled" -Value 0 -Type DWord -ErrorAction Stop
            Add-Log "SUCESSO: Chave 'RpcAuthnLevelPrivacyEnabled' definida para 0." -Type "Success"
        } catch {
             Add-Log "ERRO ao definir chave RPC: $($_.Exception.Message)" -Type "Error"
        }

        # --- Etapa 3: Permissoes do Registo (HKCU) ---
        Add-Log "[ETAPA 3/4] Definindo permissoes do Registo (HKCU)..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        try {
            $path = 'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows'
            if(Test-Path $path) {
                $acl = Get-Acl -Path $path
                $rule1 = New-Object System.Security.AccessControl.RegistryAccessRule('Todos', 'FullControl', 'ContainerInherit', 'None', 'Allow')
                $rule2 = New-Object System.Security.AccessControl.RegistryAccessRule('Convidados', 'FullControl', 'ContainerInherit', 'None', 'Allow')
                $acl.SetAccessRule($rule1)
                $acl.SetAccessRule($rule2)
                Set-Acl -Path $path -AclObject $acl -ErrorAction Stop
                Add-Log "SUCESSO: Permissoes do Registo (HKCU) definidas." -Type "Success"
            } else {
                 Add-Log "AVISO: Caminho do Registo HKCU nao encontrado: $path" -Type "Warning"
            }
        } catch {
             Add-Log "ERRO ao definir permissoes HKCU: $($_.Exception.Message)" -Type "Error"
        }

        # --- Etapa 4: Selecao Grafica e Compartilhamento ---
        Add-Log "[ETAPA 4/4] Abrindo selecao de impressora..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        $choice = Get-Printer | Select-Object Name | Out-GridView -Title 'Selecione a impressora e clique OK' -PassThru

        if ($null -eq $choice) { throw 'Nenhuma impressora foi selecionada. Script cancelado.' }

        $PrinterName = $choice.Name
        Add-Log "Voce selecionou: $PrinterName" -Type "Info"
        Set-Printer -Name $PrinterName -Shared $true -ShareName $PrinterName -ErrorAction Stop
        Add-Log "SUCESSO: Impressora '$PrinterName' compartilhada." -Type "Success"

        # --- Etapa Final: Instrucoes Manuais ---
        [System.Windows.Forms.MessageBox]::Show(
            "SUCESSO! A impressora foi compartilhada.`n`nA proxima etapa (permissoes de seguranca) deve ser feita manualmente (devido a limitacoes do seu Windows):`n`n1. Clique com o botao direito em '$PrinterName'.`n2. Selecione 'Propriedades da impressora'.`n3. Clique na aba 'Seguranca' e defina as permissoes.`n`nA janela de Impressoras sera aberta.",
            "Processo Concluido",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        start ms-settings:printers

    } catch {
        Add-Log "ERRO: $($_.Exception.Message)" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Erro no Script da Impressora", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# #############################################################################
# # (BOTAO 7) ACAO: Reiniciar Spooler
# #############################################################################
function Start-RestartSpooler {
    try {
        Add-Log "Reiniciando o Spooler de Impressao..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        Restart-Service Spooler -Force -ErrorAction Stop
        Add-Log "SUCESSO: Spooler de Impressao reiniciado." -Type "Success"
        [System.Windows.Forms.MessageBox]::Show(
            "O servico Spooler de Impressao foi reiniciado com sucesso.",
            "Spooler Reiniciado",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } catch {
        Add-Log "ERRO ao reiniciar o Spooler: $($_.Exception.Message)" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Erro ao Reiniciar Spooler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# #############################################################################
# # (BOTAO 8) ACAO: Abrir Ninite.com
# #############################################################################
function Start-OpenNinite {
    $url = "https://ninite.com"
    [System.Windows.Forms.MessageBox]::Show(
        ("O site Ninite.com sera aberto no seu navegador padrao.`nUse-o para selecionar e instalar aplicativos essenciais rapidamente."),
        "Abrir Ninite.com",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    try {
        Add-Log "Abrindo Ninite.com no navegador..."
        Start-Process $url -ErrorAction Stop
        Add-Log "SUCESSO: Link Ninite aberto." -Type "Success"
    } catch {
        Add-Log "ERRO ao abrir Ninite.com: $($_.Exception.Message)" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show(("Ocorreu um erro ao tentar abrir o Ninite:`n" + $_.Exception.Message), "Erro ao Abrir Link", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# #############################################################################
# # (BOTAO 9) ACAO: Abrir Link de Download do Office 365
# #############################################################################
function Start-OpenOfficeDownloadLink {
    # --- URL para Instalador ONLINE ---
    $url = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x64&language=pt-br&version=O16GA"
    # Mensagem inicial
    [System.Windows.Forms.MessageBox]::Show(
        ("O seu navegador padrao sera aberto no link de download do instalador ONLINE do Office 365 (PT-BR x64).`n`nURL: " + $url + "`n`nClique OK para continuar."),
        "Abrir Link de Download",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    try {
        Add-Log "Abrindo URL de download do Office 365 no navegador: $url"
        Start-Process $url -ErrorAction Stop
        Add-Log "SUCESSO: Link aberto no navegador." -Type "Success"
    } catch {
        Add-Log "ERRO ao abrir o URL no navegador: $($_.Exception.Message)" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show(("Ocorreu um erro ao tentar abrir o link no seu navegador:`n" + $_.Exception.Message), "Erro ao Abrir Link", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# #############################################################################
# # (BOTAO 10) ACAO: Criar Atalhos Office
# #############################################################################
function Start-CreateOfficeShortcuts {
    $ErrorMessages = @()
    $SuccessMessages = @()
    $DesktopPath = [Environment]::GetFolderPath("Desktop")
    # Tenta encontrar os executaveis em locais comuns
    $OfficePaths = @( "$env:ProgramFiles\Microsoft Office\root\Office16", "$env:ProgramFiles(x86)\Microsoft Office\root\Office16", "$env:ProgramFiles\Microsoft Office\Office16", "$env:ProgramFiles(x86)\Microsoft Office\Office16" )
    $Apps = @{ "WINWORD.EXE" = "Microsoft Word"; "EXCEL.EXE" = "Microsoft Excel"; "POWERPNT.EXE" = "Microsoft PowerPoint" }
    $WshShell = New-Object -ComObject WScript.Shell
    Add-Log "Criando atalhos do Office na Area de Trabalho..."
    [System.Windows.Forms.Application]::DoEvents() # V110
    
    foreach ($exeName in $Apps.Keys) {
        $foundPath = $null
        foreach ($officePath in $OfficePaths) { $fullPath = Join-Path -Path $officePath -ChildPath $exeName; if (Test-Path $fullPath) { $foundPath = $fullPath; break } }
        if ($foundPath) {
            $linkName = $Apps[$exeName] + ".lnk"; $linkPath = Join-Path -Path $DesktopPath -ChildPath $linkName
            try { $shortcut = $WshShell.CreateShortcut($linkPath); $shortcut.TargetPath = $foundPath; $shortcut.Save(); $Msg = "Atalho para $($Apps[$exeName]) criado com sucesso."; Add-Log $Msg -Type "Success"; $SuccessMessages += $Apps[$exeName] }
            catch { $ErrorMessage = "ERRO ao criar atalho para $($Apps[$exeName]): $($_.Exception.Message)"; Add-Log $ErrorMessage -Type "Error"; $ErrorMessages += $Apps[$exeName] }
        } else { $ErrorMessage = "AVISO: Executavel $($exeName) nao encontrado nos caminhos padrao."; Add-Log $ErrorMessage -Type "Warning"; $ErrorMessages += $Apps[$exeName] }
    }
    # --- Mensagem Final ---
    $FinalMessage = ""; if ($SuccessMessages.Count -gt 0) { $FinalMessage += "Atalhos criados com sucesso para: " + ($SuccessMessages -join ', ') + ".`n`n" }; if ($ErrorMessages.Count -gt 0) { $FinalMessage += "Falha ao criar/encontrar atalhos para: " + ($ErrorMessages -join ', ') + "." }
    if ($FinalMessage) { if ($ErrorMessages.Count -gt 0) { [System.Windows.Forms.MessageBox]::Show($FinalMessage, "Resultado Atalhos Office", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) } else { [System.Windows.Forms.MessageBox]::Show($FinalMessage, "Resultado Atalhos Office", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) } }
    else { [System.Windows.Forms.MessageBox]::Show("Nenhuma acao realizada para os atalhos do Office (Office nao encontrado?).", "Atalhos Office", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) }
}

# #############################################################################
# # (BOTAO 11) ACAO: Corrigir Compartilhamento de Rede (SMB)
# #############################################################################
function Start-SmbScript {
    # --- AVISO DE SEGURANCA ---
    $aviso = "AVISO DE SEGURANCA GRAVE:`n`nVoce esta prestes a desativar protecoes fundamentais do Windows (assinatura SMB e bloqueio de logon de convidado).`n`nIsto torna o PC vulneravel a ataques de rede e ransomware.`n`nUse apenas se confiar 100% na sua rede privada.`n`nDeseja continuar?"
    $resultado = [System.Windows.Forms.MessageBox]::Show($aviso, "Confirmacao de Seguranca", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($resultado -eq 'No') {
        Add-Log "Acao Corrigir Rede (SMB) cancelada pelo utilizador." -Type "Warning"
        [System.Windows.Forms.MessageBox]::Show("Acao cancelada pelo utilizador.", "Cancelado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    # --- Execucao do Script ---
    try {
        Add-Log "Iniciando Correcao de Compartilhamento de Rede (SMB)..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        Add-Log "[ETAPA 1/2] Desativando assinatura SMB..."
        Set-SmbClientConfiguration -RequireSecuritySignature $false -Confirm:$false -ErrorAction Stop
        Add-Log "[ETAPA 2/2] Ativando logons de convidado inseguros..."
        Set-SmbClientConfiguration -EnableInsecureGuestLogons $true -Confirm:$false -ErrorAction Stop
        Add-Log "SUCESSO: Configuracoes inseguras aplicadas." -Type "Error" # Visto como um log de segurança
        $msgFinal = "CONCLUIDO. A SEGURANCA DA SUA REDE FOI REDUZIDA.`n`nE necessario REINICIAR o computador agora para que as alteracoes tenham efeito.`n`nDeseja reiniciar agora?"
        $resultadoFinal = [System.Windows.Forms.MessageBox]::Show($msgFinal, "Reiniciar Agora?", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($resultadoFinal -eq 'Yes') {
            Add-Log "Reiniciando o computador..." -Type "Warning"
            Restart-Computer -Force
        }
    } catch {
        Add-Log "ERRO: $($_.Exception.Message)" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Erro no Script SMB", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# #############################################################################
# # (BOTAO 12) ACAO: Ativar Windows e Office
# #############################################################################
function Start-NewButtonAction {
    try {
        Add-Log "Iniciando Ativador..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        # --- ### COLOQUE O SEU CODIGO POWERSHELL AQUI ### ---
        irm https:\\get.activated.win | iex
        # --- Fim do seu comando ---
    } catch {
        Add-Log "ERRO na Acao do Botao 12: $($_.Exception.Message)" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Erro na Ativacao", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# #############################################################################
# # (BOTAO 13) ACAO: Desativar BitLocker (Lab)
# #############################################################################
function Start-DisableBitLocker {
    # --- AVISO DE PERIGO ---
    $aviso = "AVISO DE PERIGO (APENAS LABORATORIO):`n`nVoce esta prestes a DESATIVAR permanentemente o BitLocker na unidade C:.`n`nIsto iniciara a DESCRIPTOGRAFIA total do disco.`n`nEste processo e IRREVERSIVEL e pode demorar MUITO TEMPO.`n`nDeseja continuar?"
    $resultado = [System.Windows.Forms.MessageBox]::Show($aviso, "Confirmacao de Descriptografia", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Error)
    if ($resultado -eq 'No') {
        Add-Log "Acao Desativar BitLocker cancelada pelo utilizador." -Type "Warning"
        [System.Windows.Forms.MessageBox]::Show("Acao cancelada pelo utilizador.", "Cancelado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    try {
        Add-Log "Iniciando desativacao do BitLocker..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        Add-Log "Verificando status do BitLocker em C:..."
        $status = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue

        if ($null -eq $status) {
            throw "Nao foi possivel obter o status do BitLocker. A unidade C: existe ou o BitLocker esta instalado?"
        }

        if ($status.ProtectionStatus -eq 'Off') {
            Add-Log "AVISO: BitLocker ja esta desativado em C:." -Type "Warning"
            [System.Windows.Forms.MessageBox]::Show(
                "O BitLocker ja esta desativado na unidade C:.",
                "BitLocker Desativado",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }

        Add-Log "Iniciando desativacao (descriptografia) do BitLocker em C:..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        Disable-BitLocker -MountPoint "C:" -ErrorAction Stop

        Add-Log "SUCESSO: Descriptografia iniciada." -Type "Success"
        [System.Windows.Forms.MessageBox]::Show(
            "SUCESSO! A descriptografia do BitLocker foi iniciada na unidade C:.`n`nO processo continuara em segundo plano e pode demorar muito tempo (horas).`n`nO BitLocker sera permanentemente desativado quando a descriptografia estiver concluida.",
            "Descriptografia Iniciada",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

    } catch {
        Add-Log "ERRO ao desativar o BitLocker: $($_.Exception.Message)" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Erro ao Desativar BitLocker", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# #############################################################################
# # (BOTAO 14) ACAO: Instalar Win 11 (Bypass) - (NOVO V102)
# #############################################################################
function Start-InstallWin11 {
    # 1. Mensagem de Confirmação (como o 'pause' do .bat)
    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
        ("ATENÇÃO: Esta ação tentará iniciar a instalação do Windows 11 (ignorando os requisitos).`n`n" +
         "Certifique-se de que a imagem ISO do Windows 11 já está montada (aparecendo como uma unidade de CD/DVD, ex: D: ou E:).`n`n" +
         "Deseja continuar?"),
        "Instalador do Windows 11 (Bypass)",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($confirmResult -eq 'No') {
        Add-Log "Instalação do Win 11 cancelada pelo usuário." -Type "Warning"
        return
    }

    # 2. Encontrar a Unidade da ISO (Lógica do 'for' loop)
    Add-Log "Verificando a unidade da ISO do Windows 11..."
    [System.Windows.Forms.Application]::DoEvents() # V110
    
    $isoDrive = $null
    # Get-Volume é mais confiável para encontrar drives de CD/DVD (ISOs montadas)
    $cdDrives = Get-Volume | Where-Object { $_.DriveType -eq 'CD-ROM' -and $_.DriveLetter }

    foreach ($drive in $cdDrives) {
        $driveLetter = $drive.DriveLetter
        # O .bat procurava por setup.exe na raiz
        $setupPath = "${driveLetter}:\setup.exe"
        if (Test-Path $setupPath) {
            $isoDrive = "${driveLetter}:"
            Add-Log "SUCESSO: Unidade da ISO detectada: $isoDrive" -Type "Success"
            break # Encontramos, parar de procurar
        }
    }

    # 3. Lidar com Erro (Se não encontrar)
    if ($null -eq $isoDrive) {
        Add-Log "ERRO: Unidade da ISO não encontrada!" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show(
            "Nenhuma unidade de ISO montada com 'setup.exe' foi encontrada.`n`nPor favor, monte a imagem ISO do Windows 11 e tente novamente.",
            "Erro: ISO Não Encontrada",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # 4. Preparar e Executar o Comando (Setupprep.exe)
    $sourcesDir = Join-Path -Path $isoDrive -ChildPath "sources"
    $setupPrepExe = Join-Path -Path $sourcesDir -ChildPath "Setupprep.exe"

    if (-not (Test-Path $setupPrepExe)) {
        Add-Log "ERRO: 'Setupprep.exe' não encontrado em $sourcesDir" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show(
            "Erro: 'Setupprep.exe' não foi encontrado em '$sourcesDir'.`n`nA ISO pode estar corrompida ou não é uma ISO de instalação do Windows.",
            "Erro: Arquivo Faltando",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    
    # 5. Executar o comando do .bat
    try {
        Add-Log "Iniciando o processo de instalação: $setupPrepExe /product server"
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        Start-Process -FilePath $setupPrepExe -ArgumentList "/product server" -WorkingDirectory $sourcesDir -ErrorAction Stop
        
        [System.Windows.Forms.MessageBox]::Show(
            "O instalador do Windows 11 (Bypass) foi iniciado!`n`nSiga as instruções na nova janela para concluir a instalação.",
            "Instalação Iniciada",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

    } catch {
        Add-Log "ERRO ao executar o Setupprep.exe: $($_.Exception.Message)" -Type "Error"
        [System.Windows.Forms.MessageBox]::Show(
            ("Ocorreu um erro ao tentar iniciar o instalador:`n" + $_.Exception.Message),
            "Erro na Execução",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}


# #############################################################################
# # (EASTER EGG - Icone Matrix) ACAO: Alternar Tema (V90)
# #############################################################################
# A lógica agora está no evento Add_Click do $matrixIcon


# #############################################################################
# # (EASTER EGG - Titulo Shift+Click) ACAO: Mostrar Uptime (V90)
# #############################################################################
function Show-Uptime {
    try {
        Add-Log "Calculando Uptime do sistema..."
        [System.Windows.Forms.Application]::DoEvents() # V110
        
        $osInfo = Get-CimInstance Win32_OperatingSystem
        $lastBootUpTime = $osInfo.LastBootUpTime
        $uptime = New-TimeSpan -Start $lastBootUpTime -End (Get-Date)

        # Formatar a string
        $uptimeString = "{0:00} Dias, {1:00} Horas, {2:00} Minutos, {3:00} Segundos" -f `
                            $uptime.Days, `
                            $uptime.Hours, `
                            $uptime.Minutes, `
                            $uptime.Seconds
        
        Add-Log "Uptime: $uptimeString" -Type "Info"
        [System.Windows.Forms.MessageBox]::Show(
            ("Tempo de atividade do sistema:`n" + $uptimeString),
            "Uptime",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } catch {
        Add-Log "Erro ao obter uptime: $($_.Exception.Message)" -Type "Warning"
        [System.Windows.Forms.MessageBox]::Show(
            "Não foi possível obter o tempo de atividade do sistema.",
            "Erro Uptime",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
}

# #############################################################################
# # (EASTER EGG - Título Duplo Clique) ACAO: Tremer Janela (V93)
# #############################################################################
function Shake-Window {
    param( [System.Windows.Forms.Form]$targetForm )

    if ($targetForm -eq $null) { Add-Log "Shake-Window: targetForm is null" -Type "Warning"; return }
    Add-Log "Easter Egg: Tremer Janela!" -Type "Info"
    try {
        $originalLocation = $targetForm.Location
        $originalX = $originalLocation.X
        $originalY = $originalLocation.Y

        $shakeAmount = 7
        $shakeDurationMs = 35
        $shakeCount = 5

        [System.Windows.Forms.Application]::DoEvents()

        for ($i = 0; $i -lt $shakeCount; $i++) {
            $targetForm.Invoke([Action]{
                $targetForm.Location = New-Object System.Drawing.Point($originalX - $shakeAmount, $originalY)
                $targetForm.Refresh()
            })
            Start-Sleep -Milliseconds $shakeDurationMs

            $targetForm.Invoke([Action]{
                $targetForm.Location = New-Object System.Drawing.Point($originalX + $shakeAmount, $originalY)
                $targetForm.Refresh()
            })
            Start-Sleep -Milliseconds $shakeDurationMs
        }

        $targetForm.Invoke([Action]{
            $targetForm.Location = $originalLocation
            $targetForm.Refresh()
        })

    } catch {
        Add-Log "Erro ao tremer janela: $($_.Exception.Message)" -Type "Warning"
        try {
            if ($originalLocation -ne $null) {
                $targetForm.Invoke([Action]{
                    $targetForm.Location = $originalLocation
                    $targetForm.Refresh()
                })
            }
        } catch {}
    }
}


# --- ### FUNCAO ANIMACAO DE FECHO (V103: Som de fecho removido) ### ---
function Play-ClosingAnimationAndClose {
    # V103: Parametro $playDefaultSound removido

    if ($script:isClosing -eq $true) { return }
    $script:isClosing = $true
    Add-Log "Fechando a aplicação..." -Type "Info"

    # V103: Parar a musica de fundo IMEDIATAMENTE
    if ($backgroundMusicPlayer -ne $null) {
        try { $backgroundMusicPlayer.Stop() } catch {}
    }

    [System.Collections.Generic.List[System.Windows.Forms.Control]]$controlsToHide = [System.Collections.Generic.List[System.Windows.Forms.Control]]::new($ControlsToAnimate)
    $controlsToHide.Reverse()

    if ($labelCredit.Visible) {
        $labelCredit.Visible = $false; $form.Refresh(); Start-Sleep -Milliseconds 60
    }
    
    # V109: Esconder o LogBox primeiro
    if ($script:logBox -ne $null -and $script:logBox.Visible) {
        $script:logBox.Visible = $false; $form.Refresh(); Start-Sleep -Milliseconds 40
    }
    if ($labelLog -ne $null -and $labelLog.Visible) {
        $labelLog.Visible = $false; $form.Refresh(); Start-Sleep -Milliseconds 40
    }

    foreach ($control in $controlsToHide) {
        if ($control.Visible) {
            $control.Visible = $false
            $form.Refresh()
            Start-Sleep -Milliseconds 40
        }
    }

    Start-Sleep -Milliseconds 150

    # V103: Bloco 'if ($playDefaultSound)' e $closingSoundPlayer.PlaySync() removidos
    
    $form.Dispose()
}


# #############################################################################
# # CONSTRUCAO DA JANELA GRAFICA (GUI)
# #############################################################################

# --- Definicao das Cores do Tema Escuro ---
$darkBackColor = [System.Drawing.Color]::FromArgb(255, 32, 32, 32)
$darkForeColor = [System.Drawing.Color]::FromArgb(255, 240, 240, 240)
$darkButtonBackColor = [System.Drawing.Color]::FromArgb(255, 60, 60, 60)
$darkButtonForeColor = $darkForeColor
$darkButtonBorderColor = [System.Drawing.Color]::FromArgb(255, 100, 100, 100)
$darkButtonMouseOverColor = [System.Drawing.Color]::FromArgb(255, 80, 80, 80)
$darkButtonMouseDownColor = [System.Drawing.Color]::FromArgb(255, 100, 100, 100)
# V109: Cores do Log Escuro
$darkLogBackColor = [System.Drawing.Color]::FromArgb(255, 25, 25, 25)
$darkLogForeColor = [System.Drawing.Color]::FromArgb(255, 220, 220, 220)

# --- Definicao das Cores do Tema Claro (V90) ---
$lightBackColor = [System.Drawing.Color]::FromArgb(255, 240, 240, 240)
$lightForeColor = [System.Drawing.Color]::FromArgb(255, 20, 20, 20)
$lightButtonBackColor = [System.Drawing.Color]::FromArgb(255, 225, 225, 225)
$lightButtonForeColor = $lightForeColor
$lightButtonBorderColor = [System.Drawing.Color]::FromArgb(255, 173, 173, 173)
$lightButtonMouseOverColor = [System.Drawing.Color]::FromArgb(255, 200, 200, 200)
$lightButtonMouseDownColor = [System.Drawing.Color]::FromArgb(255, 180, 180, 180)
# V109: Cores do Log Claro
$lightLogBackColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 255)
$lightLogForeColor = [System.Drawing.Color]::FromArgb(255, 10, 10, 10)

# --- Configuracao Base do Formulario ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Ferramentas IDEIA INFORMATICA"
$form.Size = New-Object System.Drawing.Size(420, 730) # V110: Altura aumentada
$form.StartPosition = "CenterScreen"
$form.TopMost = $false
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.BackColor = $darkBackColor # Começa escuro
$form.ForeColor = $darkForeColor # Começa escuro
$form.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::None

# --- Fonte Padrao ---
$defaultFontFamily = "Segoe UI"; if (-not ([System.Drawing.FontFamily]::Families | Where-Object {$_.Name -eq $defaultFontFamily})) { $defaultFontFamily = "Arial" }
$defaultFont = New-Object System.Drawing.Font($defaultFontFamily, 9)
$buttonFont = New-Object System.Drawing.Font($defaultFontFamily, 10, [System.Drawing.FontStyle]::Bold)
$titleFont = New-Object System.Drawing.Font($defaultFontFamily, 10, [System.Drawing.FontStyle]::Bold)
$creditFont = New-Object System.Drawing.Font($defaultFontFamily, 8, [System.Drawing.FontStyle]::Italic)
$logFont = New-Object System.Drawing.Font("Consolas", 8) # V109: Fonte para o Log

# --- ToolTip Object ---
$ToolTip = New-Object System.Windows.Forms.ToolTip

# --- Lista para guardar controlos para animacao/esconder ---
$ControlsToAnimate = [System.Collections.Generic.List[System.Windows.Forms.Control]]::new()

# --- Titulo (com Easter Eggs Uptime e Shake V91) ---
$label = New-Object System.Windows.Forms.Label; $label.Location = New-Object System.Drawing.Point(20, 20); $label.Size = New-Object System.Drawing.Size(300, 30); # V107: Largura ajustada
$label.Text = "Selecione a ferramenta que deseja executar:"; $label.Font = $titleFont; $label.ForeColor = $darkForeColor; $label.BackColor = $darkBackColor
$label.Visible = $false
$label.Cursor = [System.Windows.Forms.Cursors]::Hand # Indica clicável
$label.Add_Click({
    if ([System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Shift) {
        Show-Uptime # Chama a função Uptime
    }
})
$label.Add_DoubleClick({ # V91: Adiciona DoubleClick para Shake
    Shake-Window -targetForm $form # Passa o formulário
})
$ToolTip.SetToolTip($label, "Shift+Click para ver Uptime | Duplo Clique para ...?") # Tooltip (V91)
$form.Controls.Add($label)
$ControlsToAnimate.Add($label)

# --- ### V103: ICONE DE MUDO (CORRIGIDO) ### ---
$muteIcon = $null
if ($unmuteImage -ne $null -and $muteImage -ne $null) {
    $muteIcon = New-Object System.Windows.Forms.PictureBox
    $iconSize = 30 # Tamanho padrão
    # A POSICAO SERA DEFINIDA NO EVENTO SHOWN
    $muteIcon.Size = New-Object System.Drawing.Size($iconSize, $iconSize)
    $muteIcon.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $muteIcon.BackColor = $darkBackColor # Cor inicial
    $muteIcon.Cursor = [System.Windows.Forms.Cursors]::Hand
    $muteIcon.Visible = $false
    $muteIcon.Image = $unmuteImage # Começa com som
    $muteIcon.Add_Click({
        $script:isMuted = -not $script:isMuted # Inverte o estado
        
        if ($backgroundMusicPlayer -eq $null) { return } # Não faz nada se a música não carregou
        
        if ($script:isMuted) {
            # Mutar
            Add-Log "Música mutada." -Type "Info"
            try { $backgroundMusicPlayer.Stop() } catch {}
            $muteIcon.Image = $muteImage
            $ToolTip.SetToolTip($muteIcon, "Ativar música")
        } else {
            # Desmutar
            Add-Log "Música ativada." -Type "Info"
            try { $backgroundMusicPlayer.PlayLooping() } catch {}
            $muteIcon.Image = $unmuteImage
            $ToolTip.SetToolTip($muteIcon, "Mutar música")
        }
    })
    
    $ToolTip.SetToolTip($muteIcon, "Mutar música")
    $form.Controls.Add($muteIcon)
    $ControlsToAnimate.Add($muteIcon)

} else { Add-Log "AVISO: Ícones de Mudo não carregados. Botão não será criado." -Type "Warning" }


# --- ### ICONE EASTER EGG (Alternar Tema V97) ### ---
$matrixIconPath = "$PSScriptRoot\matrix.png" # <-- FORNECER .PNG
$matrixIcon = $null
if (Test-Path $matrixIconPath) {
    $matrixIcon = New-Object System.Windows.Forms.PictureBox
    $iconSize = 30
    # A POSICAO SERA DEFINIDA NO EVENTO SHOWN
    $matrixIcon.Size = New-Object System.Drawing.Size($iconSize, $iconSize)
    $matrixIcon.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $matrixIcon.BackColor = $darkBackColor # Cor inicial
    $matrixIcon.Cursor = [System.Windows.Forms.Cursors]::Hand
    $matrixIcon.Visible = $false
    try {
        $matrixIcon.Image = [System.Drawing.Image]::FromFile($matrixIconPath)
        
        # --- V106: LÓGICA DE CLIQUE COM PlaySync() ---
        $matrixIcon.Add_Click({
            if ($themeSwitchSoundPlayer) {
                try { $themeSwitchSoundPlayer.PlaySync() } catch { Play-ClickSound } # Usa PlaySync()
            } else { Play-ClickSound }
            
            if ($script:currentTheme -eq 'dark') {
                Add-Log "Trocando para Tema Claro." -Type "Info"
                Set-Theme -ThemeName 'light'
            } else {
                Add-Log "Trocando para Tema Escuro." -Type "Info"
                Set-Theme -ThemeName 'dark'
            }
            
            Resume-BackgroundMusic # <-- V105: Reinicia a música APÓS PlaySync() terminar
        })
        # --- FIM V106 ---
        
        $ToolTip.SetToolTip($matrixIcon, "Alternar Tema (Escuro/Claro)")
        $form.Controls.Add($matrixIcon)
        $ControlsToAnimate.Add($matrixIcon)
    } catch {
        Add-Log "ERRO ao carregar 'matrix.png': $($_.Exception.Message)" -Type "Error"
        $matrixIcon = $null
    }
} else { Add-Log "Imagem 'matrix.png' nao encontrada para o Icone de Tema." -Type "Warning" }


# --- Posicao Y inicial e Espacamento para os BOTOES ---
$ButtonY = 70; $ButtonHeight = 40; $ButtonSpacing = 12
$ButtonWidth = 170
$Column1X = 20
$Column2X = 210

# --- Funcao Auxiliar para Estilizar Botoes COM TEMA (V90) ---
function Style-ButtonWithTheme {
    param(
        [System.Windows.Forms.Button]$Button,
        [string]$Text,
        [string]$ToolTipText,
        [System.Drawing.Font]$Font,
        [System.Drawing.Color]$BackColor,
        [System.Drawing.Color]$ForeColor,
        [System.Drawing.Color]$BorderColor,
        [System.Drawing.Color]$MouseOverColor,
        [System.Drawing.Color]$MouseDownColor
    )
    $Button.Text = $Text
    $Button.Font = $Font
    $Button.BackColor = $BackColor
    $Button.ForeColor = $ForeColor
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = $BorderColor
    $Button.FlatAppearance.MouseOverBackColor = $MouseOverColor
    $Button.FlatAppearance.MouseDownBackColor = $MouseDownColor
    $ToolTip.SetToolTip($Button, $ToolTipText)
}

# --- ### FUNCAO PARA MUDAR O TEMA (V109: Atualizado) ### ---
function Set-Theme {
    param( [string]$ThemeName )

    if ($ThemeName -eq 'light') {
        $bgColor = $lightBackColor; $fgColor = $lightForeColor; $btnBg = $lightButtonBackColor; $btnFg = $lightButtonForeColor
        $btnBorder = $lightButtonBorderColor; $btnHover = $lightButtonMouseOverColor; $btnDown = $lightButtonMouseDownColor
        $logBg = $lightLogBackColor; $logFg = $lightLogForeColor
    } else {
        $ThemeName = 'dark'; $bgColor = $darkBackColor; $fgColor = $darkForeColor; $btnBg = $darkButtonBackColor; $btnFg = $darkButtonForeColor
        $btnBorder = $darkButtonBorderColor; $btnHover = $darkButtonMouseOverColor; $btnDown = $darkButtonMouseDownColor
        $logBg = $darkLogBackColor; $logFg = $darkLogForeColor
    }

    $form.BackColor = $bgColor; $form.ForeColor = $fgColor
    $label.BackColor = $bgColor; $label.ForeColor = $fgColor
    $labelCredit.BackColor = $bgColor; $labelCredit.ForeColor = $fgColor
    
    # V109: Atualiza log e logbox
    if ($labelLog -ne $null) { $labelLog.BackColor = $bgColor; $labelLog.ForeColor = $fgColor }
    if ($script:logBox -ne $null) { $script:logBox.BackColor = $logBg; $script:logBox.ForeColor = $logFg }
    
    if ($matrixIcon -ne $null) { $matrixIcon.BackColor = $bgColor }
    if ($muteIcon -ne $null) { $muteIcon.BackColor = $bgColor } # V103: Adicionado
    if ($robotImage -ne $null) { $robotImage.BackColor = $bgColor }

    $allButtons = @(
        $buttonRename, $buttonDesktopIcons, $buttonPower, $buttonFastStartup,
        $buttonDriverDiag, $buttonPrinter, $buttonSpooler, $buttonNinite,
        $buttonDownloadOffice, $buttonOfficeShortcuts, $buttonSmb, $buttonNew,
        $buttonDisableBitLocker, $buttonInstallWin11,
        $buttonExit, $buttonExitRestart
    )

    foreach ($btn in $allButtons) {
        $currentFont = if ($btn -eq $buttonExit -or $btn -eq $buttonExitRestart) { $defaultFont } else { $buttonFont }
        Style-ButtonWithTheme $btn $btn.Text ($ToolTip.GetToolTip($btn)) $currentFont $btnBg $btnFg $btnBorder $btnHover $btnDown
    }
    $script:currentTheme = $ThemeName
    $form.Refresh()
}

# ######################################
# # CRIAÇÃO DOS BOTÕES (1 a 14)
# ######################################

# ######################################
# # GRUPO 1: CONFIGURACAO INICIAL ... (Botoes 1 a 4)
# ######################################
# --- Botao 1: Mudar Nomes (V108) ---
$buttonRename = New-Object System.Windows.Forms.Button
$buttonRename.Location = New-Object System.Drawing.Point($Column1X, $ButtonY)
$buttonRename.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonRename "1. Mudar Nomes" "Muda nome do PC e Nome Completo do usuário atual." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor # V107: Texto e Tooltip atualizados
$buttonRename.Add_MouseEnter({ Play-HoverSound }); $buttonRename.Add_Click({ Play-ClickSound; Start-RenameComputer })
$buttonRename.Visible = $false; $form.Controls.Add($buttonRename); $ControlsToAnimate.Add($buttonRename)

# --- Botao 2: Mostrar Icones Desktop ---
$buttonDesktopIcons = New-Object System.Windows.Forms.Button
$buttonDesktopIcons.Location = New-Object System.Drawing.Point($Column2X, $ButtonY)
$buttonDesktopIcons.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonDesktopIcons "2. Mostrar Icones" "Mostra os icones 'Computador', 'Ficheiros do Utilizador' e 'Rede' na Area de Trabalho." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonDesktopIcons.Add_MouseEnter({ Play-HoverSound }); $buttonDesktopIcons.Add_Click({ Play-ClickSound; Start-ShowDesktopIcons })
$buttonDesktopIcons.Visible = $false; $form.Controls.Add($buttonDesktopIcons); $ControlsToAnimate.Add($buttonDesktopIcons)

$ButtonY += $ButtonHeight + $ButtonSpacing # Incrementa a linha

# --- Botao 3: Definir Energia (Combinado) ---
$buttonPower = New-Object System.Windows.Forms.Button
$buttonPower.Location = New-Object System.Drawing.Point($Column1X, $ButtonY)
$buttonPower.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonPower "3. Definir Energia" "Define o plano de energia para 'Alto Desempenho' e configura para nunca desligar o monitor ou suspender (na Tomada e Bateria)." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonPower.Add_MouseEnter({ Play-HoverSound }); $buttonPower.Add_Click({ Play-ClickSound; Start-SetPowerNever })
$buttonPower.Visible = $false; $form.Controls.Add($buttonPower); $ControlsToAnimate.Add($buttonPower)

# --- Botao 4: Desativar Inicializacao Rapida ---
$buttonFastStartup = New-Object System.Windows.Forms.Button
$buttonFastStartup.Location = New-Object System.Drawing.Point($Column2X, $ButtonY)
$buttonFastStartup.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonFastStartup "4. Desat. Init. Rapida" "Desativa a funcionalidade de Inicializacao Rapida do Windows. Requer reinicio." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonFastStartup.Add_MouseEnter({ Play-HoverSound }); $buttonFastStartup.Add_Click({ Play-ClickSound; Start-DisableFastStartup })
$buttonFastStartup.Visible = $false; $form.Controls.Add($buttonFastStartup); $ControlsToAnimate.Add($buttonFastStartup)

$ButtonY += $ButtonHeight + $ButtonSpacing # Incrementa a linha

# ######################################
# # GRUPO 2: HARDWARE E DIAGNOSTICO ... (Botoes 5 a 7)
# ######################################
# --- Botao 5: Diagnostico de Drivers ---
$buttonDriverDiag = New-Object System.Windows.Forms.Button
$buttonDriverDiag.Location = New-Object System.Drawing.Point($Column1X, $ButtonY)
$buttonDriverDiag.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonDriverDiag "5. Diagnostico Drivers" "Abre Gestor de Dispositivos, DxDiag (Video/Som) e Camera para verificacao." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonDriverDiag.Add_MouseEnter({ Play-HoverSound }); $buttonDriverDiag.Add_Click({ Play-ClickSound; Start-DriverDiagnostics })
$buttonDriverDiag.Visible = $false; $form.Controls.Add($buttonDriverDiag); $ControlsToAnimate.Add($buttonDriverDiag)

# --- Botao 6: Configurar Impressora ---
$buttonPrinter = New-Object System.Windows.Forms.Button
$buttonPrinter.Location = New-Object System.Drawing.Point($Column2X, $ButtonY)
$buttonPrinter.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonPrinter "6. Configurar Impressora" "Inicia o assistente para criar um utilizador 'Impressora', aplicar correcoes e partilhar uma impressora." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonPrinter.Add_MouseEnter({ Play-HoverSound }); $buttonPrinter.Add_Click({ Play-ClickSound; Start-PrinterScript }) # V110: .Hide()/.Show() removidos
$buttonPrinter.Visible = $false; $form.Controls.Add($buttonPrinter); $ControlsToAnimate.Add($buttonPrinter)

$ButtonY += $ButtonHeight + $ButtonSpacing # Incrementa a linha

# --- Botao 7: Reiniciar Spooler ---
$buttonSpooler = New-Object System.Windows.Forms.Button
$buttonSpooler.Location = New-Object System.Drawing.Point($Column1X, $ButtonY)
$buttonSpooler.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonSpooler "7. Reiniciar Spooler" "Para e reinicia o servico Spooler de Impressao. Util para resolver problemas de impressao." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonSpooler.Add_MouseEnter({ Play-HoverSound }); $buttonSpooler.Add_Click({ Play-ClickSound; Start-RestartSpooler })
$buttonSpooler.Visible = $false; $form.Controls.Add($buttonSpooler); $ControlsToAnimate.Add($buttonSpooler)


# ######################################
# # GRUPO 3: INSTALACAO DE SOFTWARE ... (Botoes 8 a 10)
# ######################################

# --- Botao 8: Abrir Ninite.com ---
$buttonNinite = New-Object System.Windows.Forms.Button
$buttonNinite.Location = New-Object System.Drawing.Point($Column2X, $ButtonY)
$buttonNinite.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonNinite "8. Abrir Ninite.com" "Abre Ninite.com no navegador para instalar varios aplicativos populares de uma vez." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonNinite.Add_MouseEnter({ Play-HoverSound }); $buttonNinite.Add_Click({ Play-ClickSound; Start-OpenNinite })
$buttonNinite.Visible = $false; $form.Controls.Add($buttonNinite); $ControlsToAnimate.Add($buttonNinite)

$ButtonY += $ButtonHeight + $ButtonSpacing # Incrementa a linha

# --- Botao 9: Abrir Link de Download do Office 365 ---
$buttonDownloadOffice = New-Object System.Windows.Forms.Button
$buttonDownloadOffice.Location = New-Object System.Drawing.Point($Column1X, $ButtonY)
$buttonDownloadOffice.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonDownloadOffice "9. Download Office" "Abre o link direto para download do instalador online do Office 365 ProPlus x64 PT-BR." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonDownloadOffice.Add_MouseEnter({ Play-HoverSound }); $buttonDownloadOffice.Add_Click({ Play-ClickSound; Start-OpenOfficeDownloadLink })
$buttonDownloadOffice.Visible = $false; $form.Controls.Add($buttonDownloadOffice); $ControlsToAnimate.Add($buttonDownloadOffice)

# --- Botao 10: Criar Atalhos Office ---
$buttonOfficeShortcuts = New-Object System.Windows.Forms.Button
$buttonOfficeShortcuts.Location = New-Object System.Drawing.Point($Column2X, $ButtonY)
$buttonOfficeShortcuts.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonOfficeShortcuts "10. Criar Atalhos Office" "Tenta encontrar Word, Excel e PowerPoint e cria atalhos na Area de Trabalho." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonOfficeShortcuts.Add_MouseEnter({ Play-HoverSound }); $buttonOfficeShortcuts.Add_Click({ Play-ClickSound; Start-CreateOfficeShortcuts })
$buttonOfficeShortcuts.Visible = $false; $form.Controls.Add($buttonOfficeShortcuts); $ControlsToAnimate.Add($buttonOfficeShortcuts)

$ButtonY += $ButtonHeight + $ButtonSpacing # Incrementa a linha

# ######################################
# # GRUPO 4: AVANCADO / LABORATORIO ... (Botoes 11 a 14)
# ######################################

# --- Botao 11: Corrigir Compartilhamento de Rede (SMB) ---
$buttonSmb = New-Object System.Windows.Forms.Button
$buttonSmb.Location = New-Object System.Drawing.Point($Column1X, $ButtonY)
$buttonSmb.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonSmb "11. Corrigir Rede (SMB)" "AVISO: Reduz a seguranca. Permite acesso a partilhas de rede antigas. Requer reinicio." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonSmb.Add_MouseEnter({ Play-HoverSound }); $buttonSmb.Add_Click({ Play-ClickSound; Start-SmbScript }) # V110: .Hide()/.Show() removidos
$buttonSmb.Visible = $false; $form.Controls.Add($buttonSmb); $ControlsToAnimate.Add($buttonSmb)

# --- Botao 12: Ativar Windows e Office ---
$buttonNew = New-Object System.Windows.Forms.Button
$buttonNew.Location = New-Object System.Drawing.Point($Column2X, $ButtonY)
$buttonNew.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonNew "12. Ativar Windows/Office" "Executa o script de ativacao personalizado (Apenas para Laboratorio)." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonNew.Add_MouseEnter({ Play-HoverSound }); $buttonNew.Add_Click({ Play-ClickSound; Start-NewButtonAction }) # V110: .Hide()/.Show() removidos
$buttonNew.Visible = $false; $form.Controls.Add($buttonNew); $ControlsToAnimate.Add($buttonNew)

$ButtonY += $ButtonHeight + $ButtonSpacing # Incrementa a linha

# --- Botao 13: Desativar BitLocker (Lab) ---
$buttonDisableBitLocker = New-Object System.Windows.Forms.Button
$buttonDisableBitLocker.Location = New-Object System.Drawing.Point($Column1X, $ButtonY) # Coluna 1
$buttonDisableBitLocker.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonDisableBitLocker "13. Desativar BitLocker" "PERIGO: Inicia a descriptografia permanente da unidade C:. Use apenas em laboratorio." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonDisableBitLocker.Add_MouseEnter({ Play-HoverSound }); $buttonDisableBitLocker.Add_Click({ Play-ClickSound; Start-DisableBitLocker })
$buttonDisableBitLocker.Visible = $false; $form.Controls.Add($buttonDisableBitLocker); $ControlsToAnimate.Add($buttonDisableBitLocker)

# --- Botao 14: Instalar Win 11 (Bypass) - (NOVO V102) ---
$buttonInstallWin11 = New-Object System.Windows.Forms.Button
$buttonInstallWin11.Location = New-Object System.Drawing.Point($Column2X, $ButtonY) # Coluna 2
$buttonInstallWin11.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
Style-ButtonWithTheme $buttonInstallWin11 "14. Instalar Win 11" "Inicia a instalação do Win 11 a partir de uma ISO montada, ignorando os requisitos (TPM/CPU)." $buttonFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonInstallWin11.Add_MouseEnter({ Play-HoverSound }); $buttonInstallWin11.Add_Click({ Play-ClickSound; Start-InstallWin11 }) # V110: .Hide()/.Show() removidos
$buttonInstallWin11.Visible = $false; $form.Controls.Add($buttonInstallWin11); $ControlsToAnimate.Add($buttonInstallWin11)

# --- ### V109: NOVA ÁREA DE LOG ### ---

$ButtonY += $ButtonHeight + $ButtonSpacing # Incrementa Y para o Log

# --- Titulo do Log ---
$labelLog = New-Object System.Windows.Forms.Label
$labelLog.Location = New-Object System.Drawing.Point($Column1X, $ButtonY)
$labelLog.Size = New-Object System.Drawing.Size(200, 20)
$labelLog.Text = "Log de Atividades:"
$labelLog.Font = $defaultFont
$labelLog.ForeColor = $darkForeColor
$labelLog.BackColor = $darkBackColor
$labelLog.Visible = $false
$form.Controls.Add($labelLog)

$ButtonY += 20 # Espaço entre título do log e caixa

# --- Caixa de Log (RichTextBox) ---
$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = New-Object System.Drawing.Point($Column1X, $ButtonY)
$logBox.Size = New-Object System.Drawing.Size(370, 180) # (420 - 20 - 30) = 370
$logBox.Font = $logFont
$logBox.BackColor = $darkLogBackColor # V109: Cor de fundo específica
$logBox.ForeColor = $darkLogForeColor # V109: Cor de texto específica
$logBox.ReadOnly = $true
$logBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$logBox.Visible = $false
$form.Controls.Add($logBox)
$script:logBox = $logBox # Armazena na variável global

$ButtonY += 180 + 25 # V110: Espaço aumentado (era 15)

# --- ### FIM DA ÁREA DE LOG ### ---


# ######################################
# # CONTROLES DE SAIDA
# ######################################

# --- Botao Sair (Não numerado) ---
$buttonExit = New-Object System.Windows.Forms.Button; $buttonExit.Location = New-Object System.Drawing.Point(80, $ButtonY); $buttonExit.Size = New-Object System.Drawing.Size(100, 30);
Style-ButtonWithTheme $buttonExit "Sair" "Fecha a ferramenta. (Shift+Click para Easter Egg)" $defaultFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonExit.Add_MouseEnter({ Play-HoverSound });
$buttonExit.Add_Click({
    if ([System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Shift) {
        # Easter Egg Death Note
        Add-Log "Easter Egg: Death Note ativado." -Type "Warning"
        if ($deathNoteBGImage -ne $null) {
            foreach ($control in $ControlsToAnimate) { if($control.Visible){$control.Visible = $false} }
            if($labelCredit.Visible){$labelCredit.Visible = $false}
            if ($script:logBox.Visible) { $script:logBox.Visible = $false } # V109
            if ($labelLog.Visible) { $labelLog.Visible = $false } # V109
            
            $form.Refresh()
            $form.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Zoom
            $form.BackgroundImage = $deathNoteBGImage
            $form.Refresh()
            if ($deathNoteSoundPlayer) { try { $deathNoteSoundPlayer.PlaySync() } catch { Add-Log "Erro ao tocar Death_note.wav" -Type "Error"; Start-Sleep -Seconds 3 } } else { Play-ClickSound; Start-Sleep -Seconds 3 }
            $form.BackgroundImage = $null
        } else {
             if ($deathNoteSoundPlayer) { try { $deathNoteSoundPlayer.PlaySync() } catch { Play-ClickSound } } else { Play-ClickSound }
        }
        Play-ClosingAnimationAndClose # V104: CORRIGIDO (removido '()')
    } else {
        Play-ClickSound
        Play-ClosingAnimationAndClose # V104: CORRIGIDO (removido '()')
    }
})
$buttonExit.Visible = $false; $form.Controls.Add($buttonExit); $ControlsToAnimate.Add($buttonExit);

# --- Botao Reiniciar (Não numerado) ---
$buttonExitRestart = New-Object System.Windows.Forms.Button; $buttonExitRestart.Location = New-Object System.Drawing.Point(200, $ButtonY); $buttonExitRestart.Size = New-Object System.Drawing.Size(100, 30);
Style-ButtonWithTheme $buttonExitRestart "Reiniciar" "Reinicia o computador." $defaultFont $darkButtonBackColor $darkButtonForeColor $darkButtonBorderColor $darkButtonMouseOverColor $darkButtonMouseDownColor
$buttonExitRestart.Add_MouseEnter({ Play-HoverSound });
$buttonExitRestart.Add_Click({
    Play-ClickSound
    $confirmRestart = [System.Windows.Forms.MessageBox]::Show("Tem a certeza que deseja reiniciar o computador agora?", "Confirmar Reinicio", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirmRestart -eq 'Yes') {
        # V103: Parar a musica antes de reiniciar
        if ($backgroundMusicPlayer -ne $null) { try { $backgroundMusicPlayer.Stop() } catch {} }
        Add-Log "REINICIANDO O COMPUTADOR..." -Type "Error"
        Start-Sleep -Seconds 1
        Restart-Computer -Force
    }
})
$buttonExitRestart.Visible = $false; $form.Controls.Add($buttonExitRestart); $ControlsToAnimate.Add($buttonExitRestart)

# --- Imagem do Robo ---
$robotPath = "$PSScriptRoot\robo.png"
$robotImage = $null
if (Test-Path $robotPath) {
    $robotImage = New-Object System.Windows.Forms.PictureBox
    $robotImage.Location = New-Object System.Drawing.Point(310, $buttonExit.Location.Y)
    $robotImage.Size = New-Object System.Drawing.Size(30, 30)
    $robotImage.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $robotImage.BackColor = $darkBackColor
    $robotImage.Visible = $false
    $robotImage.Image = [System.Drawing.Image]::FromFile($robotPath)
    
    # --- V106: LÓGICA DE CLIQUE COM PlaySync() ---
    $robotImage.Add_Click({
        if ($robotSoundPlayer) {
            try { $robotSoundPlayer.PlaySync() } catch { Play-ClickSound } # Usa PlaySync()
        } else { Play-ClickSound }
        
        Resume-BackgroundMusic # <-- V105: Reinicia a música APÓS PlaySync() terminar
    })
    # --- FIM V106 ---
    
    $ToolTip.SetToolTip($robotImage, "Obrigado por usar a ferramenta!")
    $form.Controls.Add($robotImage)
    $ControlsToAnimate.Add($robotImage)
} else { Add-Log "Imagem 'robo.png' nao encontrada." -Type "Warning" }
$ButtonY += 30 + 10 # Incrementa Y para os Creditos

# --- Creditos (com Easter Egg V101: Som + Shake) ---
$labelCredit = New-Object System.Windows.Forms.Label; $labelCredit.Location = New-Object System.Drawing.Point(10, $ButtonY); $labelCredit.Size = New-Object System.Drawing.Size(380, 20); $labelCredit.Text = "Criado por: IDEIA INFORMATICA"; $labelCredit.TextAlign = "MiddleCenter"; $labelCredit.Font = $creditFont; $labelCredit.ForeColor = $darkForeColor; $labelCredit.BackColor = $darkBackColor;
$labelCredit.Cursor = [System.Windows.Forms.Cursors]::Hand # Cursor de mão
$labelCredit.Visible = $false;

# --- V106: LÓGICA DE CLIQUE COM PlaySync() ---
$labelCredit.Add_Click({
    $script:creditClicks++
    if ($script:creditClicks -ge 5) {
        Add-Log "Easter Egg: Fuleiro ativado!" -Type "Info"
        if ($fuleiroSoundPlayer) {
            try { $fuleiroSoundPlayer.PlaySync() } catch { Play-ClickSound } # Usa PlaySync()
        } else { Play-ClickSound }
        
        Shake-Window -targetForm $form
        
        # O MessageBox.Show() pausa o script
        [System.Windows.Forms.MessageBox]::Show("AHHHHHHH FULEIRO!!!","Easter Egg!",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        
        $script:creditClicks = 0 # Reseta
        
        Resume-BackgroundMusic # <-- V105: Reinicia a música APÓS MessageBox fechar
    }
})
# --- FIM V106 ---

$form.Controls.Add($labelCredit) # Adiciona ANTES do evento Shown

# --- ### EVENTO 'Shown' PARA ANIMACAO DE ABERTURA ### ---
$form.Add_Shown({
    # V98: Easter Egg de Abertura Rara
    if ($script:playSpecialOpening -eq $true) {
        if ($specialOpeningBGImage -ne $null -and $specialOpeningSoundPlayer -ne $null) {
            $form.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Zoom
            $form.BackgroundImage = $specialOpeningBGImage
            $form.Refresh()
            try { $specialOpeningSoundPlayer.PlaySync() } catch { Add-Log "Erro ao tocar special_opening_sound.wav" -Type "Error"; Start-Sleep -Seconds 3 } # .PlaySync() é intencional
            $form.BackgroundImage = $null
            $form.Refresh()
        }
    }
    # Fim do Easter Egg V98

    $form.Refresh()
    Start-Sleep -Milliseconds 100

    # V103: Tocar a musica de fundo (se não estiver mutado e o player existir)
    # V105: Abertura Rara (Acima) usa PlaySync(), então a música de fundo SÓ começará DEPOIS, o que é perfeito.
    #if ($backgroundMusicPlayer -ne $null -and $script:isMuted -eq $false) {
    #    try { $backgroundMusicPlayer.PlayLooping() } catch { Add-Log "Erro ao tocar background_music.wav" -Type "Error" }
    #}

    # V103: Define a posicao dos icones (Mudo e Tema)
    $iconSize = 30
    $margin = 15
    $iconSpacing = 5
    
    if ($matrixIcon -ne $null) {
        try {
            $matrixIconX = $form.ClientSize.Width - $iconSize - $margin
            $matrixIcon.Location = New-Object System.Drawing.Point($matrixIconX, $margin)
            
            # Posiciona o MuteIcon à esquerda do MatrixIcon
            if ($muteIcon -ne $null) {
                $muteIconX = $matrixIconX - $iconSize - $iconSpacing
                $muteIcon.Location = New-Object System.Drawing.Point($muteIconX, $margin)
            }
        } catch {
            Add-Log "ERRO ao calcular a posição dos ícones: $($_.Exception.Message)" -Type "Error"
        }
    }

    # Anima os Botoes (e Icones) em cascata
    foreach ($control in $ControlsToAnimate) {
        Start-Sleep -Milliseconds 50
        $control.Visible = $true
        $form.Refresh()
    }
    
    # V109: Anima o Log e os Creditos no final
    Start-Sleep -Milliseconds 50
    $labelLog.Visible = $true # Torna o titulo do log visível
    $script:logBox.Visible = $true # Torna o log visível
    $form.Refresh()
    
    Start-Sleep -Milliseconds 50
    $labelCredit.Visible = $true # Torna os créditos visíveis SÓ AGORA
    $form.Refresh()
    
    Add-Log "Interface carregada. Aguardando comandos..." -Type "Success"
})

# --- ### EVENTO 'FormClosing' PARA ANIMACAO E SOM DE FECHO ### ---
$form.Add_FormClosing({
    param($sender, $e)
    if ($script:isClosing -eq $true) {
        $e.Cancel = $false
        return
    }
    $e.Cancel = $true
    Play-ClosingAnimationAndClose # V104: CORRIGIDO (removido '()')
})

# --- Mostrar a janela ---
$form.ShowDialog()