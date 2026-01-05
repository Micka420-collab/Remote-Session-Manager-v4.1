<#
    RemoteSessionManager.ps1
    Utilitaire de gestion de sessions distantes.
    Version 4.0
    - Historique des commandes (↑/↓)
    - Boutons: Effacer Console, Exporter Log, Infos Système, Redémarrer, EventViewer
    - Timer de session
    - Raccourcis clavier (Ctrl+L, Ctrl+S, F5)
#>

# Forcer l'encodage de la console
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# Interface XAML
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Remote Session Manager (PRT)" Height="700" Width="950" WindowStartupLocation="CenterScreen" Background="#F3F3F3">
    
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="0,0,0,5"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" MinWidth="300"/>
            <ColumnDefinition Width="5"/>
            <ColumnDefinition Width="230" MinWidth="180"/>
        </Grid.ColumnDefinitions>

        <!-- BARRE DE CONNEXION -->
        <Border Name="bdrTop" Grid.Row="0" Grid.ColumnSpan="3" Margin="0,0,0,15" Padding="15" Background="White" CornerRadius="8">
            <Border.Effect>
                <DropShadowEffect Color="Gray" BlurRadius="5" ShadowDepth="1" Opacity="0.3"/>
            </Border.Effect>
            <DockPanel LastChildFill="False">
                <!-- Gauche : Saisie + Connexion -->
                <StackPanel DockPanel.Dock="Left" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Name="lblTitle" Text="CIBLE PRT" VerticalAlignment="Center" FontSize="16" FontWeight="Bold" Foreground="#333"/>
                    <TextBox Name="txtComputerNumber" Width="150" VerticalAlignment="Center" Margin="5,0,15,0" Padding="5" FontSize="14" MaxLength="30"/>
                    <CheckBox Name="chkNoPRT" Content="Sans PRT" VerticalAlignment="Center" Margin="0,0,15,0" ToolTip="Cocher pour ne pas ajouter le prefixe PRT automatiquement" FontSize="11"/>
                    
                    <Button Name="btnConnect" Content="Connecter (Session)" Background="#007ACC" Foreground="White" FontWeight="Bold" Margin="0,0,10,0"/>
                    <Button Name="btnDisconnect" Content="D&#233;connecter" Background="#D9534F" Foreground="White" FontWeight="Bold" Visibility="Collapsed"/>
                    
                    <StackPanel Orientation="Vertical" VerticalAlignment="Center" Margin="10,0,0,0">
                        <TextBlock Name="lblStatus" Text="D&#233;connect&#233;." FontSize="11" Foreground="Gray"/>
                        <TextBlock Name="lblCurrentTarget" Text="-- Aucune Cible --" FontWeight="Bold" Foreground="#D9534F"/>
                    </StackPanel>
                </StackPanel>

                <!-- Droite : Tools -->
                <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
                    <!-- Bouton Theme (Rectangle Arrondi) -->
                    <Button Name="btnTheme" Margin="0,0,10,0" ToolTip="Changer de th&#234;me" Background="#DDD">
                        <Button.Template>
                            <ControlTemplate TargetType="Button">
                                <Border CornerRadius="5" Padding="10,5" Background="{TemplateBinding Background}" BorderBrush="#999" BorderThickness="1">
                                    <TextBlock Text="Th&#234;me &#127912;" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="12" FontWeight="Bold"/>
                                </Border>
                            </ControlTemplate>
                        </Button.Template>
                    </Button>
                    
                    <!-- Bouton Scanner Reseau -->
                    <Button Name="btnScanNetwork" Content="Scan PRT" Margin="0,0,10,0" ToolTip="Rechercher les PRT sur le reseau" Background="#28A745" Foreground="White" FontWeight="Bold"/>

                    <Button Name="btnExternalConsole" Content=">_ Console Externe" Background="#333" Foreground="#00FF00" FontWeight="Bold" Margin="0,0,10,0" ToolTip="Ouvre une fenetre PowerShell separee (Enter-PSSession)"/>
                    <Button Name="btnPing" Content="PING (Test)" Background="#6C757D" Foreground="White" FontWeight="Bold" ToolTip="Lancer un Ping local vers la machine"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <!-- ZONE CENTRALE -->
        <Grid Grid.Row="1" Grid.Column="0" Margin="0,0,15,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <Border Grid.Row="0" Background="#1E1E1E" CornerRadius="5,5,0,0" Padding="5">
                <TextBox Name="txtOutput" IsReadOnly="True" 
                         VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                         FontFamily="Consolas" FontSize="12" 
                         Background="Transparent" Foreground="#E0E0E0" BorderThickness="0"
                         TextWrapping="Wrap"/>
            </Border>
            
            <Border Grid.Row="1" Background="#333" CornerRadius="0,0,5,5" Padding="5">
                <DockPanel>
                    <TextBlock Text="PS >" Foreground="#00FF00" FontFamily="Consolas" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <Button Name="btnSendCmd" DockPanel.Dock="Right" Content="Envoyer" Background="#555" Foreground="White" Width="60"/>
                    <TextBox Name="txtManualInput" FontFamily="Consolas" FontSize="12" Background="#222" Foreground="#FFF" BorderThickness="0" Padding="5" Margin="0,0,10,0"/>
                </DockPanel>
            </Border>
        </Grid>

        <!-- GRIDSPLITTER pour redimensionner -->
        <GridSplitter Grid.Row="1" Grid.Column="1" Width="5" Background="#DDD" HorizontalAlignment="Center" VerticalAlignment="Stretch" Cursor="SizeWE" ResizeBehavior="PreviousAndNext">
            <GridSplitter.Template>
                <ControlTemplate TargetType="GridSplitter">
                    <Border Background="#CCC" CornerRadius="2">
                        <Rectangle Width="2" Height="30" Fill="#999" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                </ControlTemplate>
            </GridSplitter.Template>
        </GridSplitter>

        <!-- PANNEAU COMMANDES -->
        <Border Name="bdrRight" Grid.Row="1" Grid.Column="2" Background="White" CornerRadius="5" Padding="10">
            <Border.Effect>
                <DropShadowEffect Color="Gray" BlurRadius="5" ShadowDepth="1" Opacity="0.3"/>
            </Border.Effect>
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <TextBlock Name="lblActions" Text="ACTIONS RAPIDES" Grid.Row="0" FontWeight="Bold" Foreground="#555" Margin="0,0,0,10" HorizontalAlignment="Center"/>
                
                <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="0,0,15,0">
                        <!-- SECTION NORMALE -->
                        <Border Name="bdrToolHeader" Background="#E3F2FD" CornerRadius="4" Padding="5" Margin="0,0,0,5">
                            <TextBlock Name="lblToolTitle" Text="BOITE A OUTILS" FontWeight="Bold" Foreground="#1565C0" HorizontalAlignment="Center"/>
                        </Border>
                        <StackPanel Name="pnlToolsNormal" Margin="0,0,0,15">
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </Grid>
        </Border>

        <!-- FOOTER -->
        <Grid Grid.Row="2" Grid.ColumnSpan="3" Margin="5,3,5,0">
            <TextBlock Text="Pr&#234;t." Name="lblFooterStatus" Foreground="Gray" FontSize="10" HorizontalAlignment="Left" VerticalAlignment="Center"/>
            <TextBlock Name="lblSessionTimer" Text="" Foreground="#007ACC" FontSize="10" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            <TextBlock Name="lblVersion" Text="By Hotline6 - v4.1 [???]" Foreground="Gray" FontSize="10" HorizontalAlignment="Right" VerticalAlignment="Center" Cursor="Hand" ToolTip="Il parait qu'un mot magique existe..."/>
        </Grid>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Host "Erreur XAML: " $_.Exception.Message
    exit
}

# Mapping
$txtComputerNumber = $window.FindName("txtComputerNumber")
$btnConnect = $window.FindName("btnConnect")
$btnDisconnect = $window.FindName("btnDisconnect")
$lblCurrentTarget = $window.FindName("lblCurrentTarget")
$lblStatus = $window.FindName("lblStatus")
$txtOutput = $window.FindName("txtOutput")
$pnlToolsNormal = $window.FindName("pnlToolsNormal")

$txtManualInput = $window.FindName("txtManualInput")
$btnSendCmd = $window.FindName("btnSendCmd")
$lblFooterStatus = $window.FindName("lblFooterStatus")
$btnPing = $window.FindName("btnPing")
$btnExternalConsole = $window.FindName("btnExternalConsole")

# UI Elements pour le Thème
$btnTheme = $window.FindName("btnTheme")
$bdrTop = $window.FindName("bdrTop")
$bdrRight = $window.FindName("bdrRight")
$lblTitle = $window.FindName("lblTitle")
$lblActions = $window.FindName("lblActions")
$lblToolTitle = $window.FindName("lblToolTitle")
$lblSessionTimer = $window.FindName("lblSessionTimer")
$lblVersion = $window.FindName("lblVersion")
$btnScanNetwork = $window.FindName("btnScanNetwork")
$chkNoPRT = $window.FindName("chkNoPRT")
$bdrToolHeader = $window.FindName("bdrToolHeader")

$script:currentSession = $null
$script:targetName = $null
$script:currentPath = ""

# Historique des commandes
$script:commandHistory = @()
$script:historyIndex = -1

# Timer de session
$script:sessionTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:sessionStartTime = $null
$script:sessionTimer.Interval = [TimeSpan]::FromSeconds(1)
$script:sessionTimer.Add_Tick({
        if ($script:sessionStartTime) {
            $elapsed = (Get-Date) - $script:sessionStartTime
            $hours = [math]::Floor($elapsed.TotalHours).ToString("00")
            $mins = $elapsed.Minutes.ToString("00")
            $secs = $elapsed.Seconds.ToString("00")
            $lblSessionTimer.Text = "Session: ${hours}:${mins}:${secs}"
        }
    })

# --- FUNCTIONS ---

function Log-Message {
    param([string]$Message, [string]$Type = "INFO", [bool]$IsCommandInput = $false)
    $timestamp = Get-Date -Format "HH:mm:ss"
    if ($IsCommandInput) {
        $pathShort = $script:currentPath
        $prefix = "PS $script:targetName ($pathShort) >"
        $formattedMsg = "`r`n$prefix $Message`r`n"
    }
    else {
        $prefix = "[$timestamp] [$Type]"
        $formattedMsg = "$prefix $Message`r`n"
    }
    $txtOutput.AppendText($formattedMsg)
    $txtOutput.ScrollToEnd()
}

function Close-CurrentSession {
    if ($script:currentSession) {
        # Sauvegarder le nom de la cible pour la suppression du profil après déconnexion
        $targetMachine = $script:targetName
        $adminUser = $env:USERNAME  # ex: adm_hotline6
        
        # Demander confirmation avant de supprimer le profil admin
        $confirmResult = [System.Windows.MessageBox]::Show(
            "Voulez-vous supprimer le profil admin ($adminUser) cree sur $targetMachine ?`n`nCela liberera de l'espace disque sur la machine distante.",
            "Nettoyage du profil admin",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        $deleteProfile = ($confirmResult -eq [System.Windows.MessageBoxResult]::Yes)
        
        # ETAPE 1: Fermer d'abord la session PSSession pour libérer les fichiers
        Remove-PSSession -Session $script:currentSession
        $script:currentSession = $null
        Log-Message "Session fermee." "INFO"
        
        # ETAPE 2: Supprimer le profil APRES avoir fermé la session
        if ($deleteProfile) {
            try {
                Log-Message "Suppression du profil admin sur $targetMachine..." "INFO"
                Log-Message "Attente de la liberation des fichiers (5s)..." "INFO"
                
                # Attendre plus longtemps pour que les fichiers soient libérés
                Start-Sleep -Seconds 5
                
                # Calculer d'abord la taille du profil via CIM (ne charge pas le profil)
                $profileSize = 0
                $sizeFormatted = "Inconnu"
                try {
                    $cimSession = New-CimSession -ComputerName $targetMachine -ErrorAction Stop
                    $userProfile = Get-CimInstance -CimSession $cimSession -Class Win32_UserProfile | Where-Object { 
                        $_.LocalPath -like "*$adminUser*" 
                    }
                    
                    if ($userProfile) {
                        $profilePath = $userProfile.LocalPath
                        Log-Message "Profil trouve: $profilePath" "INFO"
                        
                        # Obtenir la taille via un partage admin
                        $remotePath = "\\$targetMachine\C$\Users\$adminUser"
                        if (Test-Path $remotePath) {
                            try {
                                $profileSize = (Get-ChildItem -Path $remotePath -Recurse -Force -ErrorAction SilentlyContinue | 
                                    Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                                
                                if ($profileSize -ge 1GB) {
                                    $sizeFormatted = "{0:N2} Go" -f ($profileSize / 1GB)
                                }
                                elseif ($profileSize -ge 1MB) {
                                    $sizeFormatted = "{0:N2} Mo" -f ($profileSize / 1MB)
                                }
                                elseif ($profileSize -ge 1KB) {
                                    $sizeFormatted = "{0:N2} Ko" -f ($profileSize / 1KB)
                                }
                                else {
                                    $sizeFormatted = "$profileSize octets"
                                }
                                Log-Message "Taille du profil: $sizeFormatted" "INFO"
                            }
                            catch {
                                Log-Message "Impossible de calculer la taille." "WARN"
                            }
                        }
                        
                        # Supprimer via CIM (sans charger le profil)
                        try {
                            $userProfile | Remove-CimInstance -ErrorAction Stop
                            Log-Message "Profil supprime du Panneau de configuration." "SUCCESS"
                        }
                        catch {
                            Log-Message "Erreur CIM: $($_.Exception.Message)" "WARN"
                            
                            # Plan B: Créer une tâche planifiée pour supprimer le profil
                            Log-Message "Tentative alternative via tache planifiee..." "INFO"
                            
                            $scriptBlock = @"

`$profilePath = 'C:\Users\$adminUser'
`$profileListPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'

# Supprimer le dossier
if (Test-Path `$profilePath) {
    Remove-Item -Path `$profilePath -Recurse -Force -ErrorAction SilentlyContinue
}

# Nettoyer le registre
Get-ChildItem -Path `$profileListPath | ForEach-Object {
    `$path = (Get-ItemProperty -Path `$_.PSPath -ErrorAction SilentlyContinue).ProfileImagePath
    if (`$path -like '*$adminUser*') {
        Remove-Item -Path `$_.PSPath -Force -ErrorAction SilentlyContinue
    }
}
"@
                            
                            $encodedCmd = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($scriptBlock))
                            $taskName = "DeleteProfile_$adminUser"
                            
                            # Créer la tâche planifiée via CIM
                            Invoke-CimMethod -CimSession $cimSession -ClassName Win32_Process -MethodName Create -Arguments @{
                                CommandLine = "powershell.exe -ExecutionPolicy Bypass -EncodedCommand $encodedCmd"
                            } -ErrorAction SilentlyContinue
                            
                            Log-Message "Commande de suppression envoyee." "INFO"
                            Start-Sleep -Seconds 3
                        }
                    }
                    else {
                        Log-Message "Profil non trouve dans le Panneau de configuration." "INFO"
                        
                        # Vérifier si le dossier existe quand même
                        $remotePath = "\\$targetMachine\C$\Users\$adminUser"
                        if (Test-Path $remotePath) {
                            Log-Message "Dossier trouve, suppression directe..." "INFO"
                            try {
                                Remove-Item -Path $remotePath -Recurse -Force -ErrorAction Stop
                                Log-Message "Dossier supprime." "SUCCESS"
                            }
                            catch {
                                Log-Message "Erreur suppression dossier: $_" "WARN"
                            }
                        }
                        else {
                            Log-Message "Dossier deja supprime ou inexistant." "INFO"
                        }
                    }
                    
                    Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
                }
                catch {
                    Log-Message "Erreur connexion CIM: $_" "WARN"
                }
                
                # Vérifier si le profil a été supprimé
                Start-Sleep -Seconds 2
                $remotePath = "\\$targetMachine\C$\Users\$adminUser"
                if (-not (Test-Path $remotePath)) {
                    Log-Message "Profil $adminUser supprime avec succes!" "SUCCESS"
                    Log-Message "Espace disque libere: $sizeFormatted" "INFO"
                }
                else {
                    Log-Message "Le profil existe encore. Il sera supprime au prochain redemarrage." "WARN"
                }
                
                Log-Message "Nettoyage du profil admin termine." "SUCCESS"
            }
            catch {
                Log-Message "Erreur lors du nettoyage du profil: $_" "WARN"
            }
        }
        else {
            Log-Message "Nettoyage du profil admin ignore (choix utilisateur)." "INFO"
        }
    }
    # Arrêter le timer de session
    $script:sessionTimer.Stop()
    $script:sessionStartTime = $null
    $lblSessionTimer.Text = ""
    
    $btnConnect.Visibility = "Visible"
    $btnDisconnect.Visibility = "Collapsed"
    $lblCurrentTarget.Foreground = "#D9534F"
    $lblCurrentTarget.Text = "-- Deconnecte --"
    $script:targetName = $null
    $script:currentPath = ""
}

function Execute-OnSession {
    param(
        [string]$CommandStr, 
        [scriptblock]$ScriptBlk,
        [bool]$IsManualConsole = $false
    )

    if (-not $script:currentSession -or (Get-PSSession -Id $script:currentSession.Id -ErrorAction SilentlyContinue).State -ne 'Opened') {
        Log-Message "Erreur : Pas de connexion active." "ERROR"
        return
    }

    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    if (-not $IsManualConsole) {
        $lblFooterStatus.Text = "Execution..." 
    }
    
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::ContextIdle)

    try {
        $rawResults = $null
        
        if ($ScriptBlk) {
            # Exécution Bouton
            $rawResults = Invoke-Command -Session $script:currentSession -ScriptBlock $ScriptBlk -ErrorAction Stop
        }
        else {
            # Exécution Console (avec PATH)
            $wrapper = [scriptblock]::Create("$CommandStr ; Write-Output ""###PWD###`$(`$PWD.Path)""")
            $rawResults = Invoke-Command -Session $script:currentSession -ScriptBlock $wrapper -ErrorAction Stop
        }

        if ($rawResults) {
            # Traitement Path pour la console
            if ($IsManualConsole) {
                $lastItem = $rawResults | Select-Object -Last 1
                if ($lastItem -is [string] -and $lastItem -match "###PWD###(.*)") {
                    $script:currentPath = $matches[1]
                    $rawResults = $rawResults | Select-Object -SkipLast 1
                }
            }

            if ($rawResults) {
                $strOutput = $rawResults | Out-String -Width 160
                if (-not [string]::IsNullOrWhiteSpace($strOutput)) {
                    $txtOutput.AppendText($strOutput)
                }
            }
        }
        else {
            if (-not $IsManualConsole) {
                Log-Message "(Aucune donnée retournée)" "INFO" 
            }
        }
        $txtOutput.ScrollToEnd()
    }
    catch {
        Log-Message "Erreur : $_" "ERROR"
    }
    finally {
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        $lblFooterStatus.Text = "Pret."
    }
}

# --- HANDLERS ---

$btnConnect.Add_Click({
        $inputText = $txtComputerNumber.Text.Trim()
        if ([string]::IsNullOrEmpty($inputText)) {
            Log-Message "Nom de machine invalide." "WARN"
            return
        }

        # Si checkbox cochée OU si le texte contient des lettres (nom complet), utiliser tel quel
        if ($chkNoPRT.IsChecked -or $inputText -match '[a-zA-Z]') {
            $fullTarget = $inputText
        }
        else {
            # Sinon, ajouter le préfixe PRT au numéro
            $num = $inputText -replace "[^0-9]", ""
            $fullTarget = "PRT$num"
        }
        
        Log-Message "Connexion vers $fullTarget..." "CONNECT"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        Close-CurrentSession

        try {
            if (-not (Test-Connection -ComputerName $fullTarget -Count 1 -Quiet)) {
                throw "Ping echoue (Machine offline ?)."
            }

            $script:currentSession = New-PSSession -ComputerName $fullTarget -ErrorAction Stop
            $script:targetName = $fullTarget
    
            # Handshake - Recuperer infos + utilisateur connecte
            $initData = Invoke-Command -Session $script:currentSession -ScriptBlock { 
                Set-Location C:\
                
                # Recuperer l'utilisateur connecte (session active)
                $loggedUser = "Aucun"
                try {
                    $queryResult = query user 2>$null
                    if ($queryResult) {
                        # Chercher la session active
                        $activeSessions = $queryResult | Select-Object -Skip 1 | Where-Object { $_ -match 'Active|Actif' }
                        if ($activeSessions) {
                            # Parser le nom d'utilisateur (premier champ)
                            $firstLine = $activeSessions | Select-Object -First 1
                            if ($firstLine -match '^\s*>?(\S+)') {
                                $loggedUser = $matches[1]
                            }
                        }
                        else {
                            # Prendre le premier utilisateur si pas de session active
                            $firstLine = $queryResult | Select-Object -Skip 1 -First 1
                            if ($firstLine -match '^\s*>?(\S+)') {
                                $loggedUser = $matches[1] + " (Deconnecte)"
                            }
                        }
                    }
                }
                catch {
                    $loggedUser = "Inconnu"
                }
                
                [PSCustomObject]@{
                    Host       = $env:COMPUTERNAME
                    Path       = $PWD.Path
                    LoggedUser = $loggedUser
                }
            } -ErrorAction Stop
        
            $script:currentPath = $initData.Path
            $script:remoteUser = $initData.LoggedUser
            Log-Message "Session validee sur $($initData.Host). Path: $($script:currentPath)" "SUCCESS"
            Log-Message "Utilisateur connecte: $($initData.LoggedUser)" "INFO"

            $lblCurrentTarget.Text = "$fullTarget - User: $($initData.LoggedUser)"
            $lblCurrentTarget.Foreground = "Green"
            $lblStatus.Text = "Connecte (ID: $($script:currentSession.Id))"
            $btnConnect.Visibility = "Collapsed"
            $btnDisconnect.Visibility = "Visible"
        
            # Démarrer le timer de session
            $script:sessionStartTime = Get-Date
            $script:sessionTimer.Start()
        
            $txtManualInput.Focus()

        }
        catch {
            Log-Message "Echec : $_" "ERROR"
            [System.Windows.MessageBox]::Show("Connexion échouée : $_", "Erreur", "OK", "Error")
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    })

$btnDisconnect.Add_Click({ Close-CurrentSession })

$actionSend = {
    $cmd = $txtManualInput.Text
    if ([string]::IsNullOrWhiteSpace($cmd)) {
        return 
    }
    
    # Ajouter à l'historique
    $script:commandHistory += $cmd
    $script:historyIndex = $script:commandHistory.Count
    
    # Intéception Enter-PSSession
    if ($cmd -match "^\s*Enter-PSSession") {
        Log-Message "Commande 'Enter-PSSession' détectée." "INFO"
        Log-Message "Ouverture d'une fenêtre externe dédiée..." "ACTION"
        Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", "$cmd"
        $txtManualInput.Text = ""
        return
    }

    Log-Message $cmd -IsCommandInput $true
    Execute-OnSession -CommandStr $cmd -IsManualConsole $true
    $txtManualInput.Text = ""
    $txtManualInput.Focus()
}

$btnSendCmd.Add_Click($actionSend)

# Gestion des touches : Enter pour envoyer, ↑/↓ pour l'historique
# Utiliser PreviewKeyDown pour intercepter les touches avant le TextBox
$script:tempInput = ""

$txtManualInput.Add_PreviewKeyDown({ 
        param($sender, $e)
    
        if ($e.Key -eq 'Enter') { 
            $actionSend.Invoke() 
            $script:tempInput = ""
            $e.Handled = $true
        }
        elseif ($e.Key -eq 'Up') {
            # Historique precedent (remonter dans l'historique)
            if ($script:commandHistory.Count -gt 0) {
                # Sauvegarder la saisie actuelle si on commence à naviguer
                if ($script:historyIndex -eq $script:commandHistory.Count) {
                    $script:tempInput = $txtManualInput.Text
                    # Première pression: aller à la dernière commande
                    $script:historyIndex = $script:commandHistory.Count - 1
                    $txtManualInput.Text = $script:commandHistory[$script:historyIndex]
                    $txtManualInput.CaretIndex = $txtManualInput.Text.Length
                }
                elseif ($script:historyIndex -gt 0) {
                    # Continuer à remonter
                    $script:historyIndex--
                    $txtManualInput.Text = $script:commandHistory[$script:historyIndex]
                    $txtManualInput.CaretIndex = $txtManualInput.Text.Length
                }
            }
            $e.Handled = $true
        }
        elseif ($e.Key -eq 'Down') {
            # Historique suivant (descendre dans l'historique)
            if ($script:commandHistory.Count -gt 0) {
                if ($script:historyIndex -lt $script:commandHistory.Count - 1) {
                    $script:historyIndex++
                    $txtManualInput.Text = $script:commandHistory[$script:historyIndex]
                    $txtManualInput.CaretIndex = $txtManualInput.Text.Length
                }
                elseif ($script:historyIndex -eq $script:commandHistory.Count - 1) {
                    # On est à la dernière commande, revenir au texte temporaire
                    $script:historyIndex = $script:commandHistory.Count
                    $txtManualInput.Text = $script:tempInput
                    $txtManualInput.CaretIndex = $txtManualInput.Text.Length
                }
            }
            $e.Handled = $true
        }
    })

$btnPing.Add_Click({
        $n = $txtComputerNumber.Text -replace "[^0-9]", ""
        if ([string]::IsNullOrEmpty($n)) {
            Log-Message "Veuillez entrer un numero pour tester le Ping." "WARN"
            return
        }
        $t = "PRT$n"
        Log-Message "Ping vers $t en cours..." "PING"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::ContextIdle)

        try {
            $res = ping.exe $t
            $str = $res | Out-String
            Log-Message "`r`n$str" "RESULT"
        }
        catch {
            Log-Message "Ping echoue : $_" "ERROR"
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    })

$btnExternalConsole.Add_Click({
        $n = $txtComputerNumber.Text -replace "[^0-9]", ""
        if ([string]::IsNullOrEmpty($n)) {
            Log-Message "Numero requis pour la console externe." "WARN"
            return
        }
        $t = "PRT$n"
        Log-Message "Ouverture Console Externe vers $t..." "ACTION"
        Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", "Enter-PSSession -ComputerName $t"
    })

# Bouton Scanner Reseau - Detection automatique des postes
$btnScanNetwork.Add_Click({
        Log-Message "Scan du reseau en cours..." "ACTION"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $lblFooterStatus.Text = "Scan en cours..."
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::ContextIdle)

        # Fonction de ping rapide compatible PowerShell 5.1
        function Test-PingQuick {
            param([string]$ComputerName, [int]$TimeoutMs = 500)
            try {
                $ping = New-Object System.Net.NetworkInformation.Ping
                $result = $ping.Send($ComputerName, $TimeoutMs)
                return ($result.Status -eq [System.Net.NetworkInformation.IPStatus]::Success)
            }
            catch {
                return $false
            }
        }

        # Choix de la methode de scan
        $scanMethod = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Choisissez la methode de scan :`n`n1 = Active Directory (recommande)`n2 = Scan plage IP`n3 = Scan plage PRT (ancien)`n`nEntrez 1, 2 ou 3 :", 
            "Scanner le Reseau", 
            "1"
        )

        if ([string]::IsNullOrWhiteSpace($scanMethod)) {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            $lblFooterStatus.Text = "Pret."
            return
        }

        $foundMachines = @()

        switch ($scanMethod) {
            "1" {
                # === METHODE 1 : ACTIVE DIRECTORY ===
                Log-Message "Recherche des ordinateurs dans Active Directory..." "INFO"
                [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::ContextIdle)
                
                try {
                    # Filtre optionnel par prefixe
                    $prefixFilter = [Microsoft.VisualBasic.Interaction]::InputBox(
                        "Filtrer par prefixe (laisser vide pour tout) :`n`nExemples: PRT, PC, DESKTOP, LAP",
                        "Filtre prefixe",
                        "PRT"
                    )
                    
                    # Requete AD via ADSI
                    $searcher = [adsisearcher]"(objectCategory=computer)"
                    $searcher.PageSize = 1000
                    $searcher.PropertiesToLoad.AddRange(@("name", "operatingSystem", "lastLogonTimestamp", "distinguishedName"))
                    
                    $allComputers = $searcher.FindAll()
                    $total = $allComputers.Count
                    $current = 0
                    $filtered = 0
                    
                    Log-Message "Trouve $total ordinateurs dans AD. Filtrage et verification..." "INFO"
                    
                    foreach ($comp in $allComputers) {
                        $current++
                        $name = $comp.Properties["name"][0]
                        
                        # Filtrer par prefixe si specifie
                        if (-not [string]::IsNullOrWhiteSpace($prefixFilter) -and -not $name.StartsWith($prefixFilter, [StringComparison]::OrdinalIgnoreCase)) {
                            continue
                        }
                        
                        $filtered++
                        $lblFooterStatus.Text = "Test: $name ($filtered filtres / $current sur $total)"
                        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::ContextIdle)
                        
                        # Test de connectivite rapide avec ping .NET
                        $online = Test-PingQuick -ComputerName $name -TimeoutMs 300
                        
                        if ($online) {
                            # Obtenir l'IP
                            try {
                                $ip = ([System.Net.Dns]::GetHostAddresses($name) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1).IPAddressToString
                            }
                            catch {
                                $ip = "N/A" 
                            }
                            
                            $os = if ($comp.Properties["operatingSystem"]) {
                                $comp.Properties["operatingSystem"][0] 
                            }
                            else {
                                "Inconnu" 
                            }
                            
                            $foundMachines += [PSCustomObject]@{
                                Nom    = $name
                                IP     = $ip
                                OS     = $os
                                Statut = "EN LIGNE"
                            }
                            Log-Message "  [OK] $name ($ip)" "SUCCESS"
                        }
                    }
                    Log-Message "Scan termine: $filtered machines testees sur $total dans AD" "INFO"
                }
                catch {
                    Log-Message "Erreur AD : $_ - Essayez la methode 2 (scan IP)" "ERROR"
                }
            }
            
            "2" {
                # === METHODE 2 : SCAN PLAGE IP ===
                $ipRange = [Microsoft.VisualBasic.Interaction]::InputBox(
                    "Entrez la plage IP a scanner :`n`nFormat: 192.168.1.1-254`nou: 10.0.0.1-50",
                    "Scan IP",
                    "192.168.1.1-254"
                )
                
                if ([string]::IsNullOrWhiteSpace($ipRange)) {
                    $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                    $lblFooterStatus.Text = "Pret."
                    return
                }
                
                # Parser la plage IP
                if ($ipRange -match "^(\d+\.\d+\.\d+\.)(\d+)-(\d+)$") {
                    $baseIP = $matches[1]
                    $startIP = [int]$matches[2]
                    $endIP = [int]$matches[3]
                }
                else {
                    Log-Message "Format IP invalide. Utilisez: 192.168.1.1-254" "WARN"
                    $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                    $lblFooterStatus.Text = "Pret."
                    return
                }
                
                $total = $endIP - $startIP + 1
                $current = 0
                
                Log-Message "Scan de ${baseIP}${startIP} a ${baseIP}${endIP} ($total adresses)..." "INFO"
                Log-Message "Scan sequentiel en cours (peut prendre quelques minutes)..." "INFO"
                
                # Scan sequentiel avec ping rapide .NET
                for ($i = $startIP; $i -le $endIP; $i++) {
                    $current++
                    $ip = "${baseIP}${i}"
                    
                    $lblFooterStatus.Text = "Scan: $ip ($current/$total)"
                    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::ContextIdle)
                    
                    if (Test-PingQuick -ComputerName $ip -TimeoutMs 200) {
                        # Obtenir le nom
                        $hostName = "Inconnu"
                        try {
                            $hostName = ([System.Net.Dns]::GetHostEntry($ip)).HostName.Split('.')[0]
                        }
                        catch { 
                        }
                        
                        $foundMachines += [PSCustomObject]@{
                            Nom    = $hostName
                            IP     = $ip
                            OS     = "-"
                            Statut = "EN LIGNE"
                        }
                        Log-Message "  [OK] $hostName ($ip)" "SUCCESS"
                    }
                }
            }
            
            "3" {
                # === METHODE 3 : ANCIEN SCAN PRT ===
                $rangeInput = [Microsoft.VisualBasic.Interaction]::InputBox(
                    "Entrez la plage de numeros PRT a scanner:`n(ex: 1-50 ou 100-150)", 
                    "Scanner les PRT", 
                    "1-50"
                )

                if ([string]::IsNullOrWhiteSpace($rangeInput)) {
                    $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                    $lblFooterStatus.Text = "Pret."
                    return
                }

                if ($rangeInput -match "(\d+)-(\d+)") {
                    $start = [int]$matches[1]
                    $end = [int]$matches[2]
                }
                else {
                    Log-Message "Format invalide. Utilisez: debut-fin (ex: 1-50)" "WARN"
                    $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                    $lblFooterStatus.Text = "Pret."
                    return
                }

                $total = $end - $start + 1
                $current = 0

                Log-Message "Scan de PRT$start a PRT$end ($total machines)..." "INFO"

                for ($i = $start; $i -le $end; $i++) {
                    $current++
                    $prtName = "PRT$i"
                    $lblFooterStatus.Text = "Scan: $prtName ($current/$total)"
                    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::ContextIdle)

                    if (Test-PingQuick -ComputerName $prtName -TimeoutMs 300) {
                        try {
                            $ip = ([System.Net.Dns]::GetHostAddresses($prtName) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1).IPAddressToString
                        }
                        catch {
                            $ip = "N/A" 
                        }
                        
                        $foundMachines += [PSCustomObject]@{
                            Nom    = $prtName
                            IP     = $ip
                            OS     = "-"
                            Statut = "EN LIGNE"
                        }
                        Log-Message "  [OK] $prtName ($ip)" "SUCCESS"
                    }
                }
            }
            
            default {
                Log-Message "Choix invalide. Entrez 1, 2 ou 3." "WARN"
            }
        }

        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        $lblFooterStatus.Text = "Pret."

        if ($foundMachines.Count -gt 0) {
            Log-Message "Scan termine: $($foundMachines.Count) machine(s) en ligne trouvee(s)" "SUCCESS"
        
            # Afficher dans une grille pour selection
            $selected = $foundMachines | Out-GridView -Title "Machines disponibles - Double-cliquez pour selectionner" -PassThru
        
            if ($selected) {
                # Extraire le numero du nom (ex: PRT123 -> 123)
                if ($selected.Nom -match '\d+') {
                    $txtComputerNumber.Text = $matches[0]
                }
                else {
                    $txtComputerNumber.Text = $selected.Nom
                }
                Log-Message "Machine selectionnee: $($selected.Nom) ($($selected.IP))" "INFO"
            }
        }
        else {
            Log-Message "Aucune machine trouvee sur le reseau" "WARN"
            [System.Windows.MessageBox]::Show("Aucune machine trouvee sur le reseau.`n`nVerifiez votre connexion et reessayez.", "Scan termine", "OK", "Information")
        }
    })

# --- TASK BUTTONS ---

# Par défaut, on ajoute dans le panneau normal
function Add-TaskButton {
    param([string]$Name, [scriptblock]$Script, [System.Windows.Controls.StackPanel]$TargetPanel = $pnlToolsNormal)
    $btn = New-Object System.Windows.Controls.Button
    $btn.Content = $Name
    $btn.Add_Click({
            Log-Message "Lancement : $Name" "ACTION"
            Execute-OnSession -ScriptBlk $Script
        }.GetNewClosure())
    $TargetPanel.Children.Add($btn) | Out-Null
}

# --- BOUTONS UTILITAIRES ---

# Bouton Effacer Console
$btnClearConsole = New-Object System.Windows.Controls.Button
$btnClearConsole.Content = "Effacer Console"
$btnClearConsole.Background = "#6C757D"
$btnClearConsole.Foreground = "White"
$btnClearConsole.FontWeight = "Bold"
$btnClearConsole.Add_Click({
        $txtOutput.Clear()
        Log-Message "Console effacée." "INFO"
    })
$pnlToolsNormal.Children.Add($btnClearConsole) | Out-Null

# Bouton Exporter Log
$btnExportLog = New-Object System.Windows.Controls.Button
$btnExportLog.Content = "Exporter Log"
$btnExportLog.Background = "#17A2B8"
$btnExportLog.Foreground = "White"
$btnExportLog.FontWeight = "Bold"
$btnExportLog.Add_Click({
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "Fichier texte (*.txt)|*.txt"
        $saveDialog.FileName = "RSM_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $saveDialog.Title = "Exporter le log de la console"
    
        if ($saveDialog.ShowDialog() -eq $true) {
            $txtOutput.Text | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
            Log-Message "Log exporté vers : $($saveDialog.FileName)" "SUCCESS"
        }
    })
$pnlToolsNormal.Children.Add($btnExportLog) | Out-Null

# Séparateur visuel
$separator1 = New-Object System.Windows.Controls.Separator
$separator1.Margin = "0,10,0,10"
$pnlToolsNormal.Children.Add($separator1) | Out-Null

# --- BOUTONS DIAGNOSTIC ---

Add-TaskButton "IPConfig" { ipconfig /all }

# Bouton Infos Système (complet)
Add-TaskButton "Infos Systeme" {
    ""
    "============================================"
    "          INFORMATIONS SYSTEME              "
    "============================================"
    
    # OS
    $os = Get-CimInstance Win32_OperatingSystem
    "  [PC]      : $env:COMPUTERNAME"
    "  [OS]      : $($os.Caption)"
    "  [Version] : $($os.Version) (Build $($os.BuildNumber))"
    
    # CPU
    $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
    "  [CPU]     : $($cpu.Name)"
    
    # RAM
    $totalRAM = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
    $freeRAM = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
    "  [RAM]     : $freeRAM Go libre / $totalRAM Go total"
    
    # Uptime
    $uptime = (Get-Date) - $os.LastBootUpTime
    "  [Uptime]  : $($uptime.Days)j $($uptime.Hours)h $($uptime.Minutes)m"
    "  [Reboot]  : $($os.LastBootUpTime.ToString('dd/MM/yyyy HH:mm'))"
    
    # Numéro de série
    $bios = Get-CimInstance Win32_BIOS
    $serial = $bios.SerialNumber
    "  [S/N]     : $serial"
    
    # Fabricant
    $cs = Get-CimInstance Win32_ComputerSystem
    "  [Modèle]  : $($cs.Manufacturer) $($cs.Model)"
    
    "============================================"
    ""
}
Add-TaskButton "Disque C:" { Get-PSDrive C | Select-Object @{N = 'Total(Go)'; E = { '{0:N2}' -f ($_.Used / 1GB + $_.Free / 1GB) } }, @{N = 'Libre(Go)'; E = { '{0:N2}' -f ($_.Free / 1GB) } } }
Add-TaskButton "Services HS" { Get-Service | Where Status -eq 'Stopped' | Select -First 10 DisplayName, Status }

Add-TaskButton "Utilisateurs" { query user 2>&1 | Out-String } 

Add-TaskButton "BitLocker (Cle.)" { 
    try {
        $vols = Get-BitLockerVolume -ErrorAction Stop
        foreach ($v in $vols) {
            $key = $v.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }
            if ($key) {
                "Volume $($v.MountPoint) : $($key.RecoveryPassword)"
            }
            else {
                "Volume $($v.MountPoint) : Pas de clé trouvée."
            }
        }
    }
    catch {
        "Erreur : Impossible de recuperer BitLocker (Module absent ? Droits ?)"
        $_
    }
}

# Bouton EventViewer Errors
Add-TaskButton "EventViewer (Erreurs)" {
    ""
    "============================================"
    "    DERNIERES ERREURS SYSTEME (10 max)     "
    "============================================"
    ""
    
    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            Level   = 1, 2  # 1=Critical, 2=Error
        } -MaxEvents 10 -ErrorAction Stop
        
        if ($events) {
            foreach ($evt in $events) {
                $date = $evt.TimeCreated.ToString("dd/MM/yyyy HH:mm")
                $level = if ($evt.Level -eq 1) {
                    "CRITIQUE" 
                }
                else {
                    "ERREUR" 
                }
                $source = $evt.ProviderName
                $msg = $evt.Message -replace "`r`n", " " | Select-Object -First 1
                if ($msg.Length -gt 80) {
                    $msg = $msg.Substring(0, 77) + "..." 
                }
                
                "[$date] [$level]"
                "  Source: $source"
                "  $msg"
                ""
            }
        }
        else {
            "Aucune erreur critique trouvée. Système stable !"
        }
    }
    catch {
        "Impossible de lire les événements : $_"
    }
    
    "============================================"
}

Add-TaskButton "Nettoyage Complet" {
    #Vérifier l'état actuelle du poste coté hibernation
    #powercfg /a

    # Désactiver l'hibernation
    powercfg /hibernate off

    # Fonction pour supprimer les fichiers avec gestion d'erreurs
    function Remove-FilesSecure {
        param($path, $description)
        
        if (Test-Path $path) {
            Write-Output "Nettoyage: $description..." 
            try {
                $files = Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                $count = ($files | Measure-Object).Count
                $files | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                Write-Output "  -> $count fichiers supprimés" 
            }
            catch {
                Write-Output "  -> Erreur partielle (certains fichiers en cours d'utilisation)" 
            }
        }
    }


    # Dossiers temporaires Windows
    Remove-FilesSecure -path "C:\Windows\Temp\*" -description "Fichiers temporaires Windows"
    Remove-FilesSecure -path "$env:TEMP\*" -description "Fichiers temporaires utilisateur"


    Remove-FilesSecure -path "C:\Windows\Prefetch\*" -description "Fichiers Prefetch"
    Remove-FilesSecure -path "C:\Windows\SoftwareDistribution\Download\*" -description "Téléchargements Windows Update"

    Remove-FilesSecure -path "C:\Windows\Logs\*" -description "Fichiers de log Windows"
    Remove-FilesSecure -path "C:\ProgramData\Microsoft\Windows\WER\*" -description "Fichiers de log Windows"

    Remove-FilesSecure -path "C:\ProgramData\Dell\UpdateService\Downloads\*" -description "pilotes téléchargés par Dell Command"
    Remove-FilesSecure -path "C:\ProgramData\Dell\CommandUpdate\Downloads\*" -description "pilotes téléchargés par Dell Command"
    Remove-FilesSecure -path "C:\ProgramData\Dell\UpdateService\Temp\*" -description "pilotes téléchargés par Dell Command"

    Remove-FilesSecure -path "C:\ProgramData\HP\HP Image Assistant\*" -description "pilotes téléchargés par Hp Image Assist"


    # $env:LOCALAPPDATA → C:\Users\[Nom]\AppData\Local\ Données locales, non transférables
    # $env:APPDATA → C:\Users\[Nom]\AppData\Roaming\  Données "itinérantes", synchronisées entre ordinateurs
    # $env:TEMP → C:\Users\[Nom]\AppData\Local\Temp\  Fichiers temporaires purs

    # Cache navigateurs (si présents)
    Remove-FilesSecure -path "$env:LOCALAPPDATA\Microsoft\Windows\INetCache\*" -description "Cache Internet Explorer"

    # Cache des packages MECM
    Remove-FilesSecure -path "c:\windows\ccmcache\*" -description "Cache Package MECM "


    # Corbeille
    Write-Output "Vidage de la corbeille..." 
    try {
        Clear-RecycleBin -Force -ErrorAction SilentlyContinue
        Write-Output "  -> Corbeille videe"  
    }
    catch {
        Write-Output "  -> Impossible de vider complètement la corbeille" 
    }

    # Nettoyage disque Windows (optionnel)
    Write-Output ""
    Write-Output "Lancement de l'outil de nettoyage Windows..." 
    Start-Process cleanmgr.exe -ArgumentList "/sagerun:1" -NoNewWindow


    #Nettoyage des profils locaux
    # netplwiz (GUI - Risqué en remote)

    Write-Output ""
    Write-Output "=== Nettoyage terminé ===" 
    Write-Output ""
    Write-Output "Espace disque libéré." 
    # pause (Commenté pour éviter le blocage de la session distante)
}

Add-TaskButton "Update Dell (Silent)" { 
    $paths = @(
        "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe",
        "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe"
    )
    $exe = $paths | Where-Object { Test-Path $_ } | Select-Object -First 1
    
    if ($exe) {
        "Demarrage de Dell Command Update..."
        Start-Process -FilePath $exe -ArgumentList "/applyUpdates -silent -reboot=enable" -WindowStyle Hidden
        "Processus lancé en arrière-plan sur $env:COMPUTERNAME."
        "Les mises a jour vont s'installer. Redemarrage automatique si requis."
    }
    else {
        Write-Warning "Echec : 'dcu-cli.exe' est introuvable sur cette machine."
        "Verifiez que Dell Command Update est bien installe."
    }
}

Add-TaskButton "Lister Updates (Dell)" {
    # On utilise un dossier neutre à la racine pour éviter les soucis de droits/chemins réservés
    $destDir = "C:\Temp"
    if (-not (Test-Path $destDir)) {
        New-Item -Path $destDir -ItemType Directory -Force | Out-Null 
    }
    
    $reportPath = "$destDir\dcu-report.xml"
    
    $paths = @(
        "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe",
        "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe"
    )
    $exe = $paths | Where-Object { Test-Path $_ } | Select-Object -First 1
    
    if ($exe) {
        "Scan des mises a jour en cours (dcu-cli)..."
        
        if (Test-Path $reportPath) {
            Remove-Item $reportPath -Force -ErrorAction SilentlyContinue 
        }

        # Execution directe
        $output = & $exe /scan -report="$reportPath" 2>&1 | Out-String
        
        # Vérification existence et taille
        if (Test-Path $reportPath) {
            $item = Get-Item $reportPath
            if ($item.Length -gt 0) {
                $content = Get-Content $reportPath -Raw -ErrorAction SilentlyContinue
                
                if (-not [string]::IsNullOrWhiteSpace($content)) {
                    # Parsing regex
                    $matches = [regex]::Matches($content, 'name="([^"]+)"', 'IgnoreCase')
                    if ($matches.Count -eq 0) {
                        $matches = [regex]::Matches($content, '<Name>(.+?)</Name>', 'IgnoreCase')
                    }
        
                    if ($matches.Count -gt 0) {
                        ""
                        "=== MISES A JOUR DISPONIBLES ($($matches.Count)) ==="
                        foreach ($m in $matches) {
                            " [!] " + $m.Groups[1].Value
                        }
                        ""
                    }
                    else {
                        "Aucune mise à jour détectée (ou format XML inconnu)."
                    }
                }
                else {
                    "Le rapport est vide (0 contenu lu)."
                }
            }
            else {
                "Le fichier rapport a été créé mais est vide (0 octet)."
                "Sortie commande : $output"
            }
        }
        else {
            "ERREUR: Rapport non généré."
            "Sortie commande : $output"
        }
    }
    else {
        "Erreur: dcu-cli.exe introuvable sur la machine."
    }
}

# Bouton Redémarrer Machine (avec double confirmation)
$btnRestart = New-Object System.Windows.Controls.Button
$btnRestart.Content = "Redemarrer Machine"
$btnRestart.Background = "#DC3545"
$btnRestart.Foreground = "White"
$btnRestart.FontWeight = "Bold"
$btnRestart.Add_Click({
        if (-not $script:currentSession) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }
    
        # Première confirmation
        $confirm1 = [System.Windows.Forms.MessageBox]::Show(
            "Voulez-vous VRAIMENT redémarrer la machine $($script:targetName) ?`n`nCette action est irréversible !", 
            "Confirmation Redémarrage", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    
        if ($confirm1 -eq 'Yes') {
            # Deuxième confirmation (saisie du nom)
            $inputName = [Microsoft.VisualBasic.Interaction]::InputBox(
                "Pour confirmer, tapez le nom de la machine :`n$($script:targetName)", 
                "Confirmation Finale", 
                ""
            )
        
            if ($inputName -eq $script:targetName) {
                Log-Message "Redémarrage de $($script:targetName) en cours..." "ACTION"
            
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        Restart-Computer -Force
                    } -ErrorAction Stop
                
                    Log-Message "Commande de redémarrage envoyée avec succès." "SUCCESS"
                    Close-CurrentSession
                }
                catch {
                    Log-Message "Erreur lors du redémarrage : $_" "ERROR"
                }
            }
            else {
                Log-Message "Redémarrage annulé (nom incorrect)." "INFO"
            }
        }
        else {
            Log-Message "Redémarrage annulé." "INFO"
        }
    })
$pnlToolsNormal.Children.Add($btnRestart) | Out-Null

# Bouton BIOS (Logique Locale + Distante)
$btnBios = New-Object System.Windows.Controls.Button
$btnBios.Content = "Update BIOS (Custom)"
$btnBios.Add_Click({
        if (-not $script:currentSession) {
            Log-Message "Connectez-vous d'abord !" "WARN"; return 
        }

        # Interaction Locale (System.Windows.Forms pour éviter erreurs)
        $res = [System.Windows.Forms.MessageBox]::Show("Souhaitez-vous changer le mot de passe BIOS avant la mise à jour ?", "Update BIOS", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
    
        if ($res -eq 'Cancel') {
            return 
        }
    
        $biosPwdArg = $null
        if ($res -eq 'Yes') {
            $newPass = [Microsoft.VisualBasic.Interaction]::InputBox("Entrez le nouveau mot de passe BIOS :", "Changer MDP BIOS")
            if ([string]::IsNullOrWhiteSpace($newPass)) { 
                Log-Message "Annulé par l'utilisateur." "WARN"
                return 
            }
            $biosPwdArg = $newPass
        }

        # Execution Distante
        Log-Message "Lancement procédure BIOS..." "ACTION"
    
        $sb = {
            param($pwd)
        
            # 1. Changement MDP si demandé
            if ($pwd) {
                $cctkPaths = @("C:\Program Files\Dell\Command Configure\cctk.exe", "C:\Program Files (x86)\Dell\Command Configure\cctk.exe")
                $cctk = $cctkPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
                if ($cctk) {
                    Start-Process -FilePath $cctk -ArgumentList "--setuppwd=$pwd" -Wait -WindowStyle Hidden
                    "Mot de passe BIOS modifie via CCTK."
                }
                else {
                    "AVERTISSEMENT: CCTK introuvable. Mot de passe NON changé."
                }
            }

            # 2. Update DCU
            $paths = @("C:\Program Files\Dell\CommandUpdate\dcu-cli.exe", "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe")
            $exe = $paths | Where-Object { Test-Path $_ } | Select-Object -First 1
    
            if ($exe) {
                "Lancement DCU Update (Silent + Reboot si necessaire)..."
                Start-Process -FilePath $exe -ArgumentList "/applyUpdates -silent -reboot=enable" -WindowStyle Hidden
                "Commande envoyee."
            }
            else {
                "Erreur: dcu-cli.exe introuvable."
            }
        }
    
        Execute-OnSession -ScriptBlk $sb -ArgumentList @($biosPwdArg)
    })
$pnlToolsNormal.Children.Add($btnBios) | Out-Null

# Bouton Désinstaller App
$btnUninstall = New-Object System.Windows.Controls.Button
$btnUninstall.Content = "Supprimer une App"
$btnUninstall.Add_Click({
        if (-not $script:currentSession) {
            Log-Message "Connectez-vous d'abord !" "WARN"; return 
        }

        Log-Message "Récupération de la liste des logiciels (Win32_Product)... Cela peut être long." "ACTION"
    
        # Run in background to not freeze UI? The current design seems to be synchronous in click handlers mostly.
        # To keep it simple and consistent with other buttons:
        try {
            $apps = Invoke-Command -Session $script:currentSession -ScriptBlock {
                Get-CimInstance Win32_Product | Select-Object Name, Version, IdentifyingNumber, Vendor
            } -ErrorAction Stop

            if ($apps) {
                # Selection locale via Out-GridView
                $selected = $apps | Out-GridView -Title "Sélectionnez l'application à DÉSINSTALLER du poste distant" -PassThru
            
                if ($selected) {
                    $name = $selected.Name
                    $guid = $selected.IdentifyingNumber
                
                    $confirm = [System.Windows.Forms.MessageBox]::Show("Etes-vous SÛR de vouloir désinstaller :`n$name`n($guid) ?", "Confirmation Suppression", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                
                    if ($confirm -eq 'Yes') {
                        Log-Message "Lancement de la désinstallation pour $name..." "ACTION"
                    
                        $res = Invoke-Command -Session $script:currentSession -ScriptBlock {
                            param($g)
                            "--- DIAGNOSTIC DESINSTALLATION ---"
                            "GUID Cible : $g"
                        
                            try {
                                $p = Get-CimInstance Win32_Product -Filter "IdentifyingNumber='$g'" -ErrorAction Stop
                            
                                if ($p) {
                                    "Paquet trouvé : $($p.Name) (Version: $($p.Version))"
                                    "Appel de la méthode Uninstall()..."
                                
                                    # Appel CIM
                                    $ret = Invoke-CimMethod -InputObject $p -MethodName Uninstall
                                
                                    if ($ret) {
                                        "Objet retourné par Uninstall :"
                                        $ret | Out-String | ForEach-Object { "  > $_" }
                                    
                                        if ($ret.ReturnValue -eq 0) { 
                                            "SUCCES : La désinstallation a démarré avec succès." 
                                        }
                                        elseif ($ret.ReturnValue -eq 1603) {
                                            "ERREUR (1603) : Erreur fatale lors de l'installation (souvent droits ou fichier corrompu)."
                                        }
                                        elseif ($ret.ReturnValue -eq 1619) {
                                            "ERREUR (1619) : Package d'installation introuvable ou inaccessible."
                                        }
                                        else { 
                                            "ECHEC : Code retour non-null = $($ret.ReturnValue)" 
                                        }
                                    }
                                    else {
                                        "ERREUR CRITIQUE : La méthode n'a rien retourné (Null)."
                                    }
                                }
                                else {
                                    "ERREUR : Impossible de retrouver l'objet Win32_Product avec ce GUID sur la cible."
                                }
                            }
                            catch {
                                "EXCEPTION SYSTEME : $_"
                            }
                            "----------------------------------"
                        } -ArgumentList $guid
                    
                        Log-Message "$res" "RESULT"
                    }
                    else {
                        Log-Message "Action annulée." "INFO"
                    }
                }
                else {
                    Log-Message "Aucune application sélectionnée." "INFO"
                }
            }
            else {
                Log-Message "Aucune application trouvée via Win32_Product." "WARN"
            }
        }
        catch {
            Log-Message "Erreur lors de la récupération : $_" "ERROR"
        }
    })
$pnlToolsNormal.Children.Add($btnUninstall) | Out-Null
 

# Battery (STANDARD)
$btnBatt = New-Object System.Windows.Controls.Button
$btnBatt.Content = "Diag Batterie (Txt + HTML)"
$btnBatt.Add_Click({
        if (-not $script:currentSession) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        Log-Message "Analyse de la batterie..." "ACTION"
    
        # 1. Affichage direct console (Stats)
        try {
            $stats = Invoke-Command -Session $script:currentSession -ScriptBlock {
                $b = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
            
                if ($b) {
                    # Essai de récupération de capacité (souvent limité via WMI standard sans outils constructeur)
                    $charge = $b.EstimatedChargeRemaining
                
                    $statusText = switch ($b.BatteryStatus) {
                        1 {
                            "Decharge" 
                        }
                        2 {
                            "Secteur (AC) - Connecte" 
                        }
                        3 {
                            "Pleine charge" 
                        }
                        4 {
                            "Faible" 
                        }
                        5 {
                            "Critique" 
                        }
                        6 {
                            "En charge" 
                        }
                        7 {
                            "En charge (Haute)" 
                        }
                        8 {
                            "En charge (Faible)" 
                        }
                        9 {
                            "En charge (Critique)" 
                        }
                        Default {
                            "Inconnu ($($b.BatteryStatus))" 
                        }
                    }

                    ""
                    "=========================================="
                    "      DIAGNOSTIC BATTERIE INSTANTANE      "
                    "=========================================="
                    "  [+] Charge          : $charge %"
                    "  [+] Statut          : $statusText"
                    "  [+] Temps restant   : $(if($b.EstimatedRunTime -gt 700000){'Calcul en cours...'} else { "$($b.EstimatedRunTime) min" })"
                    "  [+] Voltage         : $($b.DesignVoltage) mV"
                    "=========================================="
                    ""
                }
                else {
                    " /!\ Aucune batterie detectee (PC Fixe ? Virtual Machine ?)"
                }
            } -ErrorAction Stop
        
            Log-Message "$stats" "RESULT"
        }
        catch {
            Log-Message "Erreur lecture batterie : $_" "ERROR"
        }

        # 2. Rapport complet HTML -> D:\Temp
        $remoteFile = "C:\Users\Public\battery-report.html"
        try {
            Invoke-Command -Session $script:currentSession -ScriptBlock {
                param($path)
                powercfg /batteryreport /output $path | Out-Null
            } -ArgumentList $remoteFile -ErrorAction SilentlyContinue
        
            # Dossier de destination force
            $destDir = "D:\Temp"
            if (-not (Test-Path $destDir)) {
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null 
            }
        
            $localDest = "$destDir\Rapport_Batterie_$($script:targetName).html"
            Copy-Item -FromSession $script:currentSession -Path $remoteFile -Destination $localDest -Force

            Log-Message "Rapport HTML complet sauvegarde : $localDest" "SUCCESS"
            Invoke-Item $localDest
        }
        catch {
        }
    })
$pnlToolsNormal.Children.Add($btnBatt) | Out-Null

# Bouton Copier Fichier Distant (DISCRET)
$btnCopyFile = New-Object System.Windows.Controls.Button
$btnCopyFile.Content = "Copier Fichier Distant"
$btnCopyFile.ToolTip = "Recuperer un fichier/dossier du PC distant vers votre PC"
$btnCopyFile.Add_Click({
        if (-not $script:currentSession -and -not $script:targetName) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        # Demander le chemin du fichier/dossier sur la machine distante
        $remotePath = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Entrez le chemin complet du fichier ou dossier a copier depuis $($script:targetName):`n`nExemples:`n- C:\Users\Public\Documents\fichier.docx`n- C:\Users\NomUser\Desktop`n- C:\Temp\rapport.pdf",
            "Copier Fichier Distant",
            "C:\Users\Public\Documents"
        )

        if ([string]::IsNullOrWhiteSpace($remotePath)) {
            Log-Message "Operation annulee." "INFO"
            return
        }

        # Dossier de destination local
        $defaultLocalDir = "D:\Temp\RecupFiles"
        $localDestDir = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Dossier de destination sur VOTRE PC:`n(Le dossier sera cree s'il n'existe pas)",
            "Destination locale",
            $defaultLocalDir
        )

        if ([string]::IsNullOrWhiteSpace($localDestDir)) {
            Log-Message "Operation annulee." "INFO"
            return
        }

        # Creer le dossier local si necessaire
        if (-not (Test-Path $localDestDir)) { 
            New-Item -ItemType Directory -Path $localDestDir -Force | Out-Null 
            Log-Message "Dossier cree : $localDestDir" "INFO"
        }

        Log-Message "Copie en cours depuis $($script:targetName)..." "ACTION"
        Log-Message "Source : $remotePath" "INFO"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $lblFooterStatus.Text = "Copie en cours..."
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::ContextIdle)

        try {
            # Methode 1 : Via PSSession (plus discret)
            if ($script:currentSession) {
                # Verifier si le chemin existe sur la machine distante
                $exists = Invoke-Command -Session $script:currentSession -ScriptBlock {
                    param($path)
                    Test-Path $path
                } -ArgumentList $remotePath -ErrorAction Stop

                if (-not $exists) {
                    Log-Message "ERREUR : Le chemin '$remotePath' n'existe pas sur $($script:targetName)" "ERROR"
                    return
                }

                # Obtenir le nom du fichier/dossier
                $itemName = Split-Path -Leaf $remotePath
                $localDest = Join-Path $localDestDir $itemName

                # Copier via PSSession
                Copy-Item -FromSession $script:currentSession -Path $remotePath -Destination $localDest -Recurse -Force -ErrorAction Stop

                # Calculer la taille
                if (Test-Path $localDest) {
                    $size = (Get-ChildItem -Path $localDest -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                    if ($size -ge 1GB) {
                        $sizeStr = "{0:N2} Go" -f ($size / 1GB)
                    }
                    elseif ($size -ge 1MB) {
                        $sizeStr = "{0:N2} Mo" -f ($size / 1MB)
                    }
                    elseif ($size -ge 1KB) {
                        $sizeStr = "{0:N2} Ko" -f ($size / 1KB)
                    }
                    else {
                        $sizeStr = "$size octets"
                    }
                    
                    Log-Message "SUCCES : Fichier copie vers $localDest" "SUCCESS"
                    Log-Message "Taille : $sizeStr" "INFO"
                    
                    # Proposer d'ouvrir le dossier
                    $openFolder = [System.Windows.Forms.MessageBox]::Show(
                        "Fichier copie avec succes !`n`nDestination : $localDest`nTaille : $sizeStr`n`nOuvrir le dossier ?",
                        "Copie terminee",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    )
                    if ($openFolder -eq 'Yes') {
                        Start-Process explorer.exe -ArgumentList "/select,`"$localDest`""
                    }
                }
            }
            # Methode 2 : Via partage admin (fallback)
            else {
                $uncPath = "\\$($script:targetName)\$($remotePath -replace ':', '$')"
                
                if (-not (Test-Path $uncPath)) {
                    Log-Message "ERREUR : Impossible d'acceder a '$uncPath'" "ERROR"
                    return
                }

                $itemName = Split-Path -Leaf $remotePath
                $localDest = Join-Path $localDestDir $itemName

                Copy-Item -Path $uncPath -Destination $localDest -Recurse -Force -ErrorAction Stop
                Log-Message "SUCCES : Fichier copie vers $localDest" "SUCCESS"
            }
        }
        catch {
            Log-Message "ERREUR lors de la copie : $_" "ERROR"
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            $lblFooterStatus.Text = "Pret."
        }
    })
$pnlToolsNormal.Children.Add($btnCopyFile) | Out-Null

# Bouton Créer Dossier Distant
$btnCreateFolder = New-Object System.Windows.Controls.Button
$btnCreateFolder.Content = "Creer Dossier Distant"
$btnCreateFolder.ToolTip = "Creer un nouveau dossier sur le PC distant"
$btnCreateFolder.Add_Click({
        if (-not $script:currentSession -and -not $script:targetName) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        # Demander le chemin du dossier à créer
        $folderPath = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Entrez le chemin complet du dossier a creer sur $($script:targetName):`n`nExemples:`n- C:\Users\Public\Documents\MonDossier`n- C:\Temp\NouveauDossier`n- D:\Data\Projet",
            "Creer Dossier Distant",
            "C:\Temp\NouveauDossier"
        )

        if ([string]::IsNullOrWhiteSpace($folderPath)) {
            Log-Message "Operation annulee." "INFO"
            return
        }

        Log-Message "Creation du dossier '$folderPath' sur $($script:targetName)..." "ACTION"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $lblFooterStatus.Text = "Creation du dossier en cours..."

        try {
            if ($script:currentSession) {
                # Via PSSession
                $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                    param($path)
                    if (Test-Path $path) {
                        return "EXISTS"
                    }
                    try {
                        New-Item -ItemType Directory -Path $path -Force | Out-Null
                        return "OK"
                    }
                    catch {
                        return "ERROR: $_"
                    }
                } -ArgumentList $folderPath -ErrorAction Stop

                if ($result -eq "OK") {
                    Log-Message "SUCCES : Dossier cree : $folderPath" "SUCCESS"
                    [System.Windows.MessageBox]::Show("Dossier cree avec succes !`n`n$folderPath", "Succes", "OK", "Information")
                }
                elseif ($result -eq "EXISTS") {
                    Log-Message "Le dossier existe deja : $folderPath" "WARN"
                    [System.Windows.MessageBox]::Show("Le dossier existe deja :`n`n$folderPath", "Information", "OK", "Information")
                }
                else {
                    Log-Message "ERREUR : $result" "ERROR"
                }
            }
            else {
                # Via UNC
                $uncPath = "\\$($script:targetName)\$($folderPath -replace ':', '$')"
                
                if (Test-Path $uncPath) {
                    Log-Message "Le dossier existe deja : $folderPath" "WARN"
                    [System.Windows.MessageBox]::Show("Le dossier existe deja :`n`n$folderPath", "Information", "OK", "Information")
                }
                else {
                    New-Item -ItemType Directory -Path $uncPath -Force | Out-Null
                    Log-Message "SUCCES : Dossier cree : $folderPath" "SUCCESS"
                    [System.Windows.MessageBox]::Show("Dossier cree avec succes !`n`n$folderPath", "Succes", "OK", "Information")
                }
            }
        }
        catch {
            Log-Message "ERREUR lors de la creation : $_" "ERROR"
            [System.Windows.MessageBox]::Show("Erreur lors de la creation du dossier :`n`n$_", "Erreur", "OK", "Error")
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            $lblFooterStatus.Text = "Pret."
        }
    })
$pnlToolsNormal.Children.Add($btnCreateFolder) | Out-Null

# Bouton Envoyer Fichier vers PC Distant
$btnSendFile = New-Object System.Windows.Controls.Button
$btnSendFile.Content = "Envoyer Fichier"
$btnSendFile.ToolTip = "Envoyer un fichier de votre PC vers le PC distant"
$btnSendFile.Add_Click({
        if (-not $script:currentSession -and -not $script:targetName) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        # Sélectionner le fichier local à envoyer
        $openFileDialog = New-Object Microsoft.Win32.OpenFileDialog
        $openFileDialog.Title = "Selectionner le fichier a envoyer vers $($script:targetName)"
        $openFileDialog.Filter = "Tous les fichiers (*.*)|*.*"
        
        if ($openFileDialog.ShowDialog() -ne $true) {
            Log-Message "Operation annulee." "INFO"
            return
        }
        
        $localFilePath = $openFileDialog.FileName
        $fileName = Split-Path -Leaf $localFilePath

        # Demander le dossier de destination sur le PC distant
        $remoteDest = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Entrez le dossier de destination sur $($script:targetName):`n`nLe fichier '$fileName' sera copie dans ce dossier.`n`nExemples:`n- C:\Temp`n- C:\Users\Public\Documents`n- D:\Data",
            "Destination sur PC distant",
            "C:\Temp"
        )

        if ([string]::IsNullOrWhiteSpace($remoteDest)) {
            Log-Message "Operation annulee." "INFO"
            return
        }

        Log-Message "Envoi de '$fileName' vers $($script:targetName):$remoteDest..." "ACTION"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $lblFooterStatus.Text = "Envoi du fichier en cours..."

        try {
            if ($script:currentSession) {
                # S'assurer que le dossier de destination existe
                Invoke-Command -Session $script:currentSession -ScriptBlock {
                    param($path)
                    if (-not (Test-Path $path)) {
                        New-Item -ItemType Directory -Path $path -Force | Out-Null
                    }
                } -ArgumentList $remoteDest -ErrorAction Stop

                # Copier via PSSession
                $remoteFullPath = Join-Path $remoteDest $fileName
                Copy-Item -ToSession $script:currentSession -Path $localFilePath -Destination $remoteFullPath -Force -ErrorAction Stop

                # Obtenir la taille du fichier
                $fileSize = (Get-Item $localFilePath).Length
                $sizeStr = if ($fileSize -gt 1MB) { "{0:N2} Mo" -f ($fileSize / 1MB) } else { "{0:N2} Ko" -f ($fileSize / 1KB) }

                Log-Message "SUCCES : Fichier envoye vers $remoteFullPath ($sizeStr)" "SUCCESS"
                [System.Windows.MessageBox]::Show("Fichier envoye avec succes !`n`nDestination: $remoteFullPath`nTaille: $sizeStr", "Succes", "OK", "Information")
            }
            else {
                # Via UNC
                $uncPath = "\\$($script:targetName)\$($remoteDest -replace ':', '$')"
                
                if (-not (Test-Path $uncPath)) {
                    New-Item -ItemType Directory -Path $uncPath -Force | Out-Null
                }

                $destFullPath = Join-Path $uncPath $fileName
                Copy-Item -Path $localFilePath -Destination $destFullPath -Force -ErrorAction Stop

                $fileSize = (Get-Item $localFilePath).Length
                $sizeStr = if ($fileSize -gt 1MB) { "{0:N2} Mo" -f ($fileSize / 1MB) } else { "{0:N2} Ko" -f ($fileSize / 1KB) }

                Log-Message "SUCCES : Fichier envoye vers $remoteDest\$fileName ($sizeStr)" "SUCCESS"
                [System.Windows.MessageBox]::Show("Fichier envoye avec succes !`n`nDestination: $remoteDest\$fileName`nTaille: $sizeStr", "Succes", "OK", "Information")
            }
        }
        catch {
            Log-Message "ERREUR lors de l'envoi : $_" "ERROR"
            [System.Windows.MessageBox]::Show("Erreur lors de l'envoi du fichier :`n`n$_", "Erreur", "OK", "Error")
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            $lblFooterStatus.Text = "Pret."
        }
    })
$pnlToolsNormal.Children.Add($btnSendFile) | Out-Null

# Séparateur
$separator2 = New-Object System.Windows.Controls.Separator
$separator2.Margin = "0,10,0,10"
$pnlToolsNormal.Children.Add($separator2) | Out-Null

# === NOUVELLES FONCTIONNALITES ===

# Bouton Performances Live
$btnPerfLive = New-Object System.Windows.Controls.Button
$btnPerfLive.Content = "[PERF] Performances Live"
$btnPerfLive.Background = "#FF5722"
$btnPerfLive.Foreground = "White"
$btnPerfLive.FontWeight = "Bold"
$btnPerfLive.ToolTip = "Affiche CPU/RAM/Disque en temps reel"
$btnPerfLive.Add_Click({
        if (-not $script:currentSession) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        Log-Message "Recuperation des performances sur $($script:targetName)..." "ACTION"
        
        try {
            $perf = Invoke-Command -Session $script:currentSession -ScriptBlock {
                $cpu = (Get-CimInstance Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average
                $os = Get-CimInstance Win32_OperatingSystem
                $ramUsed = [math]::Round(($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / 1MB, 2)
                $ramTotal = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
                $ramPercent = [math]::Round((($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / $os.TotalVisibleMemorySize) * 100, 1)
                
                $disk = Get-PSDrive C
                $diskUsed = [math]::Round($disk.Used / 1GB, 2)
                $diskFree = [math]::Round($disk.Free / 1GB, 2)
                $diskPercent = [math]::Round(($disk.Used / ($disk.Used + $disk.Free)) * 100, 1)
                
                # Top 5 processus CPU
                $topProc = Get-Process | Sort-Object CPU -Descending | Select-Object -First 5 Name, @{N = 'CPU%'; E = { [math]::Round($_.CPU, 1) } }, @{N = 'RAM(Mo)'; E = { [math]::Round($_.WorkingSet64 / 1MB, 1) } }
                
                [PSCustomObject]@{
                    CPU          = $cpu
                    RAMUsed      = $ramUsed
                    RAMTotal     = $ramTotal
                    RAMPercent   = $ramPercent
                    DiskUsed     = $diskUsed
                    DiskFree     = $diskFree
                    DiskPercent  = $diskPercent
                    TopProcesses = $topProc
                }
            } -ErrorAction Stop
            
            # Creer barres visuelles
            function Get-ProgressBar {
                param($percent, $width = 20)
                $filled = [math]::Floor($percent / (100 / $width))
                $empty = $width - $filled
                return "[" + ("█" * $filled) + ("░" * $empty) + "]"
            }
            
            $output = "`r`n"
            $output += "==========================================`r`n"
            $output += "     PERFORMANCES LIVE - $($script:targetName)`r`n"
            $output += "==========================================`r`n`r`n"
            $output += "  CPU:    $(Get-ProgressBar $perf.CPU) $($perf.CPU)%`r`n`r`n"
            $output += "  RAM:    $(Get-ProgressBar $perf.RAMPercent) $($perf.RAMPercent)%`r`n"
            $output += "          $($perf.RAMUsed) Go / $($perf.RAMTotal) Go`r`n`r`n"
            $output += "  DISQUE: $(Get-ProgressBar $perf.DiskPercent) $($perf.DiskPercent)%`r`n"
            $output += "          Utilise: $($perf.DiskUsed) Go | Libre: $($perf.DiskFree) Go`r`n`r`n"
            $output += "------------------------------------------`r`n"
            $output += "  TOP 5 PROCESSUS (CPU):`r`n"
            $output += $($perf.TopProcesses | Format-Table -AutoSize | Out-String)
            $output += "=========================================="
            
            Log-Message $output "RESULT"
        }
        catch {
            Log-Message "Erreur : $_" "ERROR"
        }
    })
$pnlToolsNormal.Children.Add($btnPerfLive) | Out-Null

# Bouton Programmes au Démarrage
$btnStartup = New-Object System.Windows.Controls.Button
$btnStartup.Content = "[START] Programmes Demarrage"
$btnStartup.Background = "#795548"
$btnStartup.Foreground = "White"
$btnStartup.FontWeight = "Bold"
$btnStartup.ToolTip = "Liste les programmes au demarrage Windows"
$btnStartup.Add_Click({
        if (-not $script:currentSession) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        Log-Message "Recuperation des programmes au demarrage..." "ACTION"
        
        try {
            $startups = Invoke-Command -Session $script:currentSession -ScriptBlock {
                $results = @()
                
                # Registre HKLM Run
                $hklmRun = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -ErrorAction SilentlyContinue
                if ($hklmRun) {
                    $hklmRun.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
                        $results += [PSCustomObject]@{ Nom = $_.Name; Emplacement = "HKLM\Run"; Commande = $_.Value.Substring(0, [Math]::Min(60, $_.Value.Length)) }
                    }
                }
                
                # Registre HKCU Run
                $hkcuRun = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -ErrorAction SilentlyContinue
                if ($hkcuRun) {
                    $hkcuRun.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
                        $results += [PSCustomObject]@{ Nom = $_.Name; Emplacement = "HKCU\Run"; Commande = $_.Value.Substring(0, [Math]::Min(60, $_.Value.Length)) }
                    }
                }
                
                # Taches planifiees au demarrage
                $tasks = Get-ScheduledTask | Where-Object { $_.Triggers -match 'LogonTrigger' -and $_.State -eq 'Ready' } | Select-Object -First 10 TaskName
                foreach ($t in $tasks) {
                    $results += [PSCustomObject]@{ Nom = $t.TaskName; Emplacement = "Task Scheduler"; Commande = "(Tache planifiee)" }
                }
                
                return $results
            } -ErrorAction Stop
            
            if ($startups.Count -gt 0) {
                $output = $startups | Format-Table -AutoSize | Out-String
                Log-Message "`r`n=== PROGRAMMES AU DEMARRAGE ($($startups.Count)) ===`r`n$output" "RESULT"
            }
            else {
                Log-Message "Aucun programme au demarrage trouve." "INFO"
            }
        }
        catch {
            Log-Message "Erreur : $_" "ERROR"
        }
    })
$pnlToolsNormal.Children.Add($btnStartup) | Out-Null

# Bouton Diagnostic Réseau
$btnNetDiag = New-Object System.Windows.Controls.Button
$btnNetDiag.Content = "[NET] Diag Reseau"
$btnNetDiag.Background = "#009688"
$btnNetDiag.Foreground = "White"
$btnNetDiag.FontWeight = "Bold"
$btnNetDiag.ToolTip = "Diagnostic reseau complet (DNS, Gateway, Internet)"
$btnNetDiag.Add_Click({
        if (-not $script:currentSession) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        Log-Message "Diagnostic reseau en cours sur $($script:targetName)..." "ACTION"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        
        try {
            $netInfo = Invoke-Command -Session $script:currentSession -ScriptBlock {
                $results = @()
                $results += "============================================"
                $results += "       DIAGNOSTIC RESEAU COMPLET           "
                $results += "============================================"
                $results += ""
                
                # IP Config
                $adapters = Get-NetIPConfiguration | Where-Object { $_.IPv4Address }
                foreach ($a in $adapters) {
                    $results += "  [Interface] $($a.InterfaceAlias)"
                    $results += "  [IP]        $($a.IPv4Address.IPAddress)"
                    $results += "  [Gateway]   $($a.IPv4DefaultGateway.NextHop)"
                    $results += "  [DNS]       $($a.DNSServer.ServerAddresses -join ', ')"
                    $results += ""
                }
                
                # Tests
                $results += "------ TESTS DE CONNECTIVITE ------"
                
                # Ping Gateway
                $gw = (Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway }).IPv4DefaultGateway.NextHop | Select-Object -First 1
                if ($gw) {
                    $pingGW = Test-Connection -ComputerName $gw -Count 1 -Quiet
                    $results += "  Gateway ($gw): $(if($pingGW){'OK'}else{'ECHEC'})"
                }
                
                # Ping DNS Google
                $pingDNS = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet
                $results += "  DNS Google (8.8.8.8): $(if($pingDNS){'OK'}else{'ECHEC'})"
                
                # Resolution DNS
                try {
                    $dns = Resolve-DnsName "google.com" -ErrorAction Stop
                    $results += "  Resolution DNS (google.com): OK"
                }
                catch {
                    $results += "  Resolution DNS (google.com): ECHEC"
                }
                
                # Ping Internet
                $pingNet = Test-Connection -ComputerName "google.com" -Count 1 -Quiet
                $results += "  Internet (google.com): $(if($pingNet){'OK'}else{'ECHEC'})"
                
                $results += ""
                $results += "============================================"
                
                return $results -join "`r`n"
            } -ErrorAction Stop
            
            Log-Message "`r`n$netInfo" "RESULT"
        }
        catch {
            Log-Message "Erreur : $_" "ERROR"
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    })
$pnlToolsNormal.Children.Add($btnNetDiag) | Out-Null

# Bouton Historique USB
$btnUSBHistory = New-Object System.Windows.Controls.Button
$btnUSBHistory.Content = "[USB] Historique USB"
$btnUSBHistory.Background = "#607D8B"
$btnUSBHistory.Foreground = "White"
$btnUSBHistory.FontWeight = "Bold"
$btnUSBHistory.ToolTip = "Liste tous les peripheriques USB connectes historiquement"
$btnUSBHistory.Add_Click({
        if (-not $script:currentSession) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        Log-Message "Recuperation de l'historique USB..." "ACTION"
        
        try {
            $usbHistory = Invoke-Command -Session $script:currentSession -ScriptBlock {
                $devices = @()
                
                # USB Storage
                $usbStore = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Enum\USBSTOR\*\*" -ErrorAction SilentlyContinue
                foreach ($u in $usbStore) {
                    if ($u.FriendlyName) {
                        $devices += [PSCustomObject]@{
                            Type      = "Stockage USB"
                            Nom       = $u.FriendlyName
                            Fabricant = $u.Mfg
                        }
                    }
                }
                
                # USB Devices generiques
                $usbGen = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Enum\USB\*\*" -ErrorAction SilentlyContinue | Select-Object -First 20
                foreach ($u in $usbGen) {
                    if ($u.DeviceDesc -and $u.DeviceDesc -notmatch "Hub|Controller") {
                        $devices += [PSCustomObject]@{
                            Type      = "Peripherique USB"
                            Nom       = ($u.DeviceDesc -split ';')[-1]
                            Fabricant = if ($u.Mfg) {
                                ($u.Mfg -split ';')[-1] 
                            }
                            else {
                                "Inconnu" 
                            }
                        }
                    }
                }
                
                return $devices | Select-Object -Unique -Property Nom, Type, Fabricant
            } -ErrorAction Stop
            
            if ($usbHistory.Count -gt 0) {
                $output = $usbHistory | Format-Table -AutoSize | Out-String
                Log-Message "`r`n=== HISTORIQUE USB ($($usbHistory.Count) peripheriques) ===`r`n$output" "RESULT"
            }
            else {
                Log-Message "Aucun historique USB trouve." "INFO"
            }
        }
        catch {
            Log-Message "Erreur : $_" "ERROR"
        }
    })
$pnlToolsNormal.Children.Add($btnUSBHistory) | Out-Null

# Bouton Kill Process
$btnKillProc = New-Object System.Windows.Controls.Button
$btnKillProc.Content = "[KILL] Kill Process"
$btnKillProc.Background = "#E91E63"
$btnKillProc.Foreground = "White"
$btnKillProc.FontWeight = "Bold"
$btnKillProc.ToolTip = "Tuer un processus par son nom"
$btnKillProc.Add_Click({
        if (-not $script:currentSession) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        $procName = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Entrez le nom du processus a tuer :`n(sans .exe)`n`nExemples: notepad, chrome, excel",
            "Kill Process",
            ""
        )

        if ([string]::IsNullOrWhiteSpace($procName)) {
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Voulez-vous vraiment tuer tous les processus '$procName' sur $($script:targetName) ?",
            "Confirmation",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($confirm -eq 'Yes') {
            try {
                $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                    param($name)
                    $procs = Get-Process -Name $name -ErrorAction SilentlyContinue
                    if ($procs) {
                        $count = $procs.Count
                        $procs | Stop-Process -Force -ErrorAction SilentlyContinue
                        return "OK: $count processus '$name' tues."
                    }
                    else {
                        return "Aucun processus '$name' trouve."
                    }
                } -ArgumentList $procName -ErrorAction Stop
                
                Log-Message $result "SUCCESS"
            }
            catch {
                Log-Message "Erreur : $_" "ERROR"
            }
        }
    })
$pnlToolsNormal.Children.Add($btnKillProc) | Out-Null

# === NOUVELLES FONCTIONNALITES SECURITE ===

# Séparateur Sécurité
$separatorSec = New-Object System.Windows.Controls.Separator
$separatorSec.Margin = "0,10,0,5"
$pnlToolsNormal.Children.Add($separatorSec) | Out-Null

$lblSecurityTitle = New-Object System.Windows.Controls.TextBlock
$lblSecurityTitle.Text = "SECURITE"
$lblSecurityTitle.FontWeight = "Bold"
$lblSecurityTitle.Foreground = "#D32F2F"
$lblSecurityTitle.HorizontalAlignment = "Center"
$lblSecurityTitle.Margin = "0,0,0,5"
$pnlToolsNormal.Children.Add($lblSecurityTitle) | Out-Null

# Bouton BitLocker Status
$btnBitlocker = New-Object System.Windows.Controls.Button
$btnBitlocker.Content = "[LOCK] BitLocker Status"
$btnBitlocker.Background = "#1976D2"
$btnBitlocker.Foreground = "White"
$btnBitlocker.FontWeight = "Bold"
$btnBitlocker.ToolTip = "Verifier si BitLocker est active sur tous les disques"
$btnBitlocker.Add_Click({
        if (-not $script:currentSession) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        Log-Message "Verification BitLocker sur $($script:targetName)..." "ACTION"
    
        try {
            $bitlockerInfo = Invoke-Command -Session $script:currentSession -ScriptBlock {
                $results = @()
                $volumes = Get-BitLockerVolume -ErrorAction SilentlyContinue
                if ($volumes) {
                    foreach ($vol in $volumes) {
                        $results += [PSCustomObject]@{
                            Lecteur     = $vol.MountPoint
                            Status      = $vol.ProtectionStatus
                            Methode     = ($vol.KeyProtector | Where-Object { $_.KeyProtectorType } | Select-Object -First 1).KeyProtectorType
                            Pourcentage = "$($vol.EncryptionPercentage)%"
                        }
                    }
                }
                else {
                    $results += [PSCustomObject]@{ Lecteur = "N/A"; Status = "BitLocker non disponible"; Methode = "-"; Pourcentage = "-" }
                }
                return $results
            } -ErrorAction Stop
        
            $output = "`r`n=== BITLOCKER STATUS ===`r`n"
            $output += $bitlockerInfo | Format-Table -AutoSize | Out-String
            Log-Message $output "RESULT"
        }
        catch {
            Log-Message "Erreur : $_" "ERROR"
        }
    })
$pnlToolsNormal.Children.Add($btnBitlocker) | Out-Null

# Bouton Logiciels Récents
$btnRecentSoftware = New-Object System.Windows.Controls.Button
$btnRecentSoftware.Content = "[NEW] Logiciels Recents"
$btnRecentSoftware.Background = "#388E3C"
$btnRecentSoftware.Foreground = "White"
$btnRecentSoftware.FontWeight = "Bold"
$btnRecentSoftware.ToolTip = "Lister les logiciels installes dans les 30 derniers jours"
$btnRecentSoftware.Add_Click({
        if (-not $script:currentSession) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        Log-Message "Recherche des logiciels recents sur $($script:targetName)..." "ACTION"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
    
        try {
            $recentApps = Invoke-Command -Session $script:currentSession -ScriptBlock {
                $cutoffDate = (Get-Date).AddDays(-30)
                $results = @()
            
                # Registre 64-bit
                $uninstallPaths = @(
                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
                    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
                )
            
                foreach ($path in $uninstallPaths) {
                    Get-ItemProperty $path -ErrorAction SilentlyContinue | Where-Object {
                        $_.DisplayName -and $_.InstallDate
                    } | ForEach-Object {
                        try {
                            $installDate = [DateTime]::ParseExact($_.InstallDate, "yyyyMMdd", $null)
                            if ($installDate -ge $cutoffDate) {
                                $results += [PSCustomObject]@{
                                    Nom     = $_.DisplayName.Substring(0, [Math]::Min(40, $_.DisplayName.Length))
                                    Date    = $installDate.ToString("dd/MM/yyyy")
                                    Version = if ($_.DisplayVersion) { $_.DisplayVersion } else { "-" }
                                }
                            }
                        }
                        catch {}
                    }
                }
            
                return $results | Sort-Object Date -Descending | Select-Object -Unique -Property Nom, Date, Version
            } -ErrorAction Stop
        
            if ($recentApps.Count -gt 0) {
                $output = "`r`n=== LOGICIELS INSTALLES (30 derniers jours) ===`r`n"
                $output += $recentApps | Format-Table -AutoSize | Out-String
                Log-Message $output "RESULT"
            }
            else {
                Log-Message "Aucun logiciel installe dans les 30 derniers jours." "INFO"
            }
        }
        catch {
            Log-Message "Erreur : $_" "ERROR"
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    })
$pnlToolsNormal.Children.Add($btnRecentSoftware) | Out-Null

# Bouton Centre Logiciel SCCM (Installation Silencieuse)
$btnSCCM = New-Object System.Windows.Controls.Button
$btnSCCM.Content = "[SCCM] Centre Logiciel"
$btnSCCM.Background = "#0078D7"
$btnSCCM.Foreground = "White"
$btnSCCM.FontWeight = "Bold"
$btnSCCM.ToolTip = "Installer un logiciel via SCCM (silencieux, invisible pour l'utilisateur)"
$btnSCCM.Add_Click({
        if (-not $script:currentSession) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        Log-Message "Recuperation des applications SCCM disponibles sur $($script:targetName)..." "ACTION"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait

        try {
            # Récupérer les applications disponibles via SCCM
            $apps = Invoke-Command -Session $script:currentSession -ScriptBlock {
                try {
                    $applications = Get-CimInstance -Namespace "root\ccm\ClientSDK" -ClassName CCM_Application -ErrorAction Stop | 
                    Where-Object { $_.ApplicableState -eq "Applicable" -or $_.InstallState -eq "NotInstalled" } |
                    Select-Object @{N = 'Nom'; E = { $_.Name } }, 
                    @{N = 'Version'; E = { $_.SoftwareVersion } },
                    @{N = 'Etat'; E = { $_.InstallState } },
                    @{N = 'ID'; E = { $_.Id } },
                    @{N = 'Revision'; E = { $_.Revision } },
                    @{N = 'IsMachineTarget'; E = { $_.IsMachineTarget } }
                    return $applications
                }
                catch {
                    return @{ Error = $_.Exception.Message }
                }
            } -ErrorAction Stop

            $window.Cursor = [System.Windows.Input.Cursors]::Arrow

            if ($apps.Error) {
                Log-Message "Erreur SCCM : $($apps.Error)" "ERROR"
                [System.Windows.MessageBox]::Show("Erreur lors de la recuperation des applications SCCM :`n`n$($apps.Error)`n`nVerifiez que le client SCCM est installe.", "Erreur SCCM", "OK", "Error")
                return
            }

            if (-not $apps -or $apps.Count -eq 0) {
                Log-Message "Aucune application SCCM disponible sur $($script:targetName)" "INFO"
                [System.Windows.MessageBox]::Show("Aucune application disponible dans le Centre Logiciel.`n`nLes applications ont peut-etre deja ete installees ou aucune n'est attribuee a cette machine.", "Centre Logiciel", "OK", "Information")
                return
            }

            Log-Message "$($apps.Count) application(s) SCCM trouvee(s)" "SUCCESS"

            # Afficher la liste pour sélection
            $selected = $apps | Out-GridView -Title "Applications disponibles - Double-cliquez pour installer" -PassThru

            if ($selected) {
                $confirm = [System.Windows.Forms.MessageBox]::Show(
                    "Voulez-vous installer '$($selected.Nom)' sur $($script:targetName) ?`n`nL'installation sera silencieuse (invisible pour l'utilisateur).`n`nVersion: $($selected.Version)",
                    "Confirmer Installation",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )

                if ($confirm -eq 'Yes') {
                    Log-Message "Installation silencieuse de '$($selected.Nom)' en cours..." "ACTION"
                    $window.Cursor = [System.Windows.Input.Cursors]::Wait
                    $lblFooterStatus.Text = "Installation SCCM en cours..."

                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        param($appId, $appRevision, $isMachine)
                        try {
                            $args = @{
                                Id              = $appId
                                Revision        = $appRevision
                                IsMachineTarget = $isMachine
                            }
                            $install = Invoke-CimMethod -Namespace "root\ccm\ClientSDK" -ClassName CCM_Application -MethodName Install -Arguments $args -ErrorAction Stop
                            
                            if ($install.ReturnValue -eq 0) {
                                return "OK"
                            }
                            else {
                                return "ERREUR: Code retour $($install.ReturnValue)"
                            }
                        }
                        catch {
                            return "ERREUR: $_"
                        }
                    } -ArgumentList $selected.ID, $selected.Revision, $selected.IsMachineTarget -ErrorAction Stop

                    $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                    $lblFooterStatus.Text = "Pret."

                    if ($result -eq "OK") {
                        Log-Message "Installation de '$($selected.Nom)' declenchee avec succes !" "SUCCESS"
                        [System.Windows.MessageBox]::Show("Installation declenchee avec succes !`n`nApplication: $($selected.Nom)`n`nL'installation se deroule en arriere-plan sur le poste distant.", "Succes", "OK", "Information")
                    }
                    else {
                        Log-Message "Echec installation : $result" "ERROR"
                        [System.Windows.MessageBox]::Show("Echec de l'installation :`n`n$result", "Erreur", "OK", "Error")
                    }
                }
            }
        }
        catch {
            Log-Message "Erreur SCCM : $_" "ERROR"
            [System.Windows.MessageBox]::Show("Erreur : $_", "Erreur", "OK", "Error")
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            $lblFooterStatus.Text = "Pret."
        }
    })
$pnlToolsNormal.Children.Add($btnSCCM) | Out-Null

# Bouton Extensions Navigateur
$btnExtensions = New-Object System.Windows.Controls.Button
$btnExtensions.Content = "[EXT] Extensions Navigateur"
$btnExtensions.Background = "#F57C00"
$btnExtensions.Foreground = "White"
$btnExtensions.FontWeight = "Bold"
$btnExtensions.ToolTip = "Lister les extensions Chrome/Edge installees"
$btnExtensions.Add_Click({
        if (-not $script:currentSession) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        Log-Message "Recherche des extensions navigateur sur $($script:targetName)..." "ACTION"
    
        try {
            $extensions = Invoke-Command -Session $script:currentSession -ScriptBlock {
                $results = @()
                $users = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch "Public|Default" }
            
                foreach ($user in $users) {
                    # Chrome Extensions
                    $chromePath = Join-Path $user.FullName "AppData\Local\Google\Chrome\User Data\Default\Extensions"
                    if (Test-Path $chromePath) {
                        Get-ChildItem $chromePath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                            $manifestPath = Get-ChildItem $_.FullName -Recurse -Filter "manifest.json" -ErrorAction SilentlyContinue | Select-Object -First 1
                            if ($manifestPath) {
                                try {
                                    $manifest = Get-Content $manifestPath.FullName -Raw | ConvertFrom-Json
                                    $results += [PSCustomObject]@{
                                        User       = $user.Name
                                        Navigateur = "Chrome"
                                        Extension  = if ($manifest.name.Length -gt 30) { $manifest.name.Substring(0, 30) } else { $manifest.name }
                                    }
                                }
                                catch {}
                            }
                        }
                    }
                
                    # Edge Extensions
                    $edgePath = Join-Path $user.FullName "AppData\Local\Microsoft\Edge\User Data\Default\Extensions"
                    if (Test-Path $edgePath) {
                        Get-ChildItem $edgePath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                            $manifestPath = Get-ChildItem $_.FullName -Recurse -Filter "manifest.json" -ErrorAction SilentlyContinue | Select-Object -First 1
                            if ($manifestPath) {
                                try {
                                    $manifest = Get-Content $manifestPath.FullName -Raw | ConvertFrom-Json
                                    $results += [PSCustomObject]@{
                                        User       = $user.Name
                                        Navigateur = "Edge"
                                        Extension  = if ($manifest.name.Length -gt 30) { $manifest.name.Substring(0, 30) } else { $manifest.name }
                                    }
                                }
                                catch {}
                            }
                        }
                    }
                }
            
                return $results
            } -ErrorAction Stop
        
            if ($extensions.Count -gt 0) {
                $output = "`r`n=== EXTENSIONS NAVIGATEUR ===`r`n"
                $output += $extensions | Format-Table -AutoSize | Out-String
                Log-Message $output "RESULT"
            }
            else {
                Log-Message "Aucune extension trouvee." "INFO"
            }
        }
        catch {
            Log-Message "Erreur : $_" "ERROR"
        }
    })
$pnlToolsNormal.Children.Add($btnExtensions) | Out-Null

# Bouton Logs Connexions
$btnLoginLogs = New-Object System.Windows.Controls.Button
$btnLoginLogs.Content = "[LOG] Connexions Recentes"
$btnLoginLogs.Background = "#C62828"
$btnLoginLogs.Foreground = "White"
$btnLoginLogs.FontWeight = "Bold"
$btnLoginLogs.ToolTip = "Afficher les dernieres connexions reussies et echouees (Event 4624/4625)"
$btnLoginLogs.Add_Click({
        if (-not $script:currentSession) { 
            Log-Message "Connectez-vous d'abord !" "WARN"
            return 
        }

        Log-Message "Recuperation des logs de connexion sur $($script:targetName)..." "ACTION"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
    
        try {
            $loginInfo = Invoke-Command -Session $script:currentSession -ScriptBlock {
                $results = @()
            
                # Connexions reussies (4624)
                $successEvents = Get-WinEvent -FilterHashtable @{
                    LogName = 'Security'
                    Id      = 4624
                } -MaxEvents 20 -ErrorAction SilentlyContinue | Where-Object {
                    $_.Properties[8].Value -in @(2, 10, 11)  # Interactive, RemoteInteractive, CachedInteractive
                } | Select-Object -First 5
            
                foreach ($evt in $successEvents) {
                    $results += [PSCustomObject]@{
                        Date   = $evt.TimeCreated.ToString("dd/MM HH:mm")
                        Status = "OK"
                        User   = $evt.Properties[5].Value
                        Type   = switch ($evt.Properties[8].Value) { 2 { "Local" } 10 { "RDP" } 11 { "Cache" } default { "Autre" } }
                    }
                }
            
                # Connexions echouees (4625)
                $failEvents = Get-WinEvent -FilterHashtable @{
                    LogName = 'Security'
                    Id      = 4625
                } -MaxEvents 10 -ErrorAction SilentlyContinue | Select-Object -First 5
            
                foreach ($evt in $failEvents) {
                    $results += [PSCustomObject]@{
                        Date   = $evt.TimeCreated.ToString("dd/MM HH:mm")
                        Status = "ECHEC"
                        User   = $evt.Properties[5].Value
                        Type   = "Tentative"
                    }
                }
            
                return $results | Sort-Object Date -Descending
            } -ErrorAction Stop
        
            $output = "`r`n=== LOGS CONNEXIONS (4624/4625) ===`r`n"
            $output += $loginInfo | Format-Table -AutoSize | Out-String
            Log-Message $output "RESULT"
        }
        catch {
            Log-Message "Erreur : $_" "ERROR"
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    })
$pnlToolsNormal.Children.Add($btnLoginLogs) | Out-Null


$script:currentTheme = 0 # 0=Light, 1=Dark, 2=Hacker

$btnTheme.Add_Click({
        $script:currentTheme = ($script:currentTheme + 1) % 3
    
        if ($script:currentTheme -eq 0) {
            # LIGHT
            $window.Background = "#F3F3F3"
            $bdrTop.Background = "White"
            $bdrRight.Background = "White"
            $lblTitle.Foreground = "#333"
            $lblActions.Foreground = "#555"
            $bdrToolHeader.Background = "#E3F2FD"
            $lblToolTitle.Foreground = "#1565C0"
        }
        elseif ($script:currentTheme -eq 1) {
            # DARK
            $window.Background = "#2D2D30"
            $bdrTop.Background = "#3E3E42"
            $bdrRight.Background = "#3E3E42"
            $lblTitle.Foreground = "White"
            $lblActions.Foreground = "White"
            $bdrToolHeader.Background = "#555"
            $lblToolTitle.Foreground = "#ADD8E6"
        }
        elseif ($script:currentTheme -eq 2) {
            # HACKER
            $window.Background = "Black"
            $bdrTop.Background = "#050505"
            $bdrRight.Background = "#050505"
            $lblTitle.Foreground = "#00FF00"
            $lblActions.Foreground = "#00FF00"
            $bdrToolHeader.Background = "#003300"
            $lblToolTitle.Foreground = "#00FF00"
        }
    })

# --- RACCOURCIS CLAVIER GLOBAUX ---
$window.Add_KeyDown({
        param($eventSender, $e)
    
        # Ctrl+L : Effacer console
        if ($e.Key -eq 'L' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') {
            $txtOutput.Clear()
            Log-Message "Console effacée (Ctrl+L)." "INFO"
            $e.Handled = $true
        }
    
        # Ctrl+S : Exporter log
        elseif ($e.Key -eq 'S' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') {
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Fichier texte (*.txt)|*.txt"
            $saveDialog.FileName = "RSM_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            $saveDialog.Title = "Exporter le log de la console"
        
            if ($saveDialog.ShowDialog() -eq $true) {
                $txtOutput.Text | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
                Log-Message "Log exporté vers : $($saveDialog.FileName)" "SUCCESS"
            }
            $e.Handled = $true
        }
    
        # F5 : Reconnecter
        elseif ($e.Key -eq 'F5') {
            $num = $txtComputerNumber.Text -replace "[^0-9]", ""
            if (-not [string]::IsNullOrEmpty($num)) {
                Log-Message "Reconnexion (F5)..." "INFO"
                $btnConnect.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
            }
            $e.Handled = $true
        }
    })

# Message d'accueil avec les raccourcis
Log-Message "=== Remote Session Manager v4.1 ===" "INFO"
Log-Message "Raccourcis: Ctrl+L (effacer) | Ctrl+S (exporter) | F5 (reconnecter)" "INFO"
Log-Message "Historique: haut/bas pour naviguer dans les commandes" "INFO"
Log-Message "" "INFO"
function Show-EasterEgg {
    [xml]$easterEggXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="????" Height="500" Width="600" 
        WindowStartupLocation="CenterScreen" Background="Black" Topmost="True">
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Contenu -->
        <Border Grid.Row="0" BorderBrush="#FF00FF" BorderThickness="3" CornerRadius="10" Background="#0A0A0A" Margin="20">
            <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                
                <!-- Titre avec effet -->
                <TextBlock Text="*** SECRET UNLOCKED! ***" FontSize="28" FontWeight="Bold" Foreground="#FF00FF" HorizontalAlignment="Center" FontFamily="Consolas">
                    <TextBlock.Effect>
                        <DropShadowEffect Color="#FF00FF" BlurRadius="20" ShadowDepth="0" Opacity="0.8"/>
                    </TextBlock.Effect>
                </TextBlock>
                
                <TextBlock Text="Tu as trouve le mot secret!" FontSize="16" Foreground="#00FF00" HorizontalAlignment="Center" FontFamily="Consolas" Margin="0,10,0,30"/>
                
                <!-- Message secret -->
                <Border Background="#1A001A" CornerRadius="8" Padding="20" Margin="20,0">
                    <StackPanel>
                        <TextBlock Text="*** ACHIEVEMENT UNLOCKED ***" FontSize="18" Foreground="#FFD700" HorizontalAlignment="Center" FontFamily="Consolas" FontWeight="Bold" Margin="0,0,0,15"/>
                        
                        <TextBlock Text="Tu as trouve l'Easter Egg secret!" FontSize="14" Foreground="#00FF00" HorizontalAlignment="Center" FontFamily="Consolas" Margin="0,5"/>
                        <TextBlock Text="Hotline6 te salue, guerrier du support!" FontSize="14" Foreground="#00FF00" HorizontalAlignment="Center" FontFamily="Consolas" Margin="0,5"/>
                        
                        <TextBlock Text="---" Foreground="#333" HorizontalAlignment="Center" Margin="0,15"/>
                        
                        <TextBlock Text="[GOD MODE ENABLED]" FontSize="16" Foreground="#FF4444" HorizontalAlignment="Center" FontFamily="Consolas" FontWeight="Bold">
                            <TextBlock.Effect>
                                <DropShadowEffect Color="Red" BlurRadius="10" ShadowDepth="0" Opacity="0.6"/>
                            </TextBlock.Effect>
                        </TextBlock>
                        
                        <TextBlock Text="(Mots secrets: GODMODE, HOTLINE6, 42)" FontSize="10" Foreground="#666" HorizontalAlignment="Center" FontFamily="Consolas" Margin="0,5,0,0"/>
                    </StackPanel>
                </Border>
                
                <!-- Stats cachees -->
                <Border Background="#001A00" CornerRadius="8" Padding="15" Margin="20,15,20,0">
                    <StackPanel>
                        <TextBlock Text="[STATS SECRETES]" FontSize="12" Foreground="#00AA00" HorizontalAlignment="Center" FontFamily="Consolas" FontWeight="Bold" Margin="0,0,0,10"/>
                        <TextBlock Name="lblSecretStats" Text="" FontSize="11" Foreground="#00FF00" HorizontalAlignment="Center" FontFamily="Consolas" TextWrapping="Wrap"/>
                    </StackPanel>
                </Border>
                
            </StackPanel>
        </Border>
        
        <!-- Bouton Fermer -->
        <Button Grid.Row="1" Name="btnCloseEgg" Content="[ RESPECT + ]" Background="#FF00FF" Foreground="Black" 
                HorizontalAlignment="Center" Padding="40,12" Margin="0,0,0,20" FontFamily="Consolas" FontWeight="Bold" FontSize="14" Cursor="Hand"/>
    </Grid>
</Window>
"@
    $eggReader = (New-Object System.Xml.XmlNodeReader $easterEggXaml)
    try {
        $eggWindow = [Windows.Markup.XamlReader]::Load($eggReader)
        
        # Stats secretes
        $lblSecretStats = $eggWindow.FindName("lblSecretStats")
        $statsText = "Commandes executees: $($script:commandHistory.Count)`r`n"
        if ($script:sessionStartTime) {
            $elapsed = (Get-Date) - $script:sessionStartTime
            $statsText += "Temps de session: $($elapsed.Hours)h $($elapsed.Minutes)m $($elapsed.Seconds)s`r`n"
        }
        $statsText += "Machine actuelle: $(if($script:targetName){$script:targetName}else{'Aucune'})`r`n"
        $statsText += "Date: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"
        $lblSecretStats.Text = $statsText
        
        # Bouton fermer
        $btnCloseEgg = $eggWindow.FindName("btnCloseEgg")
        $btnCloseEgg.Add_Click({ $eggWindow.Close() })
        
        # Jouer un son
        [console]::beep(523, 100); [console]::beep(659, 100); [console]::beep(784, 100); [console]::beep(1047, 200)
        
        Log-Message "*** SECRET UNLOCKED! Easter Egg decouvert! ***" "SUCCESS"
        $eggWindow.ShowDialog() | Out-Null
        
        # === ACTIVER LE THEME SECRET MATRIX ===
        Log-Message "*** THEME MATRIX DEBLOQUE! ***" "SUCCESS"
        
        $script:currentTheme = 3
        $matrixBg = "#0D0D0D"
        $matrixAccent = "#FF00FF"
        $matrixAccent2 = "#00FFFF"
        $matrixGreen = "#00FF00"
        
        $window.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($matrixBg)
        $bdrTop.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#1A1A2E")
        $bdrRight.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#0F0F1A")
        $txtOutput.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#050510")
        $txtOutput.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($matrixGreen)
        $lblTitle.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($matrixAccent)
        $lblSessionTimer.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($matrixAccent)
        $lblVersion.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($matrixAccent2)
        $lblStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($matrixAccent2)
        $lblCurrentTarget.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($matrixGreen)
        $btnConnect.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($matrixAccent)
        $btnConnect.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("Black")
        $btnDisconnect.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FF3366")
        $btnTheme.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($matrixAccent2)
        $btnScanNetwork.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($matrixGreen)
        $btnScanNetwork.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("Black")
        $btnPing.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#6600FF")
        $btnExternalConsole.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#1A1A2E")
        $btnExternalConsole.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($matrixGreen)
        $lblFooterStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($matrixGreen)
        
        Log-Message "" "INFO"
        Log-Message "=========================================" "INFO"
        Log-Message "   THEME MATRIX ACTIVE !" "SUCCESS"
        Log-Message "   Bienvenue dans la Matrix, THE HOTLINER DU 3150 !" "INFO"
        Log-Message "=========================================" "INFO"
    }
    catch {
        Log-Message "Easter Egg erreur: $_" "ERROR"
    }
}

# Handler pour detecter le mot secret dans le champ de saisie
$txtComputerNumber.Add_TextChanged({
        $text = $txtComputerNumber.Text.ToUpper()
        if ($text -eq "GODMODE" -or $text -eq "HOTLINE6") {
            $txtComputerNumber.Text = ""
            Show-EasterEgg
        }
    })


# --- EASTER EGG : MODE TROLL ---
$lblVersion.Add_MouseLeftButtonDown({
        # Interface XAML du Mode Troll - Version Matrix ULTIMATE
        [xml]$trollXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MODE TROLL - Hotline6 Special Edition" Height="700" Width="750" 
        WindowStartupLocation="CenterScreen" Background="Black" Topmost="True">
    
    <Window.Resources>
        <Style TargetType="Button" x:Key="TrollButton">
            <Setter Property="Padding" Value="15,10"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="Background" Value="#1A1A1A"/>
            <Setter Property="Foreground" Value="#00FF00"/>
            <Setter Property="BorderBrush" Value="#00AA00"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="10"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Background" Value="#0D0D0D"/>
            <Setter Property="Foreground" Value="#00FF00"/>
            <Setter Property="BorderBrush" Value="#00FF00"/>
            <Setter Property="BorderThickness" Value="2"/>
            <Setter Property="FontFamily" Value="Consolas"/>
        </Style>
    </Window.Resources>

    <Border BorderBrush="#00FF00" BorderThickness="2" CornerRadius="8" Background="#050505" Margin="5">
        <Grid Margin="15">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- Header avec effet glow -->
            <StackPanel Grid.Row="0" HorizontalAlignment="Center" Margin="0,5,0,15">
                <TextBlock Text="[ MODE TROLL ULTIMATE ]" FontSize="32" FontWeight="Bold" Foreground="#00FF00" HorizontalAlignment="Center" FontFamily="Consolas">
                    <TextBlock.Effect>
                        <DropShadowEffect Color="#00FF00" BlurRadius="15" ShadowDepth="0" Opacity="0.7"/>
                    </TextBlock.Effect>
                </TextBlock>
                <TextBlock Name="lblCreditsLink" Text="Hotline6 Special Edition v4.1" FontSize="11" Foreground="#00AA00" HorizontalAlignment="Center" FontFamily="Consolas" Margin="0,3,0,0" Cursor="Hand" ToolTip="Voir les credits"/>
                
                <!-- Message Easter Egg -->
                <Border Background="#1A0A1A" CornerRadius="5" Padding="10,5" Margin="0,8,0,0" BorderBrush="#FF00FF" BorderThickness="1">
                    <StackPanel>
                        <TextBlock Text="*** EASTER EGG CACHE ***" FontSize="10" Foreground="#FF00FF" HorizontalAlignment="Center" FontFamily="Consolas" FontWeight="Bold"/>
                        <TextBlock Text="Indice: Tape le 'code divin' dans le champ..." FontSize="9" Foreground="#AA00AA" HorizontalAlignment="Center" FontFamily="Consolas" FontStyle="Italic" Margin="0,3,0,0"/>
                    </StackPanel>
                </Border>
            </StackPanel>

            <!-- Cible -->
            <Border Grid.Row="1" Background="#0A0A0A" BorderBrush="#00AA00" BorderThickness="1" CornerRadius="5" Padding="10" Margin="0,0,0,10">
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                    <TextBlock Text="CIBLE: " Foreground="#00FF00" FontWeight="Bold" FontFamily="Consolas" FontSize="14" VerticalAlignment="Center"/>
                    <TextBlock Name="lblTrollTarget" Text="[ AUCUNE SESSION ]" Foreground="#FF4444" FontSize="16" FontFamily="Consolas" FontWeight="Bold" VerticalAlignment="Center"/>
                </StackPanel>
            </Border>

            <!-- Contenu Principal avec Tabs -->
            <TabControl Grid.Row="2" Background="#0A0A0A" BorderBrush="#00AA00" Margin="0,5,0,5">
                
                <!-- Tab 1: Message Popup -->
                <TabItem Header="📨 MESSAGE" Background="#1A1A1A" Foreground="#00FF00" FontFamily="Consolas" FontWeight="Bold">
                    <Border Background="#050505" Padding="15">
                        <StackPanel>
                            <TextBlock Text="MESSAGE A ENVOYER:" Foreground="#00FF00" FontWeight="Bold" Margin="0,0,0,8" FontFamily="Consolas"/>
                            <TextBox Name="txtTrollMessage" Height="100" TextWrapping="Wrap" AcceptsReturn="True" Text="Ceci est un message de la Hotline Informatique !"/>
                            <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                                <TextBlock Text="TITRE:" Foreground="#00AA00" VerticalAlignment="Center" Margin="0,0,10,0" FontFamily="Consolas" FontWeight="Bold"/>
                                <TextBox Name="txtTrollTitle" Width="250" Text="Message Important"/>
                            </StackPanel>
                            <Button Name="btnSendMessage" Content="📨 ENVOYER MESSAGE" Style="{StaticResource TrollButton}" Background="#00AA00" Foreground="Black" Margin="0,15,0,0" HorizontalAlignment="Center" Padding="30,12"/>
                        </StackPanel>
                    </Border>
                </TabItem>

                <!-- Tab 2: Sons -->
                <TabItem Header="SONS" Background="#1A1A1A" Foreground="#00FF00" FontFamily="Consolas" FontWeight="Bold">
                    <Border Background="#050505" Padding="15">
                        <StackPanel>
                            <TextBlock Text="EFFETS SONORES:" Foreground="#00FF00" FontWeight="Bold" Margin="0,0,0,15" FontFamily="Consolas"/>
                            <WrapPanel HorizontalAlignment="Center">
                                <Button Name="btnBeep" Content="BEEP" Style="{StaticResource TrollButton}" Width="140"/>
                                <Button Name="btnBeepSpam" Content="BEEP SPAM (x10)" Style="{StaticResource TrollButton}" Width="160"/>
                                <Button Name="btnSoundError" Content="SON ERREUR" Style="{StaticResource TrollButton}" Width="140"/>
                                <Button Name="btnSoundNotif" Content="SON NOTIF" Style="{StaticResource TrollButton}" Width="140"/>
                            </WrapPanel>
                            <Border BorderBrush="#333" BorderThickness="0,1,0,0" Margin="0,20,0,15" Padding="0,15,0,0">
                                <StackPanel>
                                    <TextBlock Text="TEXT-TO-SPEECH (Faire parler le PC):" Foreground="#00FF00" FontWeight="Bold" Margin="0,0,0,10" FontFamily="Consolas"/>
                                    <TextBox Name="txtTTS" Height="60" TextWrapping="Wrap" Text="Bonjour, du 3150 !"/>
                                    <Button Name="btnSpeak" Content="FAIRE PARLER" Style="{StaticResource TrollButton}" Background="#FF6600" Foreground="Black" Margin="0,10,0,0" HorizontalAlignment="Center" Padding="30,12"/>
                                </StackPanel>
                            </Border>
                        </StackPanel>
                    </Border>
                </TabItem>

                <!-- Tab 3: Ecran -->
                <TabItem Header="ECRAN" Background="#1A1A1A" Foreground="#00FF00" FontFamily="Consolas" FontWeight="Bold">
                    <Border Background="#050505" Padding="15">
                        <StackPanel>
                            <TextBlock Text="MANIPULATIONS ECRAN:" Foreground="#00FF00" FontWeight="Bold" Margin="0,0,0,15" FontFamily="Consolas"/>
                            <WrapPanel HorizontalAlignment="Center">
                                <Button Name="btnRotate90" Content="ROTATION 90°" Style="{StaticResource TrollButton}" Width="150"/>
                                <Button Name="btnRotate180" Content="ROTATION 180°" Style="{StaticResource TrollButton}" Width="150"/>
                                <Button Name="btnRotateNormal" Content="NORMAL (0°)" Style="{StaticResource TrollButton}" Width="150"/>
                            </WrapPanel>
                            <Border BorderBrush="#333" BorderThickness="0,1,0,0" Margin="0,20,0,15" Padding="0,15,0,0">
                                <StackPanel>
                                    <TextBlock Text="FOND D'ECRAN:" Foreground="#00FF00" FontWeight="Bold" Margin="0,0,0,10" FontFamily="Consolas"/>
                                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                        <Button Name="btnWallpaperBlack" Content="NOIR" Style="{StaticResource TrollButton}" Width="100"/>
                                        <Button Name="btnWallpaperBSOD" Content="BSOD FAKE" Style="{StaticResource TrollButton}" Width="120"/>
                                        <Button Name="btnWallpaperMatrix" Content="MATRIX" Style="{StaticResource TrollButton}" Width="110"/>
                                    </StackPanel>
                                </StackPanel>
                            </Border>
                        </StackPanel>
                    </Border>
                </TabItem>

                <!-- Tab 4: Peripheriques -->
                <TabItem Header="PERIPH" Background="#1A1A1A" Foreground="#00FF00" FontFamily="Consolas" FontWeight="Bold">
                    <Border Background="#050505" Padding="15">
                        <StackPanel>
                            <TextBlock Text="PERIPHERIQUES:" Foreground="#00FF00" FontWeight="Bold" Margin="0,0,0,15" FontFamily="Consolas"/>
                            <WrapPanel HorizontalAlignment="Center">
                                <Button Name="btnInvertMouse" Content="INVERSER SOURIS" Style="{StaticResource TrollButton}" Width="170"/>
                                <Button Name="btnNormalMouse" Content="SOURIS NORMALE" Style="{StaticResource TrollButton}" Width="170"/>
                            </WrapPanel>
                            <Border BorderBrush="#333" BorderThickness="0,1,0,0" Margin="0,20,0,15" Padding="0,15,0,0">
                                <StackPanel>
                                    <TextBlock Text="LECTEUR CD/DVD:" Foreground="#00FF00" FontWeight="Bold" Margin="0,0,0,10" FontFamily="Consolas"/>
                                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                        <Button Name="btnEjectCD" Content="OUVRIR CD" Style="{StaticResource TrollButton}" Width="140"/>
                                        <Button Name="btnCloseCD" Content="FERMER CD" Style="{StaticResource TrollButton}" Width="140"/>
                                    </StackPanel>
                                </StackPanel>
                            </Border>
                        </StackPanel>
                    </Border>
                </TabItem>

                <!-- Tab 5: Web -->
                <TabItem Header="WEB" Background="#1A1A1A" Foreground="#00FF00" FontFamily="Consolas" FontWeight="Bold">
                    <Border Background="#050505" Padding="15">
                        <StackPanel>
                            <TextBlock Text="OUVRIR PAGE WEB:" Foreground="#00FF00" FontWeight="Bold" Margin="0,0,0,15" FontFamily="Consolas"/>
                            <WrapPanel HorizontalAlignment="Center">
                                <Button Name="btnRickRoll" Content="RICK ROLL" Style="{StaticResource TrollButton}" Background="#FF0000" Foreground="White" Width="150"/>
                                <Button Name="btnNyanCat" Content="NYAN CAT" Style="{StaticResource TrollButton}" Background="#FF69B4" Foreground="White" Width="150"/>
                            </WrapPanel>
                            <Border BorderBrush="#333" BorderThickness="0,1,0,0" Margin="0,20,0,15" Padding="0,15,0,0">
                                <StackPanel>
                                    <TextBlock Text="URL PERSONNALISEE:" Foreground="#00FF00" FontWeight="Bold" Margin="0,0,0,10" FontFamily="Consolas"/>
                                    <TextBox Name="txtCustomURL" Text="https://www.google.com"/>
                                    <Button Name="btnOpenURL" Content="OUVRIR UN URL" Style="{StaticResource TrollButton}" Margin="0,10,0,0" HorizontalAlignment="Center" Padding="30,12"/>
                                </StackPanel>
                            </Border>
                        </StackPanel>
                    </Border>
                </TabItem>

                <!-- Tab 6: PS TOOLS (Commandes PowerShell Cachees) -->
                <TabItem Header="PS TOOLS" Background="#1A1A1A" Foreground="#00FF00" FontFamily="Consolas" FontWeight="Bold">
                    <Border Background="#050505" Padding="15">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel>
                                <TextBlock Text="COMMANDES POWERSHELL RAPIDES:" Foreground="#00FF00" FontWeight="Bold" Margin="0,0,0,15" FontFamily="Consolas"/>
                                
                                <!-- Section Navigation -->
                                <TextBlock Text="NAVIGATION:" Foreground="#00AA00" FontWeight="Bold" Margin="0,0,0,8" FontFamily="Consolas"/>
                                <WrapPanel HorizontalAlignment="Center" Margin="0,0,0,15">
                                    <Button Name="btnPsOneDrive" Content="OneDrive BRGM" Style="{StaticResource TrollButton}" Width="160" ToolTip="cd OneDrive - BRGM"/>
                                    <Button Name="btnPsDesktop" Content="Bureau User" Style="{StaticResource TrollButton}" Width="160" ToolTip="Aller au Bureau de l'utilisateur"/>
                                    <Button Name="btnPsDocuments" Content="Documents User" Style="{StaticResource TrollButton}" Width="160" ToolTip="Aller aux Documents"/>
                                    <Button Name="btnPsDownloads" Content="Telechargements" Style="{StaticResource TrollButton}" Width="160" ToolTip="Aller au dossier Telechargements"/>
                                    <Button Name="btnPsAppData" Content="AppData" Style="{StaticResource TrollButton}" Width="160" ToolTip="Aller dans AppData"/>
                                    <Button Name="btnPsTemp" Content="Temp" Style="{StaticResource TrollButton}" Width="160" ToolTip="Aller dans le dossier Temp"/>
                                </WrapPanel>
                                
                                <Border BorderBrush="#333" BorderThickness="0,1,0,0" Margin="0,10,0,10" Padding="0,15,0,0">
                                    <StackPanel>
                                        <!-- Section Profils -->
                                        <TextBlock Text="PROFILS UTILISATEUR:" Foreground="#00AA00" FontWeight="Bold" Margin="0,0,0,8" FontFamily="Consolas"/>
                                        <WrapPanel HorizontalAlignment="Center" Margin="0,0,0,15">
                                            <Button Name="btnPsListProfiles" Content="Lister Profils" Style="{StaticResource TrollButton}" Width="160" ToolTip="Lister tous les profils utilisateur"/>
                                            <Button Name="btnPsCurrentUser" Content="User Connecte" Style="{StaticResource TrollButton}" Width="160" ToolTip="Afficher l'utilisateur connecte"/>
                                            <Button Name="btnPsGoToUserProfile" Content="Profil User" Style="{StaticResource TrollButton}" Width="160" ToolTip="Aller dans le profil de l'utilisateur connecte"/>
                                        </WrapPanel>
                                    </StackPanel>
                                </Border>

                                <Border BorderBrush="#333" BorderThickness="0,1,0,0" Margin="0,10,0,10" Padding="0,15,0,0">
                                    <StackPanel>
                                        <!-- Section Systeme -->
                                        <TextBlock Text="SYSTEME:" Foreground="#00AA00" FontWeight="Bold" Margin="0,0,0,8" FontFamily="Consolas"/>
                                        <WrapPanel HorizontalAlignment="Center" Margin="0,0,0,15">
                                            <Button Name="btnPsServices" Content="Services" Style="{StaticResource TrollButton}" Width="160" ToolTip="Get-Service"/>
                                            <Button Name="btnPsProcesses" Content="Processus" Style="{StaticResource TrollButton}" Width="160" ToolTip="Get-Process Top 10"/>
                                            <Button Name="btnPsDiskSpace" Content="Espace Disque" Style="{StaticResource TrollButton}" Width="160" ToolTip="Get-PSDrive"/>
                                            <Button Name="btnPsEnvVars" Content="Variables Env" Style="{StaticResource TrollButton}" Width="160" ToolTip="Variables d'environnement"/>
                                        </WrapPanel>
                                    </StackPanel>
                                </Border>

                                <Border BorderBrush="#333" BorderThickness="0,1,0,0" Margin="0,10,0,10" Padding="0,15,0,0">
                                    <StackPanel>
                                        <!-- Section Commande Custom -->
                                        <TextBlock Text="COMMANDE PERSONNALISEE:" Foreground="#00AA00" FontWeight="Bold" Margin="0,0,0,8" FontFamily="Consolas"/>
                                        <TextBox Name="txtPsCustomCmd" Height="80" TextWrapping="Wrap" AcceptsReturn="True" Text="Get-ChildItem -Path . | Select-Object Name, Length, LastWriteTime"/>
                                        <Button Name="btnPsRunCustom" Content="EXECUTER" Style="{StaticResource TrollButton}" Background="#FF6600" Foreground="Black" Margin="0,10,0,0" HorizontalAlignment="Center" Padding="40,12"/>
                                    </StackPanel>
                                </Border>

                                <!-- Exemples de commandes -->
                                <Border BorderBrush="#333" BorderThickness="0,1,0,0" Margin="0,10,0,0" Padding="0,15,0,0">
                                    <StackPanel>
                                        <TextBlock Text="EXEMPLES DE COMMANDES:" Foreground="#666" FontWeight="Bold" Margin="0,0,0,8" FontFamily="Consolas"/>
                                        <TextBlock Text="cd '.\OneDrive - BRGM'" Foreground="#00AA00" FontFamily="Consolas" FontSize="11" Margin="5,2"/>
                                        <TextBlock Text="Get-ChildItem -Recurse -Filter *.docx" Foreground="#00AA00" FontFamily="Consolas" FontSize="11" Margin="5,2"/>
                                        <TextBlock Text="Get-Process | Sort-Object CPU -Descending | Select -First 10" Foreground="#00AA00" FontFamily="Consolas" FontSize="11" Margin="5,2"/>
                                        <TextBlock Text="Get-EventLog -LogName System -Newest 20" Foreground="#00AA00" FontFamily="Consolas" FontSize="11" Margin="5,2"/>
                                    </StackPanel>
                                </Border>
                            </StackPanel>
                        </ScrollViewer>
                    </Border>
                </TabItem>

                <!-- Tab 7: CHAOS -->
                <TabItem Header="CHAOS" Background="#1A1A1A" Foreground="#FF0000" FontFamily="Consolas" FontWeight="Bold">
                    <Border Background="#050505" Padding="15">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel>
                                <TextBlock Text="MODE CHAOS - UTILISER AVEC PRECAUTION!" Foreground="#FF0000" FontWeight="Bold" Margin="0,0,0,15" FontFamily="Consolas" HorizontalAlignment="Center"/>
                                
                                <!-- Souris Folle -->
                                <Border BorderBrush="#333" BorderThickness="1" CornerRadius="5" Padding="10" Margin="0,0,0,10" Background="#0A0A0A">
                                    <StackPanel>
                                        <TextBlock Text="SOURIS FOLLE" Foreground="#FF6600" FontWeight="Bold" FontFamily="Consolas"/>
                                        <TextBlock Text="La souris bouge toute seule pendant 5 secondes" Foreground="#888" FontFamily="Consolas" FontSize="10" Margin="0,3,0,8"/>
                                        <Button Name="btnCrazyCursor" Content="ACTIVER SOURIS FOLLE" Style="{StaticResource TrollButton}" Background="#FF6600" Foreground="Black"/>
                                    </StackPanel>
                                </Border>

                                <!-- Caps Lock Disco -->
                                <Border BorderBrush="#333" BorderThickness="1" CornerRadius="5" Padding="10" Margin="0,0,0,10" Background="#0A0A0A">
                                    <StackPanel>
                                        <TextBlock Text="CAPS LOCK DISCO" Foreground="#FF00FF" FontWeight="Bold" FontFamily="Consolas"/>
                                        <TextBlock Text="Active/desactive Caps Lock rapidement (10 fois)" Foreground="#888" FontFamily="Consolas" FontSize="10" Margin="0,3,0,8"/>
                                        <Button Name="btnCapsLockDisco" Content="CAPS LOCK DISCO" Style="{StaticResource TrollButton}" Background="#FF00FF" Foreground="Black"/>
                                    </StackPanel>
                                </Border>

                                <!-- Volume Aleatoire -->
                                <Border BorderBrush="#333" BorderThickness="1" CornerRadius="5" Padding="10" Margin="0,0,0,10" Background="#0A0A0A">
                                    <StackPanel>
                                        <TextBlock Text="VOLUME ALEATOIRE" Foreground="#00BFFF" FontWeight="Bold" FontFamily="Consolas"/>
                                        <TextBlock Text="Change le volume du PC de maniere aleatoire" Foreground="#888" FontFamily="Consolas" FontSize="10" Margin="0,3,0,8"/>
                                        <Button Name="btnRandomVolume" Content="CHANGER VOLUME" Style="{StaticResource TrollButton}" Background="#00BFFF" Foreground="Black"/>
                                    </StackPanel>
                                </Border>

                                <!-- Fausse Mise à Jour Windows -->
                                <Border BorderBrush="#333" BorderThickness="1" CornerRadius="5" Padding="10" Margin="0,0,0,10" Background="#0A0A0A">
                                    <StackPanel>
                                        <TextBlock Text="FAUSSE MISE A JOUR WINDOWS" Foreground="#0078D7" FontWeight="Bold" FontFamily="Consolas"/>
                                        <TextBlock Text="Ouvre une page web simulant une MAJ Windows" Foreground="#888" FontFamily="Consolas" FontSize="10" Margin="0,3,0,8"/>
                                        <Button Name="btnFakeUpdate" Content="FAUSSE MAJ WINDOWS" Style="{StaticResource TrollButton}" Background="#0078D7" Foreground="White"/>
                                    </StackPanel>
                                </Border>

                                <!-- Taskbar Hide -->
                                <Border BorderBrush="#333" BorderThickness="1" CornerRadius="5" Padding="10" Margin="0,0,0,10" Background="#0A0A0A">
                                    <StackPanel>
                                        <TextBlock Text="BARRE DES TACHES" Foreground="#FFD700" FontWeight="Bold" FontFamily="Consolas"/>
                                        <TextBlock Text="Cacher/Afficher la barre des taches" Foreground="#888" FontFamily="Consolas" FontSize="10" Margin="0,3,0,8"/>
                                        <WrapPanel HorizontalAlignment="Center">
                                            <Button Name="btnHideTaskbar" Content="CACHER" Style="{StaticResource TrollButton}" Background="#FFD700" Foreground="Black" Width="120"/>
                                            <Button Name="btnShowTaskbar" Content="AFFICHER" Style="{StaticResource TrollButton}" Background="#32CD32" Foreground="Black" Width="120"/>
                                        </WrapPanel>
                                    </StackPanel>
                                </Border>
                            </StackPanel>
                        </ScrollViewer>
                    </Border>
                </TabItem>

                <!-- Tab 8: CAPTURE -->
                <TabItem Header="CAPTURE" Background="#1A1A1A" Foreground="#00FF00" FontFamily="Consolas" FontWeight="Bold">
                    <Border Background="#050505" Padding="15">
                        <StackPanel>
                            <TextBlock Text="CAPTURE D'ECRAN DISTANTE" Foreground="#00FF00" FontWeight="Bold" Margin="0,0,0,15" FontFamily="Consolas" HorizontalAlignment="Center"/>
                            
                            <Border BorderBrush="#00AA00" BorderThickness="1" CornerRadius="5" Padding="15" Background="#0A0A0A">
                                <StackPanel>
                                    <TextBlock Text="Prendre une capture d'ecran du poste distant" Foreground="#00AA00" FontFamily="Consolas" HorizontalAlignment="Center" Margin="0,0,0,10"/>
                                    <TextBlock Text="La capture sera sauvegardee et ouverte automatiquement" Foreground="#666" FontFamily="Consolas" FontSize="10" HorizontalAlignment="Center" Margin="0,0,0,15"/>
                                    <Button Name="btnTrollScreenshot" Content="CAPTURER L'ECRAN" Style="{StaticResource TrollButton}" Background="#00AA00" Foreground="Black" HorizontalAlignment="Center" Padding="40,15" FontSize="14"/>
                                </StackPanel>
                            </Border>

                            <Border BorderBrush="#333" BorderThickness="0,1,0,0" Margin="0,20,0,0" Padding="0,15,0,0">
                                <StackPanel>
                                    <TextBlock Text="ENREGISTREMENT (Coming Soon)" Foreground="#333" FontWeight="Bold" FontFamily="Consolas" HorizontalAlignment="Center"/>
                                    <TextBlock Text="Enregistrement video de l'ecran distant" Foreground="#333" FontFamily="Consolas" FontSize="10" HorizontalAlignment="Center" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>
                        </StackPanel>
                    </Border>
                </TabItem>

            </TabControl>

            <!-- Bouton Fermer -->
            <Button Grid.Row="3" Name="btnCloseTroll" Content="[ FERMER ]" Background="#222" Foreground="#00FF00" BorderBrush="#00FF00" BorderThickness="1" HorizontalAlignment="Center" Padding="40,12" Margin="0,10,0,5" FontFamily="Consolas" FontWeight="Bold" FontSize="14" Cursor="Hand"/>
        </Grid>
    </Border>
</Window>
"@

        $trollReader = (New-Object System.Xml.XmlNodeReader $trollXaml)
        try {
            $trollWindow = [Windows.Markup.XamlReader]::Load($trollReader)
        }
        catch {
            [System.Windows.MessageBox]::Show("Erreur chargement Mode Troll: $_", "Erreur")
            return
        }

        # Elements - Tab Message
        $lblTrollTarget = $trollWindow.FindName("lblTrollTarget")
        $txtTrollMessage = $trollWindow.FindName("txtTrollMessage")
        $txtTrollTitle = $trollWindow.FindName("txtTrollTitle")
        $btnSendMessage = $trollWindow.FindName("btnSendMessage")
        $btnCloseTroll = $trollWindow.FindName("btnCloseTroll")
        
        # Elements - Tab Sons
        $btnBeep = $trollWindow.FindName("btnBeep")
        $btnBeepSpam = $trollWindow.FindName("btnBeepSpam")
        $btnSoundError = $trollWindow.FindName("btnSoundError")
        $btnSoundNotif = $trollWindow.FindName("btnSoundNotif")
        $txtTTS = $trollWindow.FindName("txtTTS")
        $btnSpeak = $trollWindow.FindName("btnSpeak")
        
        # Elements - Tab Ecran
        $btnRotate90 = $trollWindow.FindName("btnRotate90")
        $btnRotate180 = $trollWindow.FindName("btnRotate180")
        $btnRotateNormal = $trollWindow.FindName("btnRotateNormal")
        $btnWallpaperBlack = $trollWindow.FindName("btnWallpaperBlack")
        $btnWallpaperBSOD = $trollWindow.FindName("btnWallpaperBSOD")
        $btnWallpaperMatrix = $trollWindow.FindName("btnWallpaperMatrix")
        
        # Elements - Tab Peripheriques
        $btnInvertMouse = $trollWindow.FindName("btnInvertMouse")
        $btnNormalMouse = $trollWindow.FindName("btnNormalMouse")
        $btnEjectCD = $trollWindow.FindName("btnEjectCD")
        $btnCloseCD = $trollWindow.FindName("btnCloseCD")
        
        # Elements - Tab Web
        $btnRickRoll = $trollWindow.FindName("btnRickRoll")
        $btnNyanCat = $trollWindow.FindName("btnNyanCat")
        $txtCustomURL = $trollWindow.FindName("txtCustomURL")
        $btnOpenURL = $trollWindow.FindName("btnOpenURL")
        
        # Elements - Tab PS Tools
        $btnPsOneDrive = $trollWindow.FindName("btnPsOneDrive")
        $btnPsDesktop = $trollWindow.FindName("btnPsDesktop")
        $btnPsDocuments = $trollWindow.FindName("btnPsDocuments")
        $btnPsDownloads = $trollWindow.FindName("btnPsDownloads")
        $btnPsAppData = $trollWindow.FindName("btnPsAppData")
        $btnPsTemp = $trollWindow.FindName("btnPsTemp")
        $btnPsListProfiles = $trollWindow.FindName("btnPsListProfiles")
        $btnPsCurrentUser = $trollWindow.FindName("btnPsCurrentUser")
        $btnPsGoToUserProfile = $trollWindow.FindName("btnPsGoToUserProfile")
        $btnPsServices = $trollWindow.FindName("btnPsServices")
        $btnPsProcesses = $trollWindow.FindName("btnPsProcesses")
        $btnPsDiskSpace = $trollWindow.FindName("btnPsDiskSpace")
        $btnPsEnvVars = $trollWindow.FindName("btnPsEnvVars")
        $txtPsCustomCmd = $trollWindow.FindName("txtPsCustomCmd")
        $btnPsRunCustom = $trollWindow.FindName("btnPsRunCustom")

        # Afficher la cible actuelle
        if ($script:currentSession -and $script:targetName) {
            $lblTrollTarget.Text = $script:targetName
            $lblTrollTarget.Foreground = "#00ff00"
        }

        # === FONCTION HELPER POUR EXECUTER SUR LA CIBLE ===
        $script:TrollExecute = {
            param($ScriptBlock, $Args)
            if (-not $script:currentSession) {
                [System.Windows.MessageBox]::Show("Connecte-toi d'abord a une machine !", "Pas de cible", "OK", "Warning")
                return $null
            }
            try {
                $result = Invoke-Command -Session $script:currentSession -ScriptBlock $ScriptBlock -ArgumentList $Args -ErrorAction Stop
                return $result
            }
            catch {
                [System.Windows.MessageBox]::Show("Erreur: $_", "Echec!", "OK", "Error")
                return $null
            }
        }

        # === TAB MESSAGE ===
        $btnSendMessage.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord a une machine !", "Pas de cible", "OK", "Warning")
                    return
                }

                $message = $txtTrollMessage.Text
                $title = $txtTrollTitle.Text

                if ([string]::IsNullOrWhiteSpace($message)) {
                    [System.Windows.MessageBox]::Show("Entre un message !", "Message vide", "OK", "Warning")
                    return
                }

                try {
                    $fullMessage = "[$title]`n`n$message"
                
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        param($msgText)
                        $output = msg * /TIME:300 $msgText 2>&1
                        if ($LASTEXITCODE -eq 0) {
                            return "OK" 
                        }
                        else {
                            return "ERREUR: $output" 
                        }
                    } -ArgumentList $fullMessage -ErrorAction Stop

                    if ($result -eq "OK") {
                        Log-Message "[TROLL] Message envoye a $($script:targetName): $message" "ACTION"
                        [System.Windows.MessageBox]::Show("Message envoye!", "Troll reussi!", "OK", "Information")
                    }
                    else {
                        [System.Windows.MessageBox]::Show("Probleme: $result", "Attention", "OK", "Warning")
                    }
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec!", "OK", "Error")
                }
            })

        # === TAB SONS ===
        # Beep simple - via tâche planifiée dans le contexte utilisateur
        $btnBeep.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\beep_troll.ps1"
                        @'
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Beeper {
    [DllImport("kernel32.dll")]
    public static extern bool Beep(int frequency, int duration);
}
"@
[Beeper]::Beep(800, 500)
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollBeep_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 1500
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[TROLL] Beep envoye a $($script:targetName)" "ACTION"
                    [System.Windows.MessageBox]::Show("Beep envoye!", "Troll !", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Beep Spam - multiple beeps
        $btnBeepSpam.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord !", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\beep_spam.ps1"
                        @'
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class BeeperSpam {
    [DllImport("kernel32.dll")]
    public static extern bool Beep(int frequency, int duration);
}
"@
for ($i = 0; $i -lt 10; $i++) {
    $freq = Get-Random -Minimum 300 -Maximum 2000
    [BeeperSpam]::Beep($freq, 200)
    Start-Sleep -Milliseconds 100
}
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollBeepSpam_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 4000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[TROLL] Beep SPAM envoye a $($script:targetName)" "ACTION"
                    [System.Windows.MessageBox]::Show("Beep Spam envoye!", "Troll!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Son Erreur Windows - via lecteur audio
        $btnSoundError.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\error_sound.ps1"
                        @'
$player = New-Object System.Media.SoundPlayer
$player.SoundLocation = "C:\Windows\Media\Windows Critical Stop.wav"
$player.PlaySync()
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollError_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 2000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[TROLL] Son ERREUR envoye a $($script:targetName)" "ACTION"
                    [System.Windows.MessageBox]::Show("Son erreur envoye!", "Troll!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Son Notification
        $btnSoundNotif.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\notif_sound.ps1"
                        @'
$player = New-Object System.Media.SoundPlayer
$player.SoundLocation = "C:\Windows\Media\Windows Notify System Generic.wav"
$player.PlaySync()
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollNotif_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 2000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[TROLL] Son NOTIF envoye a $($script:targetName)" "ACTION"
                    [System.Windows.MessageBox]::Show("Son notification envoye!", "Troll!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Text-to-Speech - via tâche planifiée
        $btnSpeak.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                $textToSpeak = $txtTTS.Text
                if ([string]::IsNullOrWhiteSpace($textToSpeak)) {
                    return 
                }
            
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        param($txt)
                        $scriptPath = "$env:TEMP\tts_troll.ps1"
                        @"
Add-Type -AssemblyName System.Speech
`$synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
`$synth.Rate = 0
`$synth.Volume = 100
`$synth.Speak('$($txt -replace "'", "''")')
"@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollTTS_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 5000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ArgumentList $textToSpeak -ErrorAction Stop
                    Log-Message "[TROLL] TTS envoye a $($script:targetName): $textToSpeak" "ACTION"
                    [System.Windows.MessageBox]::Show("TTS envoye!", "Troll!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # === TAB ECRAN ===
        # Rotation ecran via tâche planifiée - ROTATION 90°
        $btnRotate90.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\rotate_screen.ps1"
                        # Script de rotation d'écran complet
                        @'
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class ScreenRotation {
    [DllImport("user32.dll")]
    public static extern int EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE devMode);
    
    [DllImport("user32.dll")]
    public static extern int ChangeDisplaySettings(ref DEVMODE devMode, int flags);
    
    public const int ENUM_CURRENT_SETTINGS = -1;
    public const int CDS_UPDATEREGISTRY = 0x01;
    public const int CDS_TEST = 0x02;
    public const int DISP_CHANGE_SUCCESSFUL = 0;
    public const int DMDO_DEFAULT = 0;
    public const int DMDO_90 = 1;
    public const int DMDO_180 = 2;
    public const int DMDO_270 = 3;
    
    [StructLayout(LayoutKind.Sequential)]
    public struct DEVMODE {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmDeviceName;
        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;
        public int dmPositionX;
        public int dmPositionY;
        public int dmDisplayOrientation;
        public int dmDisplayFixedOutput;
        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmFormName;
        public short dmLogPixels;
        public int dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;
        public int dmICMMethod;
        public int dmICMIntent;
        public int dmMediaType;
        public int dmDitherType;
        public int dmReserved1;
        public int dmReserved2;
        public int dmPanningWidth;
        public int dmPanningHeight;
    }
    
    public static void Rotate(int orientation) {
        DEVMODE dm = new DEVMODE();
        dm.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));
        EnumDisplaySettings(null, ENUM_CURRENT_SETTINGS, ref dm);
        
        int temp = dm.dmPelsWidth;
        dm.dmPelsWidth = dm.dmPelsHeight;
        dm.dmPelsHeight = temp;
        dm.dmDisplayOrientation = orientation;
        
        ChangeDisplaySettings(ref dm, CDS_UPDATEREGISTRY);
    }
}
"@
[ScreenRotation]::Rotate(1)
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollRotate90_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 2000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[TROLL] Rotation 90 appliquee sur $($script:targetName)" "ACTION"
                    [System.Windows.MessageBox]::Show("Rotation 90 degres appliquee!", "Troll!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # ROTATION 180°
        $btnRotate180.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\rotate_180.ps1"
                        @'
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class ScreenRotation180 {
    [DllImport("user32.dll")]
    public static extern int EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE devMode);
    
    [DllImport("user32.dll")]
    public static extern int ChangeDisplaySettings(ref DEVMODE devMode, int flags);
    
    public const int ENUM_CURRENT_SETTINGS = -1;
    public const int CDS_UPDATEREGISTRY = 0x01;
    
    [StructLayout(LayoutKind.Sequential)]
    public struct DEVMODE {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmDeviceName;
        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;
        public int dmPositionX;
        public int dmPositionY;
        public int dmDisplayOrientation;
        public int dmDisplayFixedOutput;
        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmFormName;
        public short dmLogPixels;
        public int dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;
        public int dmICMMethod;
        public int dmICMIntent;
        public int dmMediaType;
        public int dmDitherType;
        public int dmReserved1;
        public int dmReserved2;
        public int dmPanningWidth;
        public int dmPanningHeight;
    }
    
    public static void Rotate() {
        DEVMODE dm = new DEVMODE();
        dm.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));
        EnumDisplaySettings(null, ENUM_CURRENT_SETTINGS, ref dm);
        dm.dmDisplayOrientation = 2; // 180 degrees
        ChangeDisplaySettings(ref dm, CDS_UPDATEREGISTRY);
    }
}
"@
[ScreenRotation180]::Rotate()
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollRotate180_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 2000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[TROLL] Rotation 180 appliquee sur $($script:targetName)" "ACTION"
                    [System.Windows.MessageBox]::Show("Rotation 180 degres appliquee!", "Troll!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # ROTATION NORMAL (0°)
        $btnRotateNormal.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\rotate_normal.ps1"
                        @'
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class ScreenRotationNormal {
    [DllImport("user32.dll")]
    public static extern int EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE devMode);
    
    [DllImport("user32.dll")]
    public static extern int ChangeDisplaySettings(ref DEVMODE devMode, int flags);
    
    public const int ENUM_CURRENT_SETTINGS = -1;
    public const int CDS_UPDATEREGISTRY = 0x01;
    
    [StructLayout(LayoutKind.Sequential)]
    public struct DEVMODE {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmDeviceName;
        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;
        public int dmPositionX;
        public int dmPositionY;
        public int dmDisplayOrientation;
        public int dmDisplayFixedOutput;
        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmFormName;
        public short dmLogPixels;
        public int dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;
        public int dmICMMethod;
        public int dmICMIntent;
        public int dmMediaType;
        public int dmDitherType;
        public int dmReserved1;
        public int dmReserved2;
        public int dmPanningWidth;
        public int dmPanningHeight;
    }
    
    public static void Rotate() {
        DEVMODE dm = new DEVMODE();
        dm.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));
        EnumDisplaySettings(null, ENUM_CURRENT_SETTINGS, ref dm);
        dm.dmDisplayOrientation = 0; // Normal
        ChangeDisplaySettings(ref dm, CDS_UPDATEREGISTRY);
    }
}
"@
[ScreenRotationNormal]::Rotate()
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollRotateNormal_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 2000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[TROLL] Rotation remise a 0 sur $($script:targetName)" "ACTION"
                    [System.Windows.MessageBox]::Show("Rotation remise a la normale!", "Troll!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Wallpaper Noir
        $btnWallpaperBlack.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                Invoke-Command -Session $script:currentSession -ScriptBlock {
                    # Creer une image noire
                    Add-Type -AssemblyName System.Drawing
                    $bmp = New-Object System.Drawing.Bitmap(1920, 1080)
                    $g = [System.Drawing.Graphics]::FromImage($bmp)
                    $g.Clear([System.Drawing.Color]::Black)
                    $path = "$env:TEMP\black_wallpaper.bmp"
                    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Bmp)
                    $g.Dispose()
                    $bmp.Dispose()
                
                    # Appliquer le wallpaper
                    Add-Type @"
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@
                    [Wallpaper]::SystemParametersInfo(0x0014, 0, $path, 0x01 -bor 0x02)
                } -ErrorAction SilentlyContinue
                Log-Message "[TROLL] Wallpaper NOIR applique sur $($script:targetName)" "ACTION"
            })

        # Wallpaper BSOD Fake
        $btnWallpaperBSOD.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                Invoke-Command -Session $script:currentSession -ScriptBlock {
                    Add-Type -AssemblyName System.Drawing
                    $bmp = New-Object System.Drawing.Bitmap(1920, 1080)
                    $g = [System.Drawing.Graphics]::FromImage($bmp)
                    $g.Clear([System.Drawing.Color]::FromArgb(0, 120, 215))
                    $font = New-Object System.Drawing.Font("Segoe UI", 28)
                    $fontSmall = New-Object System.Drawing.Font("Segoe UI", 14)
                    $brush = [System.Drawing.Brushes]::White
                    $g.DrawString(":(", $font, $brush, 100, 200)
                    $g.DrawString("Your PC ran into a problem and needs to restart.", $fontSmall, $brush, 100, 280)
                    $g.DrawString("We're just collecting some error info, and then we'll restart for you.", $fontSmall, $brush, 100, 320)
                    $g.DrawString("0% complete", $fontSmall, $brush, 100, 400)
                    $g.DrawString("Stop code: CRITICAL_PROCESS_DIED", $fontSmall, $brush, 100, 480)
                    $path = "$env:TEMP\bsod_fake.bmp"
                    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Bmp)
                    $g.Dispose()
                    $bmp.Dispose()
                
                    Add-Type @"
using System.Runtime.InteropServices;
public class WallpaperBSOD {
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@
                    [WallpaperBSOD]::SystemParametersInfo(0x0014, 0, $path, 0x01 -bor 0x02)
                } -ErrorAction SilentlyContinue
                Log-Message "[TROLL] Wallpaper BSOD FAKE applique sur $($script:targetName)" "ACTION"
            })

        # Wallpaper Matrix
        $btnWallpaperMatrix.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                Invoke-Command -Session $script:currentSession -ScriptBlock {
                    Add-Type -AssemblyName System.Drawing
                    $bmp = New-Object System.Drawing.Bitmap(1920, 1080)
                    $g = [System.Drawing.Graphics]::FromImage($bmp)
                    $g.Clear([System.Drawing.Color]::Black)
                    $font = New-Object System.Drawing.Font("Consolas", 12)
                    $rnd = New-Object System.Random
                    $chars = "0123456789ABCDEFabcdef!@#$%"
                    for ($x = 0; $x -lt 1920; $x += 15) {
                        for ($y = 0; $y -lt 1080; $y += 18) {
                            $c = $chars[$rnd.Next($chars.Length)]
                            $green = $rnd.Next(100, 256)
                            $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb($green / 3, $green, $green / 3))
                            $g.DrawString($c, $font, $brush, $x, $y)
                            $brush.Dispose()
                        }
                    }
                    $path = "$env:TEMP\matrix_wallpaper.bmp"
                    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Bmp)
                    $g.Dispose()
                    $bmp.Dispose()
                
                    Add-Type @"
using System.Runtime.InteropServices;
public class WallpaperMatrix {
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@
                    [WallpaperMatrix]::SystemParametersInfo(0x0014, 0, $path, 0x01 -bor 0x02)
                } -ErrorAction SilentlyContinue
                Log-Message "[TROLL] Wallpaper MATRIX applique sur $($script:targetName)" "ACTION"
            })

        # === TAB PERIPHERIQUES ===
        # Inverser Souris - via tâche interactive
        $btnInvertMouse.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\invert_mouse.ps1"
                        @'
Add-Type @"
using System.Runtime.InteropServices;
public class MouseInvert {
    [DllImport("user32.dll")]
    public static extern bool SwapMouseButton(bool fSwap);
}
"@
[MouseInvert]::SwapMouseButton($true)
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollMouse_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 1500
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[TROLL] Souris INVERSEE sur $($script:targetName)" "ACTION"
                    [System.Windows.MessageBox]::Show("Souris inversee!", "Troll!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Souris Normale - via tâche interactive
        $btnNormalMouse.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\normal_mouse.ps1"
                        @'
Add-Type @"
using System.Runtime.InteropServices;
public class MouseNormal {
    [DllImport("user32.dll")]
    public static extern bool SwapMouseButton(bool fSwap);
}
"@
[MouseNormal]::SwapMouseButton($false)
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollMouseNormal_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 1500
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[TROLL] Souris REMISE A NORMAL sur $($script:targetName)" "ACTION"
                    [System.Windows.MessageBox]::Show("Souris remise a la normale!", "Troll!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Ouvrir CD
        $btnEjectCD.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                Invoke-Command -Session $script:currentSession -ScriptBlock {
                    $shell = New-Object -ComObject Shell.Application
                    $shell.Namespace(17).Items() | Where-Object { $_.Type -eq "CD Drive" } | ForEach-Object { $_.InvokeVerb("Eject") }
                } -ErrorAction SilentlyContinue
                Log-Message "[TROLL] CD OUVERT sur $($script:targetName)" "ACTION"
            })

        # Fermer CD
        $btnCloseCD.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                Invoke-Command -Session $script:currentSession -ScriptBlock {
                    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class CDTray {
    [DllImport("winmm.dll", EntryPoint="mciSendStringA")]
    public static extern int mciSendString(string lpstrCommand, string lpstrReturnString, int uReturnLength, IntPtr hwndCallback);
}
"@
                    [CDTray]::mciSendString("set cdaudio door closed", $null, 0, [IntPtr]::Zero)
                } -ErrorAction SilentlyContinue
                Log-Message "[TROLL] CD FERME sur $($script:targetName)" "ACTION"
            })

        # === TAB WEB ===
        # Rick Roll - via tâche interactive
        $btnRickRoll.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        # Récupérer l'utilisateur connecté
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollRickRoll_$(Get-Random)"
                            # Créer tâche qui s'exécute sous l'utilisateur connecté avec /IT (Interactive Token)
                            schtasks /create /tn $taskName /tr "cmd /c start https://www.youtube.com/watch?v=dQw4w9WgXcQ" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 2000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                            return "OK"
                        }
                        else {
                            return "Aucun utilisateur connecte"
                        }
                    } -ErrorAction Stop
                    Log-Message "[TROLL] RICK ROLL envoye a $($script:targetName)!" "ACTION"
                    [System.Windows.MessageBox]::Show("Never gonna give you up!", "Rick Roll!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Nyan Cat
        $btnNyanCat.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollNyanCat_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "cmd /c start https://www.nyan.cat/" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 2000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[TROLL] NYAN CAT envoye a $($script:targetName)!" "ACTION"
                    [System.Windows.MessageBox]::Show("Nyan Cat lance!", "Troll!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # URL Custom
        $btnOpenURL.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                $url = $txtCustomURL.Text
                if ([string]::IsNullOrWhiteSpace($url)) {
                    return 
                }
            
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        param($targetUrl)
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollURL_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "cmd /c start $targetUrl" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 2000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ArgumentList $url -ErrorAction Stop
                    Log-Message "[TROLL] URL ouverte sur $($script:targetName): $url" "ACTION"
                    [System.Windows.MessageBox]::Show("URL ouverte!", "Troll!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # === TAB PS TOOLS ===
        
        # Navigation - OneDrive BRGM
        $btnPsOneDrive.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $user = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($user) {
                            $username = $user.Split('\')[1]
                            $oneDrivePath = "C:\Users\$username\OneDrive - BRGM"
                            if (Test-Path $oneDrivePath) {
                                Set-Location $oneDrivePath
                                return "OK: Navigue vers $oneDrivePath"
                            }
                            else {
                                return "ERREUR: OneDrive BRGM non trouve ($oneDrivePath)"
                            }
                        }
                        return "ERREUR: Aucun utilisateur connecte"
                    } -ErrorAction Stop
                    Log-Message "[PS TOOLS] $result" "ACTION"
                    [System.Windows.MessageBox]::Show($result, "OneDrive BRGM", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Navigation - Bureau
        $btnPsDesktop.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $user = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($user) {
                            $username = $user.Split('\')[1]
                            $path = "C:\Users\$username\Desktop"
                            if (Test-Path $path) {
                                Set-Location $path
                                Get-ChildItem | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize | Out-String
                            }
                            else {
                                "Dossier non trouve" 
                            }
                        }
                    } -ErrorAction Stop
                    Log-Message "[PS TOOLS] Bureau utilisateur:`r`n$result" "ACTION"
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Navigation - Documents
        $btnPsDocuments.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord !", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $user = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($user) {
                            $username = $user.Split('\')[1]
                            $path = "C:\Users\$username\Documents"
                            if (Test-Path $path) {
                                Set-Location $path
                                Get-ChildItem | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize | Out-String
                            }
                            else {
                                "Dossier non trouve" 
                            }
                        }
                    } -ErrorAction Stop
                    Log-Message "[PS TOOLS] Documents utilisateur:`r`n$result" "ACTION"
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Navigation - Telechargements
        $btnPsDownloads.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $user = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($user) {
                            $username = $user.Split('\')[1]
                            $path = "C:\Users\$username\Downloads"
                            if (Test-Path $path) {
                                Set-Location $path
                                Get-ChildItem | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize | Out-String
                            }
                            else {
                                "Dossier non trouve" 
                            }
                        }
                    } -ErrorAction Stop
                    Log-Message "[PS TOOLS] Telechargements:`r`n$result" "ACTION"
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Navigation - AppData
        $btnPsAppData.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $user = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($user) {
                            $username = $user.Split('\')[1]
                            $path = "C:\Users\$username\AppData"
                            if (Test-Path $path) {
                                Set-Location $path
                                Get-ChildItem | Select-Object Name, LastWriteTime | Format-Table -AutoSize | Out-String
                            }
                            else {
                                "Dossier non trouve" 
                            }
                        }
                    } -ErrorAction Stop
                    Log-Message "[PS TOOLS] AppData:`r`n$result" "ACTION"
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Navigation - Temp
        $btnPsTemp.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        Set-Location $env:TEMP
                        $items = Get-ChildItem | Measure-Object
                        $size = (Get-ChildItem -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum / 1MB
                        "Dossier Temp: $($env:TEMP)`r`nNombre de fichiers: $($items.Count)`r`nTaille totale: $([math]::Round($size, 2)) Mo"
                    } -ErrorAction Stop
                    Log-Message "[PS TOOLS] Temp:`r`n$result" "ACTION"
                    [System.Windows.MessageBox]::Show($result, "Dossier Temp", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Profils - Lister tous les profils
        $btnPsListProfiles.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        Get-ChildItem "C:\Users" -Directory | Select-Object Name, LastWriteTime | Format-Table -AutoSize | Out-String
                    } -ErrorAction Stop
                    Log-Message "[PS TOOLS] Profils utilisateurs:`r`n$result" "ACTION"
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Profils - Utilisateur connecte
        $btnPsCurrentUser.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $user = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        $sessions = query user 2>&1 | Out-String
                        "Utilisateur connecte: $user`r`n`r`nSessions actives:`r`n$sessions"
                    } -ErrorAction Stop
                    Log-Message "[PS TOOLS] Utilisateur:`r`n$result" "ACTION"
                    [System.Windows.MessageBox]::Show($result, "Utilisateur Connecte", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Profils - Aller dans le profil utilisateur
        $btnPsGoToUserProfile.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $user = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($user) {
                            $username = $user.Split('\')[1]
                            $path = "C:\Users\$username"
                            Set-Location $path
                            Get-ChildItem | Select-Object Name, @{N = 'Type'; E = { if ($_.PSIsContainer) {
                                        'Dossier' 
                                    }
                                    else {
                                        'Fichier' 
                                    } } 
                            }, LastWriteTime | Format-Table -AutoSize | Out-String
                        }
                    } -ErrorAction Stop
                    Log-Message "[PS TOOLS] Profil utilisateur:`r`n$result" "ACTION"
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Systeme - Services
        $btnPsServices.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        Get-Service | Where-Object { $_.Status -eq 'Running' } | Select-Object -First 20 Name, DisplayName, Status | Format-Table -AutoSize | Out-String
                    } -ErrorAction Stop
                    Log-Message "[PS TOOLS] Services actifs (Top 20):`r`n$result" "ACTION"
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Systeme - Processus
        $btnPsProcesses.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        Get-Process | Sort-Object CPU -Descending | Select-Object -First 15 Name, Id, @{N = 'CPU(s)'; E = { [math]::Round($_.CPU, 2) } }, @{N = 'Mem(Mo)'; E = { [math]::Round($_.WorkingSet64 / 1MB, 2) } } | Format-Table -AutoSize | Out-String
                    } -ErrorAction Stop
                    Log-Message "[PS TOOLS] Processus (Top 15 CPU):`r`n$result" "ACTION"
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Systeme - Espace Disque
        $btnPsDiskSpace.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        Get-PSDrive -PSProvider FileSystem | Select-Object Name, @{N = 'Utilise(Go)'; E = { [math]::Round($_.Used / 1GB, 2) } }, @{N = 'Libre(Go)'; E = { [math]::Round($_.Free / 1GB, 2) } }, @{N = 'Total(Go)'; E = { [math]::Round(($_.Used + $_.Free) / 1GB, 2) } } | Format-Table -AutoSize | Out-String
                    } -ErrorAction Stop
                    Log-Message "[PS TOOLS] Espace disque:`r`n$result" "ACTION"
                    [System.Windows.MessageBox]::Show($result, "Espace Disque", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Systeme - Variables Environnement
        $btnPsEnvVars.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $env = @{
                            'COMPUTERNAME'           = $env:COMPUTERNAME
                            'USERNAME'               = $env:USERNAME
                            'USERPROFILE'            = $env:USERPROFILE
                            'TEMP'                   = $env:TEMP
                            'APPDATA'                = $env:APPDATA
                            'PROGRAMFILES'           = $env:ProgramFiles
                            'WINDIR'                 = $env:WINDIR
                            'PROCESSOR_ARCHITECTURE' = $env:PROCESSOR_ARCHITECTURE
                        }
                        ($env.GetEnumerator() | ForEach-Object { "$($_.Key) = $($_.Value)" }) -join "`r`n"
                    } -ErrorAction Stop
                    Log-Message "[PS TOOLS] Variables d'environnement:`r`n$result" "ACTION"
                    [System.Windows.MessageBox]::Show($result, "Variables Environnement", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Commande personnalisee
        $btnPsRunCustom.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                $customCmd = $txtPsCustomCmd.Text
                if ([string]::IsNullOrWhiteSpace($customCmd)) {
                    return 
                }
                
                try {
                    $scriptBlock = [scriptblock]::Create($customCmd)
                    $result = Invoke-Command -Session $script:currentSession -ScriptBlock $scriptBlock -ErrorAction Stop
                    $output = $result | Out-String
                    Log-Message "[PS TOOLS] Commande: $customCmd`r`nResultat:`r`n$output" "ACTION"
                    [System.Windows.MessageBox]::Show("Commande executee. Voir la console principale pour les resultats.", "Succes", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # === TAB CHAOS - Handlers ===
        
        # Elements CHAOS
        $btnCrazyCursor = $trollWindow.FindName("btnCrazyCursor")
        $btnCapsLockDisco = $trollWindow.FindName("btnCapsLockDisco")
        $btnRandomVolume = $trollWindow.FindName("btnRandomVolume")
        $btnFakeUpdate = $trollWindow.FindName("btnFakeUpdate")
        $btnHideTaskbar = $trollWindow.FindName("btnHideTaskbar")
        $btnShowTaskbar = $trollWindow.FindName("btnShowTaskbar")
        
        # Elements CAPTURE
        $btnTrollScreenshot = $trollWindow.FindName("btnTrollScreenshot")

        # Souris Folle - fait bouger la souris aléatoirement pendant 5s
        $btnCrazyCursor.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\crazy_cursor.ps1"
                        @'
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class CrazyCursor {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT lpPoint);
    public struct POINT { public int X; public int Y; }
}
"@
$rnd = New-Object System.Random
$endTime = (Get-Date).AddSeconds(5)
while ((Get-Date) -lt $endTime) {
    $x = $rnd.Next(100, 1800)
    $y = $rnd.Next(100, 900)
    [CrazyCursor]::SetCursorPos($x, $y)
    Start-Sleep -Milliseconds 50
}
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollCrazyCursor_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 6000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[CHAOS] Souris folle activee sur $($script:targetName)!" "ACTION"
                    [System.Windows.MessageBox]::Show("Souris folle activee pendant 5s!", "Chaos!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Caps Lock Disco - toggle caps lock rapidement
        $btnCapsLockDisco.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\capslock_disco.ps1"
                        @'
$wsh = New-Object -ComObject WScript.Shell
for ($i = 0; $i -lt 20; $i++) {
    $wsh.SendKeys("{CAPSLOCK}")
    Start-Sleep -Milliseconds 150
}
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollCapsLock_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 4000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[CHAOS] Caps Lock Disco sur $($script:targetName)!" "ACTION"
                    [System.Windows.MessageBox]::Show("Caps Lock Disco active!", "Chaos!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Volume Aleatoire
        $btnRandomVolume.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\random_volume.ps1"
                        @'
Add-Type -TypeDefinition @"
using System.Runtime.InteropServices;
[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume {
    int NotImpl1();
    int NotImpl2();
    int GetChannelCount(out int pnChannelCount);
    int SetMasterVolumeLevel(float fLevelDB, System.Guid pguidEventContext);
    int SetMasterVolumeLevelScalar(float fLevel, System.Guid pguidEventContext);
    int GetMasterVolumeLevel(out float pfLevelDB);
    int GetMasterVolumeLevelScalar(out float pfLevel);
}
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice {
    int Activate(ref System.Guid iid, int dwClsCtx, System.IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator {
    int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ppDevice);
}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }
public class Audio {
    static IAudioEndpointVolume Vol() {
        var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
        IMMDevice dev = null;
        enumerator.GetDefaultAudioEndpoint(0, 1, out dev);
        System.Guid IID = typeof(IAudioEndpointVolume).GUID;
        object o;
        dev.Activate(ref IID, 1, System.IntPtr.Zero, out o);
        return o as IAudioEndpointVolume;
    }
    public static void SetVolume(float level) { Vol().SetMasterVolumeLevelScalar(level, System.Guid.Empty); }
}
"@
$rnd = Get-Random -Minimum 10 -Maximum 100
[Audio]::SetVolume($rnd / 100)
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollVolume_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 2000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[CHAOS] Volume change aleatoirement sur $($script:targetName)!" "ACTION"
                    [System.Windows.MessageBox]::Show("Volume change aleatoirement!", "Chaos!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Fausse Mise a Jour Windows
        $btnFakeUpdate.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollFakeUpdate_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "cmd /c start https://fakeupdate.net/win10ue/" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 2000
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[CHAOS] Fausse MAJ Windows ouverte sur $($script:targetName)!" "ACTION"
                    [System.Windows.MessageBox]::Show("Fausse mise a jour Windows ouverte!`nAppuyer sur F11 pour plein ecran!", "Chaos!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Cacher Barre des Taches
        $btnHideTaskbar.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\hide_taskbar.ps1"
                        @'
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Taskbar {
    [DllImport("user32.dll")]
    public static extern int FindWindow(string className, string windowText);
    [DllImport("user32.dll")]
    public static extern int ShowWindow(int hwnd, int command);
    public static void Hide() {
        int hwnd = FindWindow("Shell_TrayWnd", "");
        ShowWindow(hwnd, 0);
    }
}
"@
[Taskbar]::Hide()
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollHideTaskbar_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 1500
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[CHAOS] Barre des taches CACHEE sur $($script:targetName)!" "ACTION"
                    [System.Windows.MessageBox]::Show("Barre des taches cachee!", "Chaos!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Afficher Barre des Taches
        $btnShowTaskbar.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    Invoke-Command -Session $script:currentSession -ScriptBlock {
                        $scriptPath = "$env:TEMP\show_taskbar.ps1"
                        @'
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class TaskbarShow {
    [DllImport("user32.dll")]
    public static extern int FindWindow(string className, string windowText);
    [DllImport("user32.dll")]
    public static extern int ShowWindow(int hwnd, int command);
    public static void Show() {
        int hwnd = FindWindow("Shell_TrayWnd", "");
        ShowWindow(hwnd, 5);
    }
}
"@
[TaskbarShow]::Show()
'@ | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
                    
                        $loggedUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName
                        if ($loggedUser) {
                            $taskName = "TrollShowTaskbar_$(Get-Random)"
                            schtasks /create /tn $taskName /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" /sc once /st 00:00 /f /ru $loggedUser /it 2>&1 | Out-Null
                            schtasks /run /tn $taskName 2>&1 | Out-Null
                            Start-Sleep -Milliseconds 1500
                            schtasks /delete /tn $taskName /f 2>&1 | Out-Null
                        }
                    } -ErrorAction Stop
                    Log-Message "[CHAOS] Barre des taches RESTAUREE sur $($script:targetName)!" "ACTION"
                    [System.Windows.MessageBox]::Show("Barre des taches restauree!", "OK!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # === TAB CAPTURE - Screenshot ===
        $btnTrollScreenshot.Add_Click({
                if (-not $script:currentSession) {
                    [System.Windows.MessageBox]::Show("Connecte-toi d'abord!", "Erreur", "OK", "Warning"); return 
                }
                try {
                    $remoteScreenPath = Invoke-Command -Session $script:currentSession -ScriptBlock {
                        Add-Type -AssemblyName System.Windows.Forms
                        Add-Type -AssemblyName System.Drawing
                        
                        $screens = [System.Windows.Forms.Screen]::AllScreens
                        $top = ($screens.Bounds.Top | Measure-Object -Minimum).Minimum
                        $left = ($screens.Bounds.Left | Measure-Object -Minimum).Minimum
                        $width = ($screens.Bounds.Right | Measure-Object -Maximum).Maximum - $left
                        $height = ($screens.Bounds.Bottom | Measure-Object -Maximum).Maximum - $top
                        
                        $bitmap = New-Object System.Drawing.Bitmap($width, $height)
                        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
                        $graphics.CopyFromScreen($left, $top, 0, 0, $bitmap.Size)
                        
                        $path = "$env:TEMP\troll_screenshot_$(Get-Date -Format 'yyyyMMdd_HHmmss').png"
                        $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
                        
                        $graphics.Dispose()
                        $bitmap.Dispose()
                        
                        return $path
                    } -ErrorAction Stop

                    $destDir = "D:\Temp\Screenshots"
                    if (-not (Test-Path $destDir)) {
                        New-Item -ItemType Directory -Path $destDir -Force | Out-Null 
                    }
                    
                    $localDest = "$destDir\TrollCapture_$($script:targetName)_$(Get-Date -Format 'yyyyMMdd_HHmmss').png"
                    Copy-Item -FromSession $script:currentSession -Path $remoteScreenPath -Destination $localDest -Force

                    Log-Message "[CAPTURE] Screenshot sauvegarde: $localDest" "SUCCESS"
                    Start-Process $localDest
                    [System.Windows.MessageBox]::Show("Capture sauvegardee et ouverte!", "Capture!", "OK", "Information")
                }
                catch {
                    [System.Windows.MessageBox]::Show("Erreur: $_", "Echec", "OK", "Error")
                }
            })

        # Bouton Fermer
        $btnCloseTroll.Add_Click({
                $trollWindow.Close()
            })

        # Credits Link
        $lblCreditsLink = $trollWindow.FindName("lblCreditsLink")
        $lblCreditsLink.Add_MouseLeftButtonDown({
                [xml]$creditsXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="CREDITS" Height="450" Width="500" 
        WindowStartupLocation="CenterScreen" Background="Black" Topmost="True">
    
    <Border BorderBrush="#00FF00" BorderThickness="2" CornerRadius="8" Background="#050505" Margin="5">
        <Grid Margin="25">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- Header -->
            <TextBlock Grid.Row="0" Text="[ CREDITS ]" FontSize="32" FontWeight="Bold" Foreground="#00FF00" HorizontalAlignment="Center" FontFamily="Consolas" Margin="0,10,0,30">
                <TextBlock.Effect>
                    <DropShadowEffect Color="#00FF00" BlurRadius="15" ShadowDepth="0" Opacity="0.7"/>
                </TextBlock.Effect>
            </TextBlock>

            <!-- Credits Content -->
            <!-- Hidden: Script co-created by Micka -->
            <StackPanel Grid.Row="1" HorizontalAlignment="Center">
                <TextBlock Text="REMOTE SESSION MANAGER" FontSize="18" Foreground="#00FF00" HorizontalAlignment="Center" FontFamily="Consolas" FontWeight="Bold"/>
                <TextBlock Text="Version 4.1 - TROLL ULTIMATE" FontSize="14" Foreground="#00AA00" HorizontalAlignment="Center" FontFamily="Consolas" Margin="0,5,0,20"/>
                
                <TextBlock Text="Developpe par:" Foreground="#888" FontFamily="Consolas" HorizontalAlignment="Center" Margin="0,10,0,5"/>
                <TextBlock Text="HOTLINE6" FontSize="24" Foreground="#00FF00" HorizontalAlignment="Center" FontFamily="Consolas" FontWeight="Bold">
                    <TextBlock.Effect>
                        <DropShadowEffect Color="#00FF00" BlurRadius="10" ShadowDepth="0" Opacity="0.5"/>
                    </TextBlock.Effect>
                </TextBlock>
                
                <TextBlock Text="---" Foreground="#333" HorizontalAlignment="Center" Margin="0,20,0,20"/>
                
                <TextBlock Text="Mode Troll Features:" Foreground="#888" FontFamily="Consolas" HorizontalAlignment="Center" Margin="0,0,0,8"/>
                <TextBlock Text="Messages | Sons | TTS | Wallpapers" Foreground="#00AA00" FontFamily="Consolas" HorizontalAlignment="Center"/>
                <TextBlock Text="Rotation | Souris | CD | Web Trolls" Foreground="#00AA00" FontFamily="Consolas" HorizontalAlignment="Center"/>
                
                <TextBlock Text="---" Foreground="#333" HorizontalAlignment="Center" Margin="0,20,0,20"/>
                
                <TextBlock Text="Support Informatique" Foreground="#888" FontFamily="Consolas" HorizontalAlignment="Center"/>
                <TextBlock Text="2024-2025" Foreground="#00AA00" FontFamily="Consolas" HorizontalAlignment="Center" Margin="0,5,0,0"/>
                <!-- ! M1ck4 w4s h3r3 ! -->
            </StackPanel>

            <!-- Bouton Fermer -->
            <Button Grid.Row="2" Name="btnCloseCredits" Content="[ OK ]" Background="#222" Foreground="#00FF00" 
                    HorizontalAlignment="Center" Padding="30,10" Margin="0,20,0,10" FontFamily="Consolas" FontWeight="Bold" Cursor="Hand"/>
        </Grid>
    </Border>
</Window>
"@
                $creditsReader = (New-Object System.Xml.XmlNodeReader $creditsXaml)
                $creditsWindow = [Windows.Markup.XamlReader]::Load($creditsReader)
                $btnCloseCredits = $creditsWindow.FindName("btnCloseCredits")
                $btnCloseCredits.Add_Click({ $creditsWindow.Close() })
                $creditsWindow.ShowDialog() | Out-Null
            })

        $trollWindow.ShowDialog() | Out-Null
    })

$window.ShowDialog() | Out-Null
Close-CurrentSession