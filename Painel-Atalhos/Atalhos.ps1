param([switch]$InstallStartup, [switch]$RemoveStartup)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName @('PresentationFramework', 'PresentationCore', 'WindowsBase')
$dataDir = Join-Path $env:LOCALAPPDATA 'PainelAtalhos'
[void][IO.Directory]::CreateDirectory($dataDir)
$storeFile = Join-Path $dataDir 'atalhos.json'
$startupFile = Join-Path ([Environment]::GetFolderPath('Startup')) 'Meus Atalhos.lnk'
if ($RemoveStartup) {
    Remove-Item -LiteralPath $startupFile -Force -ErrorAction SilentlyContinue
    [void][Windows.MessageBox]::Show('Inicio automatico desativado.', 'Meus atalhos'); exit
}
if ($InstallStartup) {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($startupFile)
    $shortcut.TargetPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $shortcut.Arguments = '-NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $PSCommandPath + '"'
    $shortcut.WorkingDirectory = $PSScriptRoot; $shortcut.Save()
    [void][Windows.MessageBox]::Show('O painel abrira no proximo login. Mantenha a pasta no mesmo local.', 'Meus atalhos'); exit
}
$restoreEvent = New-Object Threading.EventWaitHandle($false, [Threading.EventResetMode]::AutoReset, 'Local\PainelAtalhosRestore')
$created = $false
$mutex = New-Object Threading.Mutex($true, 'Local\PainelAtalhos', [ref]$created)
if (-not $created) { [void]$restoreEvent.Set(); $restoreEvent.Dispose(); $mutex.Dispose(); exit }
$timer = $null
try {
    function Load-Window([string]$file) {
        [xml]$xml = Get-Content -LiteralPath (Join-Path $PSScriptRoot $file) -Raw -Encoding UTF8
        # Insert the shared theme before loading controls so StaticResource resolves.
        [xml]$theme = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Tema.xaml') -Raw -Encoding UTF8
        $resources = $xml.CreateElement('Window.Resources', $xml.DocumentElement.NamespaceURI)
        [void]$resources.AppendChild($xml.ImportNode($theme.DocumentElement, $true))
        [void]$xml.DocumentElement.InsertBefore($resources, $xml.DocumentElement.FirstChild)
        $reader = New-Object Xml.XmlNodeReader $xml
        return [Windows.Markup.XamlReader]::Load($reader)
    }
    function Get-Controls($dialog, [string[]]$names) {
        $controls = @{}
        foreach ($name in $names) { $controls[$name] = $dialog.FindName($name) }
        return $controls
    }
    $window = Load-Window 'Painel.xaml'
    $ui = Get-Controls $window @('FitBox', 'Header', 'CloseButton', 'AddButton', 'BatchButton', 'ImportButton', 'ExportButton', 'Search', 'Groups', 'Tiles', 'EmptyText', 'PrevButton', 'NextButton', 'PageText', 'SemiOption', 'FullOption', 'PinCheck', 'Status')
    $script:items = @(); $script:page = 0; $script:changingGroups = $false; $script:preferences = $null
    function Normalize-Target([string]$inputTarget) {
        $target = $inputTarget.Trim()
        if ($target.Length -eq 0) { throw 'Informe uma URL ou um caminho.' }
        if ($target.Length -gt 4096) { throw 'Destino muito longo (maximo 4096 caracteres).' }
        if ($target -match '^www\.') { $target = 'https://' + $target }
        if ($target -match '^https?://') {
            $uri = $null
            if (-not [Uri]::TryCreate($target, [UriKind]::Absolute, [ref]$uri) -or -not $uri.Host) { throw 'URL invalida.' }
            return [pscustomobject]@{ Target = $uri.AbsoluteUri; Kind = 'Site'; DefaultName = $uri.Host }
        }
        if ($target -match '^file://') {
            $fileUri = $null
            if (-not [Uri]::TryCreate($target, [UriKind]::Absolute, [ref]$fileUri) -or -not $fileUri.IsFile) { throw 'Endereco file:// invalido.' }
            $target = $fileUri.LocalPath
        }
        if ($target.StartsWith('"') -and $target.EndsWith('"') -and $target.Length -gt 1) { $target = $target.Substring(1, $target.Length - 2) }
        $target = [Environment]::ExpandEnvironmentVariables($target)
        if ($target -match '[<>"|?*\x00-\x1F]') { throw 'O caminho contem caracteres invalidos.' }
        $kind = 'Pasta'
        if ($target -match '^\\\\[^\\/:]+(?:\\.*)?$') {
            $kind = 'Rede'
            if ($target -notmatch '^\\\\[^\\]+\\[^\\]+') { $target = $target.TrimEnd('\') }
            else { $target = [IO.Path]::GetFullPath($target) }
        } elseif ($target -match '^[A-Za-z]:[\\/]') {
            $target = [IO.Path]::GetFullPath($target)
        } else { throw 'Use http://, https://, C:\pasta ou \\servidor\pasta.' }
        $leaf = ($target.TrimEnd('\') -split '\\')[-1]
        if (-not $leaf) { $leaf = $target }
        return [pscustomobject]@{ Target = $target; Kind = $kind; DefaultName = $leaf }
    }
    function New-Shortcut([string]$name, [string]$target, [string]$group, [string]$id = '') {
        $normalized = Normalize-Target $target
        if ([string]::IsNullOrWhiteSpace($name)) { $name = $normalized.DefaultName }
        if ([string]::IsNullOrWhiteSpace($group)) { $group = 'Geral' }
        $name = $name.Trim(); $group = $group.Trim()
        if ($name.Length -gt 120 -or $group.Length -gt 80) { throw 'Nome: ate 120 caracteres. Grupo: ate 80.' }
        if (-not $id) { $id = [Guid]::NewGuid().ToString() }
        return [pscustomobject]@{ Id = $id; Name = $name; Target = $normalized.Target; Group = $group; Kind = $normalized.Kind }
    }
    function Same-Target($first, $second) {
        if ($first.Kind -ne $second.Kind) { return $false }
        $comparison = [StringComparison]::OrdinalIgnoreCase
        if ($first.Kind -eq 'Site') { $comparison = [StringComparison]::Ordinal }
        return [string]::Equals($first.Target.TrimEnd('\'), $second.Target.TrimEnd('\'), $comparison)
    }
    if (Test-Path -LiteralPath $storeFile) {
        $stored = Get-Content -LiteralPath $storeFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($stored.Version -ne 1 -or $null -eq $stored.Items) { throw 'Arquivo de atalhos invalido. O arquivo original foi preservado.' }
        $ids = @{}
        foreach ($item in @($stored.Items)) {
            $parsedId = [Guid]::Empty
            if (-not [Guid]::TryParse([string]$item.Id, [ref]$parsedId) -or $ids.ContainsKey([string]$item.Id)) { throw 'Identificadores invalidos no arquivo de atalhos. O arquivo original foi preservado.' }
            $ids[[string]$item.Id] = $true
            $script:items += New-Shortcut $item.Name $item.Target $item.Group ([string]$item.Id)
        }
        $script:preferences = $stored.Preferences
    }
    function Save-Store($items) {
        $preferences = @{ Left = $window.Left; Top = $window.Top; Topmost = $window.Topmost; Opacity = ((Get-Opacity) * 100) }
        # NaN window coordinates before first rendering are stored as null.
        if ([double]::IsNaN($window.Left)) { $preferences.Left = $null }
        if ([double]::IsNaN($window.Top)) { $preferences.Top = $null }
        $json = @{ Version = 1; Items = @($items); Preferences = $preferences } | ConvertTo-Json -Depth 6
        $tempFile = Join-Path $dataDir ('atalhos-' + [Guid]::NewGuid().ToString() + '.tmp')
        try {
            [IO.File]::WriteAllText($tempFile, $json, [Text.UTF8Encoding]::new($true))
            if ([IO.File]::Exists($storeFile)) { [IO.File]::Replace($tempFile, $storeFile, ($storeFile + '.bak')) }
            else { [IO.File]::Move($tempFile, $storeFile) }
        } finally { if ([IO.File]::Exists($tempFile)) { [IO.File]::Delete($tempFile) } }
    }
    function Animate-Value($element, $property, [double]$from, [double]$to, [int]$duration) {
        if (-not [Windows.SystemParameters]::ClientAreaAnimation -or [Math]::Abs($to - $from) -lt 0.01) { return }
        $animation = New-Object Windows.Media.Animation.DoubleAnimation
        $animation.From = $from; $animation.To = $to
        $animation.Duration = [Windows.Duration]::new([TimeSpan]::FromMilliseconds($duration))
        $ease = New-Object Windows.Media.Animation.CubicEase
        $ease.EasingMode = [Windows.Media.Animation.EasingMode]::EaseOut
        $animation.EasingFunction = $ease; $animation.FillBehavior = [Windows.Media.Animation.FillBehavior]::Stop
        $element.BeginAnimation($property, $animation)
    }
    function Get-Opacity { if ($ui.SemiOption.IsChecked -eq $true) { return 0.5 }; return 1.0 }
    function Set-Opacity([bool]$animate = $true) {
        $from = $window.Opacity; $window.BeginAnimation([Windows.UIElement]::OpacityProperty, $null)
        $target = Get-Opacity; $window.Opacity = $target
        if ($animate) { Animate-Value $window ([Windows.UIElement]::OpacityProperty) $from $target 220 }
    }
    function Fit-Panel {
        $area = [Windows.SystemParameters]::WorkArea
        $ui.FitBox.MaxWidth = [Math]::Max(1, $area.Width - 24); $ui.FitBox.MaxHeight = [Math]::Max(1, $area.Height - 24)
        if ($window.IsLoaded) {
            $window.UpdateLayout()
            $window.Left = [Math]::Max($area.Left, [Math]::Min($window.Left, $area.Right - $window.ActualWidth))
            $window.Top = [Math]::Max($area.Top, [Math]::Min($window.Top, $area.Bottom - $window.ActualHeight))
        }
    }
    function Update-Groups {
        $selectedGroup = $null
        if ($null -ne $ui.Groups.SelectedItem) { $selectedGroup = $ui.Groups.SelectedItem.Value }
        $options = @([pscustomobject]@{ Label = 'Todos os grupos'; Value = $null })
        $options += @($script:items | Select-Object -ExpandProperty Group | Sort-Object -Unique | ForEach-Object { [pscustomobject]@{ Label = $_; Value = $_ } })
        $script:changingGroups = $true
        try {
            $ui.Groups.DisplayMemberPath = 'Label'; $ui.Groups.ItemsSource = $options
            $ui.Groups.SelectedIndex = 0
            for ($i = 1; $i -lt $options.Count; $i++) { if ($options[$i].Value -eq $selectedGroup) { $ui.Groups.SelectedIndex = $i; break } }
        } finally { $script:changingGroups = $false }
    }
    function Find-Shortcut([string]$id) { return ($script:items | Where-Object Id -eq $id | Select-Object -First 1) }
    function Open-Shortcut([string]$id) {
        $item = Find-Shortcut $id
        if ($null -eq $item) { return }
        try {
            # Use the Windows shell association directly; never interpolate into a command line.
            $start = New-Object Diagnostics.ProcessStartInfo
            $start.FileName = $item.Target; $start.UseShellExecute = $true
            [void][Diagnostics.Process]::Start($start)
            $ui.Status.Text = 'Abrindo: ' + $item.Name
        } catch { $ui.Status.Text = 'Nao foi possivel abrir ' + $item.Name + ': ' + $_.Exception.Message }
    }
    function Render-Tiles {
        $query = $ui.Search.Text.Trim(); $group = $ui.Groups.SelectedItem.Value
        $filtered = @($script:items | Where-Object {
            $content = $_.Name + ' ' + $_.Group + ' ' + $_.Target
            (-not $group -or $_.Group -eq $group) -and (-not $query -or $content.IndexOf($query, [StringComparison]::OrdinalIgnoreCase) -ge 0)
        } | Sort-Object -Property @('Group', 'Name'))
        $pageCount = [Math]::Max(1, [Math]::Ceiling($filtered.Count / 8.0))
        $script:page = [int][Math]::Max(0, [Math]::Min($script:page, $pageCount - 1))
        $ui.Tiles.Children.Clear()
        foreach ($item in @($filtered | Select-Object -Skip ($script:page * 8) -First 8)) {
            $button = New-Object Windows.Controls.Button
            $button.Style = $window.Resources['ShortcutStyle']; $button.Tag = $item.Id
            $button.ToolTip = $item.Name + "`n" + $item.Target
            $button.RenderTransform = New-Object Windows.Media.ScaleTransform
            $content = New-Object Windows.Controls.Grid
            $iconColumn = New-Object Windows.Controls.ColumnDefinition
            $iconColumn.Width = [Windows.GridLength]::new(42)
            [void]$content.ColumnDefinitions.Add($iconColumn)
            [void]$content.ColumnDefinitions.Add((New-Object Windows.Controls.ColumnDefinition))
            $iconBorder = New-Object Windows.Controls.Border
            $iconBorder.Width = 32; $iconBorder.Height = 32; $iconBorder.CornerRadius = '9'
            $iconBorder.HorizontalAlignment = [Windows.HorizontalAlignment]::Left
            $iconBorder.Background = [Windows.Media.BrushConverter]::new().ConvertFromString('#182535')
            $accent = '#7DB8FF'
            $geometry = 'M 4,18 L 18,4 M 8,4 L 18,4 L 18,14 M 4,4 L 4,4 M 4,10 L 4,20 L 14,20'
            if ($item.Kind -eq 'Pasta') {
                $accent = '#F4BF66'
                $geometry = 'M 2,7 L 2,19 L 22,19 L 22,7 Z M 2,7 L 2,4 L 9,4 L 12,7'
            } elseif ($item.Kind -eq 'Rede') {
                $accent = '#5EDDC5'
                $geometry = 'M 9,2 L 15,2 L 15,8 L 9,8 Z M 12,8 L 12,13 M 4,13 L 20,13 M 4,13 L 4,16 M 20,13 L 20,16 M 1,16 L 7,16 L 7,22 L 1,22 Z M 17,16 L 23,16 L 23,22 L 17,22 Z'
            }
            $icon = New-Object Windows.Shapes.Path
            $icon.Data = [Windows.Media.Geometry]::Parse($geometry)
            $icon.Stroke = [Windows.Media.BrushConverter]::new().ConvertFromString($accent)
            $icon.StrokeThickness = 1.5; $icon.Width = 23; $icon.Height = 23
            $icon.Stretch = [Windows.Media.Stretch]::Uniform
            $icon.HorizontalAlignment = [Windows.HorizontalAlignment]::Center
            $icon.VerticalAlignment = [Windows.VerticalAlignment]::Center
            $iconBorder.Child = $icon; [void]$content.Children.Add($iconBorder)
            $stack = New-Object Windows.Controls.StackPanel
            [Windows.Controls.Grid]::SetColumn($stack, 1)
            $label = New-Object Windows.Controls.TextBlock
            $label.Text = $item.Name; $label.FontSize = 14; $label.FontWeight = [Windows.FontWeights]::SemiBold
            $label.TextTrimming = [Windows.TextTrimming]::CharacterEllipsis; $label.MaxWidth = 139
            $detail = New-Object Windows.Controls.TextBlock
            $detail.Text = $item.Kind.ToUpper() + ' · ' + $item.Group; $detail.FontSize = 10; $detail.Margin = '0,4,0,0'
            $detail.Foreground = $icon.Stroke
            $detail.TextTrimming = [Windows.TextTrimming]::CharacterEllipsis; $detail.MaxWidth = 139
            $targetLabel = New-Object Windows.Controls.TextBlock
            $targetLabel.Text = $item.Target; $targetLabel.FontSize = 10; $targetLabel.Margin = '0,5,0,0'
            $targetLabel.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'Muted')
            $targetLabel.TextTrimming = [Windows.TextTrimming]::CharacterEllipsis; $targetLabel.MaxWidth = 139
            [void]$stack.Children.Add($label); [void]$stack.Children.Add($detail); [void]$stack.Children.Add($targetLabel)
            [void]$content.Children.Add($stack); $button.Content = $content
            $button.Add_Click({ param($sender, $eventArgs); Open-Shortcut ([string]$sender.Tag) })
            $button.Add_MouseEnter({
                param($sender, $eventArgs)
                Animate-Value $sender.RenderTransform ([Windows.Media.ScaleTransform]::ScaleXProperty) $sender.RenderTransform.ScaleX 1.02 120
                Animate-Value $sender.RenderTransform ([Windows.Media.ScaleTransform]::ScaleYProperty) $sender.RenderTransform.ScaleY 1.02 120
                $sender.RenderTransform.ScaleX = 1.02; $sender.RenderTransform.ScaleY = 1.02
            })
            $button.Add_MouseLeave({
                param($sender, $eventArgs)
                Animate-Value $sender.RenderTransform ([Windows.Media.ScaleTransform]::ScaleXProperty) $sender.RenderTransform.ScaleX 1 120
                Animate-Value $sender.RenderTransform ([Windows.Media.ScaleTransform]::ScaleYProperty) $sender.RenderTransform.ScaleY 1 120
                $sender.RenderTransform.ScaleX = 1; $sender.RenderTransform.ScaleY = 1
            })
            $menu = New-Object Windows.Controls.ContextMenu
            $menu.Background = $window.Resources['Card']; $menu.Foreground = $window.Resources['Text']
            foreach ($action in @('Abrir', 'Editar', 'Copiar destino', 'Remover')) {
                $entry = New-Object Windows.Controls.MenuItem
                $entry.Header = $action; $entry.Tag = $item.Id
                $entry.Background = $window.Resources['Card']; $entry.Foreground = $window.Resources['Text']
                $entry.Add_Click({
                    param($sender, $eventArgs)
                    $selectedItem = Find-Shortcut ([string]$sender.Tag)
                    if ($null -eq $selectedItem) { return }
                    switch ([string]$sender.Header) {
                        'Abrir' { Open-Shortcut $selectedItem.Id }
                        'Editar' { Show-Editor $selectedItem }
                        'Copiar destino' { try { [Windows.Clipboard]::SetText($selectedItem.Target); $ui.Status.Text = 'Destino copiado.' } catch { $ui.Status.Text = 'Nao foi possivel copiar.' } }
                        'Remover' {
                            if ([Windows.MessageBox]::Show($window, ('Remover o atalho "' + $selectedItem.Name + '"?'), 'Meus atalhos', [Windows.MessageBoxButton]::YesNo) -eq [Windows.MessageBoxResult]::Yes) {
                                try { $candidate = @($script:items | Where-Object Id -ne $selectedItem.Id); Save-Store $candidate; $script:items = $candidate; Update-Groups; Render-Tiles; $ui.Status.Text = 'Atalho removido.' }
                                catch { $ui.Status.Text = 'Nao foi possivel salvar: ' + $_.Exception.Message }
                            }
                        }
                    }
                })
                [void]$menu.Items.Add($entry)
            }
            $button.ContextMenu = $menu; [void]$ui.Tiles.Children.Add($button)
        }
        $ui.EmptyText.Visibility = [Windows.Visibility]::Collapsed
        if ($filtered.Count -eq 0) {
            $ui.EmptyText.Visibility = [Windows.Visibility]::Visible
            $ui.EmptyText.Text = 'Nenhum atalho corresponde a busca.'
            if ($script:items.Count -eq 0) { $ui.EmptyText.Text = 'Adicione seu primeiro atalho para comecar.' }
        }
        $ui.PageText.Text = '{0} atalhos | pagina {1}/{2}' -f $filtered.Count, ($script:page + 1), $pageCount
        $ui.PrevButton.IsEnabled = ($script:page -gt 0); $ui.NextButton.IsEnabled = ($script:page -lt $pageCount - 1)
        Fit-Panel
    }
    function Reveal-Shortcut([string]$id) {
        $ui.Search.Text = ''; $ui.Groups.SelectedIndex = 0
        $ordered = @($script:items | Sort-Object -Property @('Group', 'Name'))
        for ($index = 0; $index -lt $ordered.Count; $index++) {
            if ($ordered[$index].Id -eq $id) { $script:page = [int][Math]::Floor($index / 8.0); break }
        }
        Render-Tiles
    }
    function Commit-Additions($newItems) {
        $candidate = @($script:items); $added = 0; $duplicates = 0; $revealId = $null
        foreach ($newItem in $newItems) {
            $exists = @($candidate | Where-Object { Same-Target $_ $newItem })
            if ($exists.Count -gt 0) { $duplicates++; continue }
            $candidate += $newItem; $added++; if (-not $revealId) { $revealId = $newItem.Id }
        }
        if ($added -gt 0) { Save-Store $candidate; $script:items = $candidate; Update-Groups; Reveal-Shortcut $revealId }
        $ui.Status.Text = '{0} adicionados; {1} destinos duplicados ignorados.' -f $added, $duplicates
    }
    function Show-Editor($item = $null) {
        $script:editDialog = Load-Window 'Editar.xaml'; $script:editDialog.Owner = $window
        $script:editUi = Get-Controls $script:editDialog @('NameInput', 'TargetInput', 'GroupInput', 'ErrorText', 'SaveButton', 'CancelButton')
        $script:editingItem = $item
        if ($null -ne $item) {
            $script:editDialog.Title = 'Editar atalho'; $script:editUi.NameInput.Text = $item.Name
            $script:editUi.TargetInput.Text = $item.Target; $script:editUi.GroupInput.Text = $item.Group
        }
        $script:editUi.SaveButton.Add_Click({
            try {
                $id = ''; if ($null -ne $script:editingItem) { $id = $script:editingItem.Id }
                $newItem = New-Shortcut $script:editUi.NameInput.Text $script:editUi.TargetInput.Text $script:editUi.GroupInput.Text $id
                if ($null -eq $script:editingItem) { Commit-Additions @($newItem) }
                else {
                    $duplicate = @($script:items | Where-Object { $_.Id -ne $id -and (Same-Target $_ $newItem) })
                    if ($duplicate.Count) { throw 'Ja existe um atalho para esse destino.' }
                    $candidate = @($script:items | ForEach-Object { if ($_.Id -eq $id) { $newItem } else { $_ } })
                    Save-Store $candidate; $script:items = $candidate; Update-Groups; Reveal-Shortcut $newItem.Id; $ui.Status.Text = 'Atalho atualizado.'
                }
                $script:editDialog.DialogResult = $true
            } catch { $script:editUi.ErrorText.Text = $_.Exception.Message }
        })
        [void]$script:editDialog.ShowDialog()
    }
    function Show-Batch {
        $script:batchDialog = Load-Window 'Adicionar-Varias.xaml'; $script:batchDialog.Owner = $window
        $script:batchUi = Get-Controls $script:batchDialog @('BatchInput', 'GroupInput', 'ErrorText', 'SaveButton', 'CancelButton')
        $script:batchUi.SaveButton.Add_Click({
            try {
                $newItems = @(); $lineNumber = 0
                foreach ($line in ($script:batchUi.BatchInput.Text -split '\r?\n')) {
                    $lineNumber++; if ([string]::IsNullOrWhiteSpace($line)) { continue }
                    $parts = $line.Trim() -split '\|', 2; $name = ''; $target = $parts[0]
                    if ($parts.Count -eq 2) { $name = $parts[0]; $target = $parts[1] }
                    try { $newItems += New-Shortcut $name $target $script:batchUi.GroupInput.Text }
                    catch { throw ('Linha ' + $lineNumber + ': ' + $_.Exception.Message) }
                }
                if ($newItems.Count -eq 0) { throw 'Cole pelo menos um destino.' }
                Commit-Additions $newItems; $script:batchDialog.DialogResult = $true
            } catch { $script:batchUi.ErrorText.Text = $_.Exception.Message }
        })
        [void]$script:batchDialog.ShowDialog()
    }
    $ui.AddButton.Add_Click({ Show-Editor })
    $ui.BatchButton.Add_Click({ Show-Batch })
    $ui.Search.Add_TextChanged({ $script:page = 0; Render-Tiles })
    $ui.Groups.Add_SelectionChanged({ if (-not $script:changingGroups) { $script:page = 0; Render-Tiles } })
    $ui.PrevButton.Add_Click({ $script:page--; Render-Tiles })
    $ui.NextButton.Add_Click({ $script:page++; Render-Tiles })
    $ui.SemiOption.Add_Checked({ Set-Opacity })
    $ui.FullOption.Add_Checked({ Set-Opacity })
    $ui.PinCheck.Add_Checked({ $window.Topmost = $true })
    $ui.PinCheck.Add_Unchecked({ $window.Topmost = $false })
    $ui.CloseButton.Add_Click({ $window.Close() })
    $ui.Header.Add_MouseLeftButtonDown({
        param($sender, $eventArgs)
        $source = $eventArgs.OriginalSource
        while ($null -ne $source) {
            if ($source -is [Windows.Controls.Button]) { return }
            try { $source = [Windows.Media.VisualTreeHelper]::GetParent($source) } catch { break }
        }
        $window.DragMove()
    })
    $ui.ExportButton.Add_Click({
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = 'Lista de atalhos (*.json)|*.json'; $dialog.FileName = 'Meus-Atalhos.json'
        if ($dialog.ShowDialog() -eq $true) {
            try {
                $json = @{ Version = 1; Items = @($script:items) } | ConvertTo-Json -Depth 6
                [IO.File]::WriteAllText($dialog.FileName, $json, [Text.UTF8Encoding]::new($true))
                $ui.Status.Text = 'Lista exportada.'
            } catch { $ui.Status.Text = 'Erro ao exportar: ' + $_.Exception.Message }
        }
    })
    $ui.ImportButton.Add_Click({
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = 'Lista de atalhos (*.json)|*.json'
        if ($dialog.ShowDialog() -eq $true) {
            try {
                $imported = Get-Content -LiteralPath $dialog.FileName -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($imported.Version -ne 1 -or $null -eq $imported.Items) { throw 'Use uma lista exportada por este painel.' }
                $newItems = @($imported.Items | ForEach-Object { New-Shortcut $_.Name $_.Target $_.Group })
                Commit-Additions $newItems
            } catch { $ui.Status.Text = 'Erro ao importar: ' + $_.Exception.Message }
        }
    })
    $window.Add_ContentRendered({
        Fit-Panel; $area = [Windows.SystemParameters]::WorkArea
        $window.Left = $area.Right - $window.ActualWidth - 12; $window.Top = $area.Top + 12
        if ($null -ne $script:preferences) {
            if ($null -ne $script:preferences.Left) { $window.Left = [double]$script:preferences.Left }
            if ($null -ne $script:preferences.Top) { $window.Top = [double]$script:preferences.Top }
            $ui.PinCheck.IsChecked = [bool]$script:preferences.Topmost
            if ($script:preferences.Opacity -eq 50) { $ui.SemiOption.IsChecked = $true }
        }
        Fit-Panel; Set-Opacity -animate $false
        Animate-Value $window ([Windows.UIElement]::OpacityProperty) 0 (Get-Opacity) 350
    })
    $window.Add_Closing({
        try { Save-Store $script:items }
        catch { [void][Windows.MessageBox]::Show($window, ('Nao foi possivel salvar preferencias: ' + $_.Exception.Message), 'Meus atalhos') }
    })
    $timer = New-Object Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(250)
    $timer.Add_Tick({ if ($restoreEvent.WaitOne(0)) { $ui.FullOption.IsChecked = $true; Set-Opacity; [void]$window.Activate() } })
    Update-Groups; Render-Tiles; Fit-Panel; $timer.Start()
    [void]$window.ShowDialog()
} catch {
    $_ | Out-String | Set-Content -LiteralPath (Join-Path $dataDir 'erro.txt') -Encoding UTF8
    [void][Windows.MessageBox]::Show('Nao foi possivel abrir o painel. Detalhes em ' + $dataDir + '\erro.txt. Se houver atalhos.json, ele foi preservado.', 'Meus atalhos')
} finally {
    if ($null -ne $timer) { $timer.Stop() }
    $restoreEvent.Dispose(); $mutex.ReleaseMutex(); $mutex.Dispose()
}
