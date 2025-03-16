Add-Type -AssemblyName System.Windows.Forms

function Get-RelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FullPath,
        [Parameter(Mandatory = $false)]
        [string]$BasePath = (Get-Location).Path
    )
    # ベースパスの末尾にディレクトリセパレータを追加して URI を作成
    $baseUri = New-Object System.Uri((Join-Path $BasePath ""))
    $fileUri = New-Object System.Uri($FullPath)
    $relativeUri = $baseUri.MakeRelativeUri($fileUri)
    $relativePath = [System.Uri]::UnescapeDataString($relativeUri.ToString())
    # URI のスラッシュを Windows のパス区切りに変換
    return $relativePath -replace '/', '\'
}

function Export-FilesToClipboard {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true, ValueFromRemainingArguments = $true)]
        [Alias("FullName")]
        [Object]$Path
    )

    begin {
        $output = New-Object System.Text.StringBuilder
        # テキストファイルの拡張子リスト
        $textExtensions = @(
            '.txt', '.md', '.csv', '.log', '.json', '.xml', '.yml', '.yaml',
            '.ini', '.conf', '.config', '.ps1', '.psm1', '.psd1', '.bat', '.cmd',
            '.sh', '.js', '.ts', '.jsx', '.tsx', '.css', '.scss', '.html', '.htm',
            '.sql', '.py', '.rb', '.java', '.c', '.cpp', '.cs', '.go', '.rs',
            '.php', '.r', '.swift', '.kt', '.kts', '.dart', '.lua', '.pl', '.pm',
            '.t', '.coffee', '.scala', '.groovy', '.vb', '.fs', '.fsx', '.erl',
            '.ex', '.exs'
        )
    }

    process {
        foreach ($item in @($Path)) {
            if ($item -is [System.IO.FileInfo] -or $item -is [System.IO.DirectoryInfo]) {
                $p = $item.FullName
            }
            else {
                $p = [string]$item
            }

            if (-not (Test-Path $p)) { continue }

            if (Test-Path $p -PathType Container) {
                # フォルダの場合は再帰的にテキストファイルを取得
                $files = Get-ChildItem -Path $p -File -Recurse
                foreach ($file in $files) {
                    if ($textExtensions -contains $file.Extension.ToLower()) {
                        try {
                            $content = Get-Content -Path $file.FullName -Raw -Encoding UTF8
                            $relPath = Get-RelativePath -FullPath $file.FullName
                            $null = $output.AppendLine("## $relPath")
                            $null = $output.AppendLine("")
                            $null = $output.AppendLine("``````")
                            $null = $output.AppendLine($content)
                            $null = $output.AppendLine("``````")
                            $null = $output.AppendLine("")
                            $null = $output.AppendLine("")
                        }
                        catch {
                            Write-Warning "Failed to process file: $($file.FullName)"
                            continue
                        }
                    }
                }
            }
            elseif (Test-Path $p -PathType Leaf) {
                # ファイルの場合
                $extension = [System.IO.Path]::GetExtension($p).ToLower()
                if ($textExtensions -contains $extension) {
                    try {
                        $content = Get-Content -Path $p -Raw -Encoding UTF8
                        $relPath = Get-RelativePath -FullPath $p
                        $null = $output.AppendLine("## $relPath")
                        $null = $output.AppendLine("")
                        $null = $output.AppendLine("``````")
                        $null = $output.AppendLine($content)
                        $null = $output.AppendLine("``````")
                        $null = $output.AppendLine("")
                        $null = $output.AppendLine("")
                    }
                    catch {
                        Write-Warning "Failed to process file: $p"
                        continue
                    }
                }
            }
        }
    }

    end {
        if ($output.Length -gt 0) {
            try {
                [System.Windows.Forms.Clipboard]::SetText($output.ToString())
                Write-Host "Content copied to clipboard successfully"
            }
            catch {
                Write-Error "Failed to copy to clipboard: $_"
            }
        }
        else {
            Write-Warning "No content was generated to copy to clipboard"
        }
    }
}

# パイプラインで引数を渡す場合
$args | Export-FilesToClipboard
