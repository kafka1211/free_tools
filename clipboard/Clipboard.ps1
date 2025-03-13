Add-Type -AssemblyName System.Windows.Forms

function Export-FilesToClipboard {
   begin {
       $output = New-Object System.Text.StringBuilder
       # テキストファイルの拡張子リスト
       $textExtensions = @('.txt', '.md', '.csv', '.log', '.json', '.xml', '.yml', '.yaml', '.ini', '.conf', '.config', '.ps1', '.psm1', '.psd1', '.bat', '.cmd', '.sh', '.js', '.ts', '.jsx', '.tsx', '.css', '.scss', '.html', '.htm', '.sql', '.py', '.rb', '.java', '.c', '.cpp', '.cs', '.go', '.rs', '.php', '.r', '.swift', '.kt', '.kts', '.dart', '.lua', '.pl', '.pm', '.t', '.coffee', '.scala', '.groovy', '.vb', '.fs', '.fsx', '.erl', '.ex', '.exs')
   }
   
   process {
       $paths = @($input)
       if (-not $paths) {
           $paths = $Path
       }
       
       foreach ($p in $paths) {
           if (-not $p) { continue }
           
           if (Test-Path $p -PathType Container) {
               $files = Get-ChildItem -Path $p -File -Recurse
               foreach ($file in $files) {
                   # ファイル拡張子がテキストファイルリストにある場合のみ処理
                   if ($textExtensions -contains $file.Extension.ToLower()) {
                       try {
                           $content = Get-Content -Path $file.FullName -Raw -Encoding UTF8
                           [void]$output.AppendLine("## $($file.Name)")
                           [void]$output.AppendLine("")
                           [void]$output.AppendLine("``````")
                           [void]$output.AppendLine($content)
                           [void]$output.AppendLine("``````")
                           [void]$output.AppendLine("")
                           [void]$output.AppendLine("")
                       }
                       catch {
                           Write-Warning "Failed to process file: $($file.FullName)"
                           continue
                       }
                   }
               }
           }
           elseif (Test-Path $p -PathType Leaf) {
               # ファイル拡張子がテキストファイルリストにある場合のみ処理
               $extension = [System.IO.Path]::GetExtension($p).ToLower()
               if ($textExtensions -contains $extension) {
                   try {
                       $content = Get-Content -Path $p -Raw -Encoding UTF8
                       $filename = Split-Path $p -Leaf
                       [void]$output.AppendLine("## $filename")
                       [void]$output.AppendLine("")
                       [void]$output.AppendLine("``````")
                       [void]$output.AppendLine($content)
                       [void]$output.AppendLine("``````")
                       [void]$output.AppendLine("")
                       [void]$output.AppendLine("")
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

# パイプラインで引数を渡す
$args | Export-FilesToClipboard
