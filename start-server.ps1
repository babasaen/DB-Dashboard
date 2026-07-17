$port = 8080
$root = $PSScriptRoot
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()
Write-Host "Dashboard: http://localhost:$port/" -ForegroundColor Green
Write-Host "Press Ctrl+C to stop"
while ($listener.IsListening) {
    $ctx = $listener.GetContext()
    $path = $ctx.Request.Url.LocalPath
    if ($path -eq '/') { $path = '/index.html' }
    $file = Join-Path $root $path.TrimStart('/')
    if (Test-Path $file) {
        $bytes = [IO.File]::ReadAllBytes($file)
        $ext = [IO.Path]::GetExtension($file)
        $mimeMap = @{ '.html'='text/html'; '.js'='application/javascript'; '.json'='application/json'; '.xlsx'='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
        $mime = $mimeMap[$ext]
        if (-not $mime) { $mime = 'application/octet-stream' }
        $ctx.Response.ContentType = "$mime; charset=utf-8"
        $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } else {
        $ctx.Response.StatusCode = 404
    }
    $ctx.Response.Close()
}
