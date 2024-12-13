Set-StrictMode -Version Latest

[string]$DirectoryPath = "C:\Users\mikehow\OneDrive\Desktop\SamsungPics"
$PowerPoint = $null
$HttpListener = $null
$CurrentPresentation = $null

function Get-PowerPointInstance {
    try {
        $PowerPoint = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
        Write-Host "Connected to existing PowerPoint instance."
    } catch {
        Write-Host "No active PowerPoint instance found. Starting a new one..."
        $PowerPoint = New-Object -ComObject PowerPoint.Application
        $PowerPoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    }
    return $PowerPoint
}

function BringToFront {
    param (
        [System.__ComObject]$PowerPoint
    )
    $PowerPoint.Activate()
}

function Generate-HTML {
    param (
        [string]$DirectoryPath,
        [string]$Message
    )
    $Files = Get-ChildItem -Path $DirectoryPath -Filter "*.pptx" | Select-Object -ExpandProperty Name
    $HTML = @"
<!DOCTYPE html>
<html>
<head>
    <title>PowerPoint File Selector</title>
</head>
<body>
    <h1>Select a PowerPoint File</h1>
    <ul>
"@
    foreach ($File in $Files) {
        $HTML += "<li><a href='/$File'>$File</a></li>`n"
    }
    $HTML += @"
    </ul>
    <h1>Navigation</h1>
    <ul>
        <li><a href='/forward'>Forward Slide</a></li>
        <li><a href='/back'>Back Slide</a></li>
        <li><a href='/'>Home</a></li>
    </ul>
"@
    if ($Message) {
        $HTML += "<p>$Message</p>"
    }
    $HTML += @"
</body>
</html>
"@
    return $HTML
}

function Handle-Request {
    param (
        [System.Net.HttpListenerContext]$Context
    )
    $RequestedFile = $Context.Request.Url.AbsolutePath.TrimStart("/")
    $Response = $Context.Response

    if ($RequestedFile -eq "") {
        $HTML = Generate-HTML -DirectoryPath $DirectoryPath -Message ""
    } elseif ($RequestedFile -eq "forward") {
        if ($CurrentPresentation -ne $null) {
            $CurrentPresentation.SlideShowWindow.View.Next()
            $HTML = Generate-HTML -DirectoryPath $DirectoryPath -Message "Moved to next slide."
        } else {
            $HTML = Generate-HTML -DirectoryPath $DirectoryPath -Message "No presentation loaded."
        }
    } elseif ($RequestedFile -eq "back") {
        if ($CurrentPresentation -ne $null) {
            $CurrentPresentation.SlideShowWindow.View.Previous()
            $HTML = Generate-HTML -DirectoryPath $DirectoryPath -Message "Moved to previous slide."
        } else {
            $HTML = Generate-HTML -DirectoryPath $DirectoryPath -Message "No presentation loaded."
        }
    } elseif ($RequestedFile -match "\.pptx$") {
        $FullFilePath = Join-Path -Path $DirectoryPath -ChildPath $RequestedFile
        if (Test-Path $FullFilePath) {
            if ($CurrentPresentation -ne $null) {
                $CurrentPresentation.Close()
            }
            $CurrentPresentation = $PowerPoint.Presentations.Open($FullFilePath)
            $SlideShow = $CurrentPresentation.SlideShowSettings
            $SlideShow.StartingSlide = 1
            $SlideShow.EndingSlide = $CurrentPresentation.Slides.Count
            $SlideShow.Run()
            BringToFront -PowerPoint $PowerPoint
            $HTML = Generate-HTML -DirectoryPath $DirectoryPath -Message "Loaded presentation: $RequestedFile"
        } else {
            $HTML = Generate-HTML -DirectoryPath $DirectoryPath -Message "File not found: $RequestedFile"
        }
    } else {
        $HTML = Generate-HTML -DirectoryPath $DirectoryPath -Message "Invalid request: $RequestedFile"
    }

    $Response.ContentType = "text/html"
    $Response.StatusCode = 200
    $Response.OutputStream.Write(([System.Text.Encoding]::UTF8.GetBytes($HTML)), 0, ([System.Text.Encoding]::UTF8.GetByteCount($HTML)))
    $Response.OutputStream.Close()
}

try {
    $PowerPoint = Get-PowerPointInstance

    $Files = @(Get-ChildItem -Path $DirectoryPath -Filter "*.pptx")
    if ($Files.Count -gt 0) {
        $FullFilePath = Join-Path -Path $DirectoryPath -ChildPath $Files[0].Name
        $CurrentPresentation = $PowerPoint.Presentations.Open($FullFilePath)
        $SlideShow = $CurrentPresentation.SlideShowSettings
        $SlideShow.StartingSlide = 1
        $SlideShow.EndingSlide = $CurrentPresentation.Slides.Count
        $SlideShow.Run()
        BringToFront -PowerPoint $PowerPoint
        Write-Host "Loaded presentation: $FullFilePath"
    } else {
        Write-Host "No presentations found in directory."
    }

    $HttpListener = New-Object System.Net.HttpListener
    $HttpListener.Prefixes.Add("http://localhost:6969/")
    $HttpListener.Start()
    Write-Host "Web server started at http://localhost:6969"

    while ($HttpListener.IsListening) {
        $Context = $HttpListener.GetContext()
        Handle-Request -Context $Context
    }
} catch {
    Write-Error "An error occurred: $_"
} finally {
    if ($HttpListener) {
        $HttpListener.Stop()
        Write-Host "Web server stopped."
    }
    if ($PowerPoint) {
        $PowerPoint.Quit()
        Write-Host "PowerPoint closed."
    }
}
