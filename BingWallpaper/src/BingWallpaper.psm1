# https://github.com/Jaykul/MultiMonitorHelper
# Requires -Assembly MultiMonitorHelper.dll

Add-Type -Name Windows -Namespace System -MemberDefinition '
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo (int uAction, int uParam, string lpvParam, int fuWinIni);
'
$urlbase = "http://www.bing.com/"

# A few resolutions of these wallpapers exist, as far as I know ...
$KnownAvailable = [Ordered]@{
    0.56 = @('1080x1920','768x1366')
    0.6  = @('240x400','480x800','768x1280')
    0.75 = @('240x320','360x480','480x640','768x1024')
    0.8  = @('176x220')
    1.0  = @('240x240','320x320')
    1.25 = @('220x176')
    1.33 = @('1024x768','320x240','480x360','640x480','800x600')
    1.6  = @('1920x1200')
    1.67 = @('1280x768','400x240','800x480')
    1.78 = @('1280x720','1366x768','1920x1080')
}
$Sizes = @($KnownAvailable.Values | % { $_ })
$Ratios = @($knownAvailable.Keys)

function Get-Size {
    param($Size)
    if($Sizes -Contains $Size) {
        return $Size
    }

    [int]$Width, [int]$Height = $Size -split 'x'
    $Ratio = [Math]::Round(($Width / $Height), 2)
    Write-Debug ("{0}x{1} = {2}" -f $Width, $Height, $Ratio)

    $r = [array]::BinarySearch( $Ratios, $Ratio )
    Write-Debug "Index of $Ratio = $r in $Ratios"
    if($r -ge 0) {
       $Key = $Ratios[$r]
    } else {
        $r = [Math]::Abs($r+2)
        if(($Ratios[$r] - $Ratio) -gt ($Ratio - $Ratios[($r-1)])) {
           $Key = $Ratios[($r-1)]
        } else {
           $Key = $Ratios[$r]
        }
    }

    foreach($Sz in $KnownAvailable[$Key]) {
        $W = [int]($Sz -split 'x')[0]
        Write-Debug ("Can't match {0}x{1}, try {2}" -f $Width, $Height, $W)
        if($W -ge $Width) {
            return $Sz
        }
    }
    return $KnownAvailable[$Key][-1]
}

function Get-ActiveDisplays {
    #.Synopsis
    #  Get the currently available displays

    #.Notes
    # Windows.Forms.Screen alters the size of low-DPI monitors in mixed-DPI systems
    # System.Windows.SystemParameters alters the size of hight-DPI monitors in mixed-DPI systems

    # # For example, the following code proved buggy on my systems:
    # [System.Windows.Forms.Screen]::AllScreens | Select DeviceName -Expand Bounds

    # # Given a 3200x1800 high DPI laptop screen and two 1080p screens (one rotated):
    # # It produces this information:

    # Name             X    Y Width Height
    # ----             -    - ----- ------
    # \\.\DISPLAY1     0    0  3200   1800
    # \\.\DISPLAY2 -3840 1448  3840   2160
    # \\.\DISPLAY3 -6000  330  2160   3840

    # # The code below, using https://github.com/ChrisEelmaa/MultiMonitorHelper
    # # Produces the following correct information on my system:
    # Name             X   Y Height Width  Rotation
    # ----             -   - ------ -----  --------
    # \\.\DISPLAY1     0   0   1800  3200   Default
    # \\.\DISPLAY2 -1920 724   1080  1920   Default
    # \\.\DISPLAY3 -3000 165   1920  1080 Rotated90

    if(!("MultiMonitorHelper.DisplayFactory" -As [Type])) {
        Add-Type -AssemblyName System.Windows.Forms
        # This should work on "simple" systems without mixed DPI displays
        @([System.Windows.Forms.Screen]::AllScreens | Select DeviceName -Expand Bounds)
    } else {
        @(@([MultiMonitorHelper.DisplayFactory]::GetDisplayModel().GetActiveDisplays()) |
                Select Name,
                       @{Name = "X"; Expr = { $_.Origin.X }},
                       @{Name = "Y"; Expr = { $_.Origin.Y }},
                       @{Name = "Height"; Expr = { if($_.Rotation -in "Default", "Rotated180") { $_.Resolution.Height } else { $_.Resolution.Width } }},
                       @{Name = "Width"; Expr = { if($_.Rotation -in "Default", "Rotated180") { $_.Resolution.Width } else { $_.Resolution.Height } }},
                       Rotatione, IsPrimary, IsActive)
    }
}


function Write-OutlineText {
    <#
        .SYNOPSIS
            Writes text to an image.
        .DESCRIPTION
            Writes one or more lines of text onto a rectangle on an image.

            By default, it writes the computer name on the bottom right corner, with an outline around so you can read it regardless of the picture.
        .EXAMPLE
            Write-OutlineText "$Env:UserName`n$Env:ComputerName`n$Env:UserDnsDomain" -Path "~\Wallpaper.png"
    #>
    [CmdletBinding()]
    param(
        [Parameter(ParameterSetName="OnImageFile", Mandatory)]
        [ValidateScript({
            if (!(Test-Path "$_")) {
                throw "The path must point to an existing image. Can't find '$_'"
            }
            $true
        })]
        [string]$Path,

        [Parameter(ParameterSetName="OnExistingGraphics", ValueFromPipeline, Mandatory)]
        [System.Drawing.Graphics]$Graphics,

        # By default, the bounds are 20 pixels in from the size of the graphic, with an extra 40px off the bottom.
        [System.Drawing.RectangleF]$Bounds,

        [string]$Text = $Env:ComputerName,

        [string]$FontName = "Cascadia Code",

        [int]$FontSize = 18,

        [System.Drawing.FontStyle]$FontStyle = "Bold",

        # The alignment relative to the bounds. Note that in a left-to-right layout, the far position is bottom, and the near position is top.
        [System.Drawing.StringAlignment]$VerticalAlignment = "Far",

        # The alignment relative to the bounds. Note that in a left-to-right layout, the far position is right, and the near position is left. However, in a right-to-left layout, the far position is left.
        [System.Drawing.StringAlignment]$HorizontalAlignment = "Far",

        # The stroke color (defaults to [System.Drawing.Brushes]::Black)
        [System.Drawing.Brush]$StrokeBrush = [System.Drawing.Brushes]::Black,

        # The fill color (defaults to [System.Drawing.Brushes]::White)
        [System.Drawing.Brush]$FillBrush = [System.Drawing.Brushes]::White
    )

    if ($PSCmdlet.ParameterSetName -eq "OnImageFile") {
        $Source = [System.Drawing.Image]::FromFile((Convert-Path $Path))
        $Graphics = [System.Drawing.Graphics]::FromImage($Source)
    }
    try {
        if (!$Bounds) {
            $Bounds = [System.Drawing.RectangleF]::new($Graphics.VisibleClipBounds.Location + [System.Drawing.SizeF]::new(20, 20), $Graphics.VisibleClipBounds.Size - [System.Drawing.SizeF]::new(40, 80))
            Write-Verbose "Using Bounds: $($Bounds.Top), $($Bounds.Left), $($Bounds.Bottom, $Bounds.Right)"
        }
        $Font = try {
            [System.Drawing.FontFamily]::new($FontName)
        } catch {
            [System.Drawing.FontFamily]::GenericMonospace
        }
        Write-Verbose "Using FontFamily $Font"

        $Format = [System.Drawing.StringFormat]::GenericTypographic
        $Format.Alignment = $HorizontalAlignment
        $Format.LineAlignment = $VerticalAlignment

        $GraphicsPath = [System.Drawing.Drawing2D.GraphicsPath]::new()
        $GraphicsPath.AddString(
            $Text,
            $Font,
            $FontStyle,
            ($Graphics.DpiY * $FontSize / 72),
            $Bounds,
            $Format);

        Write-Verbose "Adding '$Text' to the image in $FillBrush on $StrokeBrush"
        $Graphics.DrawPath($StrokeBrush, $GraphicsPath);
        $Graphics.FillPath($FillBrush, $GraphicsPath);
    } catch {
        Write-Warning "Unhandled Error: $_"
        Write-Warning "Unhandled Error: $($_.ScriptStackTrace)"
        throw
    } finally {
        if ($PSCmdlet.ParameterSetName -eq "OnImageFile") {
            $Graphics.Dispose()
            try {
                # Save as png to avoid asking them, and dealing with image format as a parameter
                $NewPath = "{0}\{1}-1.png" -f [IO.Path]::GetDirectoryName($Path), [IO.Path]::GetFileNameWithoutExtension($Path)
                $Source.Save($NewPath, [System.Drawing.Imaging.ImageFormat]::Png)
                Get-Item $NewPath
            } finally {
                $Source.Dispose()
            }
        }
    }
}

function Set-BingWallpaper {
    #.Synopsis
    #  Fetches Bing Homepage images and generates a WallPaper
    #.Description
    #  With support for mixed-DPI multi-display configurations, this command will download one or more Bing homepage images from the last several days and generate a custom wallpaper for all your connected screens.
    #.Example
    #  Set-BingWallpaper
    #  Sets the wallpaper on each of your displays to a different recent Bing homepage image
    #.Example
    #  Set-BingWallpaper -Offset 1
    #  Sets the wallpaper on each of your displays using the image(s) from yesterday
    #.Notes
    #  Uses the CCD APIs which are new in Windows 7 and requires WDDM with display miniport drivers.
    #  https://msdn.microsoft.com/en-us/library/windows/hardware/hh406259%28v=vs.85%29.aspx
    [CmdletBinding()]
    param(
        # If you want to try the bing images from other countries, fiddle around with this.
        # As far as I know, the valid values are: en-US, zh-CN, ja-JP, en-AU, en-UK, de-DE, en-NZ, en-CA
        # NOTE: as far as I can tell, the images are usually the same, except offset by a day or two in some locales.
        [ValidateSet('es-AR','en-AU','de-AT','nl-BE','fr-BE','pt-BR','en-CA','fr-CA','fr-FR','de-DE','zh-HK','en-IN','en-ID','it-IT','ja-JP','ko-KR','en-MY','es-MX','nl-NL','nb-NO','zh-CN','pl-PL','ru-RU','ar-SA','en-ZA','es-ES','sv-SE','fr-CH','de-CH','zh-TW','tr-TR','en-GB','en-US','es-US')]
        [System.Globalization.CultureInfo]$Culture,
        # If you want to (re)use yesterday's wallpapers, fiddle around with this
        [Int]$Offset = 0,
        # Force re-downloading and creating the wallpaper even if none of the files have changed
        [Switch]$Force,

        # If set, supports writing text to the wallpaper in the bottom right corner
        [string]$Text,

        [string]$FontName = "Cascadia Code",

        [int]$FontSize = 18,

        [System.Drawing.FontStyle]$FontStyle = "Bold"
    )
    begin {
    }
    end {
        # Figure out how many wallpapers we need
        $Screens = @(Get-ActiveDisplays)

        # Use THAT information to calculate our virtual wallpaper size.
        # NOTE: See .Notes in Get-ActiveDisplays as to why we can't use SystemParameters here
        $count = $screens.Count
        $MinX, $MaxX = $Screens | % { $_.X, ($_.X + $_.Width) } | sort | select -first 1 -last 1
        $MinY, $MaxY = $Screens | % { $_.Y, ($_.Y + $_.Height) } | sort | select -first 1 -last 1

        $Width = $MaxX - $MinX
        $Height = $MaxY - $MinY
        $Top = $MinY
        $Left = $MinX

        # Fetch Bing's Image Archive information
        # It would be fun to tell you about these images, right?
        # TODO: We should do a notification with the details about the new wallpaper
        $APIUrl = "http://www.bing.com/HPImageArchive.aspx?format=js&idx=${Offset}&n=${count}"
        if($Culture) {
            $APIUrl += "&mkt=${Culture}"
        }

        $BingImages = Invoke-RestMethod $APIUrl

        $datespan = $Culture + $BingImages.images[-1].startdate + "-" + $BingImages.images[0].enddate

        $TempPath = [System.IO.Path]::GetTempPath()

        $WallPaperPath = Join-Path $TempPath "${datespan}.jpg"

        if(-not $Force -and (Test-Path $WallPaperPath)) {
            Write-Verbose "Update wallpaper from cached image file $WallPaperPath"
        } else {
            $ErrorActionPreference = "Stop"
            Write-Verbose ("Create {0}x{1} Wallpaper from Bing images" -f $Width, $Height)
            try {
                Write-Debug ("Full Wallpaper Size {0}x{1} offset to {2},{3}" -f $Width, $Height, $Left, $Top)
                $Wallpaper = New-Object System.Drawing.Bitmap ([int[]]($Width, $Height))
                $Graphics = [System.Drawing.Graphics]::FromImage($Wallpaper)

                for($i = 0; $i -lt $Count; $i++) {
                    $Size = "{0}x{1}" -f $Screens[$i].Width, $Screens[$i].Height
                    # Figure out the best image size available ...
                    Write-Debug "Actual Size $Size"
                    $Size = Get-Size $Size
                    Write-Debug "Image Size $Size"

                    # # I wanted to use ProjectOxford to trim the wallpaper, but it's output is grainy
                    # $OCPKey = (BetterCredentials\Get-Credential ComputerVisionApis@api.ProjectOxford.ai -Store).GetNetworkCredential().Password
                    # Invoke-WebRequest -OutFile $WallPaperPath -Method "POST" -ContentType "application/json" `
                    #     -Headers @{ "Ocp-Apim-Subscription-Key" = $OCPKey } `
                    #     -Body (ConvertTo-Json @{ Url = "http://www.bing.com/iod/1920/1200/{0:yyyyMMdd}" -f (get-date) }) `
                    #     -Uri "https://api.projectoxford.ai/vision/v1/thumbnails?width=1024&height=768&smartCropping=true"

                    $File = (Join-Path $TempPath $BingImages.Images[$i].fullstartdate) + "_" + $Size + ".jpg"

                    if(-not $Force -and (Test-Path $File)){
                        Write-Verbose "Using cached image file $File"
                    } else {
                        $ImageUrl = $urlbase + $BingImages.Images[$i].UrlBase + "_" + $Size + ".jpg"
                        Write-Verbose "Download Image $ImageUrl to $File"
                        Invoke-WebRequest $ImageUrl -OutFile $File
                    }

                    Write-Debug ("Place $(Split-Path $File -Leaf) at {0}, {1} for ({2},{3})" -f ($Screens[$i].X - $Left), ($Screens[$i].Y - $Top), $Screens[$i].Width, $Screens[$i].Height)

                    try {
                        $Source = [System.Drawing.Image]::FromFile($File)
                        # Putting the wallpaper in the right place, relatively speaking, is the tricky bit
                        $Graphics.DrawImage($Source, $Screens[$i].X - $Left, $Screens[$i].Y - $Top, $Screens[$i].Width, $Screens[$i].Height)
                    } finally {
                        $Source.Dispose()
                    }
                }
                if ($Text) {
                    # We want to make sure the text is within the bounds of the last screen (in a jagged screen layout, it's possible the actual CORNER of a rectangle is off-screen)
                    $LastScreenBounds = [System.Drawing.RectangleF]::new($Screens[-1].X - $Left + 20, $Screens[-1].Y - $Top + 20, $Screens[-1].Width - 40, $Screens[-1].Height - 80)
                    Write-OutlineText -Graphics $Graphics -Text $Text -FontName $FontName -FontSize $FontSize -FontStyle $FontStyle -Bounds $LastScreenBounds
                }
            } catch {
                Write-Warning "Unhandled Error: $_"
                Write-Warning "Unhandled Error: $($_.ScriptStackTrace)"
            } finally {
                $Graphics.Dispose()
                # Save as jpeg to save a little disk space
                Write-Verbose "Write wallpaper to cached image file $WallPaperPath"
                $Wallpaper.Save($WallPaperPath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
            }
        }
        # Tell windows about our new wallpaper ...
        Set-ItemProperty "HKCU:\Control Panel\Desktop" WallpaperStyle 1
        Set-ItemProperty "HKCU:\Control Panel\Desktop" TileWallpaper 1
        # Please excuse the magic numbers, I can't remember the constants anymore
        # $Result should be 1, but we're not checking it, because what would we do about it?
        $Result = [Windows]::SystemParametersInfo( 20, 0, $WallPaperPath, 3 )
    }
}
