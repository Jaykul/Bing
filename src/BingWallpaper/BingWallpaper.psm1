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
    1    = @('240x240','320x320')
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
        [Switch]$Force
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

                    $File = Join-Path $TempPath ((Split-Path $BingImages.Images[$i].UrlBase -Leaf) + "_" + $Size + ".jpg")

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
            } finally {
                $Graphics.Dispose()
                # Save as jpeg to save a little disk space
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

# SIG # Begin signature block
# MIIXxAYJKoZIhvcNAQcCoIIXtTCCF7ECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUmudFgMprEtmT5fta354ZUoTh
# 8ASgghL3MIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
# AQUFADCBizELMAkGA1UEBhMCWkExFTATBgNVBAgTDFdlc3Rlcm4gQ2FwZTEUMBIG
# A1UEBxMLRHVyYmFudmlsbGUxDzANBgNVBAoTBlRoYXd0ZTEdMBsGA1UECxMUVGhh
# d3RlIENlcnRpZmljYXRpb24xHzAdBgNVBAMTFlRoYXd0ZSBUaW1lc3RhbXBpbmcg
# Q0EwHhcNMTIxMjIxMDAwMDAwWhcNMjAxMjMwMjM1OTU5WjBeMQswCQYDVQQGEwJV
# UzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFu
# dGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBALGss0lUS5ccEgrYJXmRIlcqb9y4JsRDc2vCvy5Q
# WvsUwnaOQwElQ7Sh4kX06Ld7w3TMIte0lAAC903tv7S3RCRrzV9FO9FEzkMScxeC
# i2m0K8uZHqxyGyZNcR+xMd37UWECU6aq9UksBXhFpS+JzueZ5/6M4lc/PcaS3Er4
# ezPkeQr78HWIQZz/xQNRmarXbJ+TaYdlKYOFwmAUxMjJOxTawIHwHw103pIiq8r3
# +3R8J+b3Sht/p8OeLa6K6qbmqicWfWH3mHERvOJQoUvlXfrlDqcsn6plINPYlujI
# fKVOSET/GeJEB5IL12iEgF1qeGRFzWBGflTBE3zFefHJwXECAwEAAaOB+jCB9zAd
# BgNVHQ4EFgQUX5r1blzMzHSa1N197z/b7EyALt0wMgYIKwYBBQUHAQEEJjAkMCIG
# CCsGAQUFBzABhhZodHRwOi8vb2NzcC50aGF3dGUuY29tMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwPwYDVR0fBDgwNjA0oDKgMIYuaHR0cDovL2NybC50aGF3dGUuY29tL1Ro
# YXd0ZVRpbWVzdGFtcGluZ0NBLmNybDATBgNVHSUEDDAKBggrBgEFBQcDCDAOBgNV
# HQ8BAf8EBAMCAQYwKAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0y
# MDQ4LTEwDQYJKoZIhvcNAQEFBQADgYEAAwmbj3nvf1kwqu9otfrjCR27T4IGXTdf
# plKfFo3qHJIJRG71betYfDDo+WmNI3MLEm9Hqa45EfgqsZuwGsOO61mWAK3ODE2y
# 0DGmCFwqevzieh1XTKhlGOl5QGIllm7HxzdqgyEIjkHq3dlXPx13SYcqFgZepjhq
# IhKjURmDfrYwggSjMIIDi6ADAgECAhAOz/Q4yP6/NW4E2GqYGxpQMA0GCSqGSIb3
# DQEBBQUAMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBD
# QSAtIEcyMB4XDTEyMTAxODAwMDAwMFoXDTIwMTIyOTIzNTk1OVowYjELMAkGA1UE
# BhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTQwMgYDVQQDEytT
# eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIFNpZ25lciAtIEc0MIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAomMLOUS4uyOnREm7Dv+h8GEKU5Ow
# mNutLA9KxW7/hjxTVQ8VzgQ/K/2plpbZvmF5C1vJTIZ25eBDSyKV7sIrQ8Gf2Gi0
# jkBP7oU4uRHFI/JkWPAVMm9OV6GuiKQC1yoezUvh3WPVF4kyW7BemVqonShQDhfu
# ltthO0VRHc8SVguSR/yrrvZmPUescHLnkudfzRC5xINklBm9JYDh6NIipdC6Anqh
# d5NbZcPuF3S8QYYq3AhMjJKMkS2ed0QfaNaodHfbDlsyi1aLM73ZY8hJnTrFxeoz
# C9Lxoxv0i77Zs1eLO94Ep3oisiSuLsdwxb5OgyYI+wu9qU+ZCOEQKHKqzQIDAQAB
# o4IBVzCCAVMwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAO
# BgNVHQ8BAf8EBAMCB4AwcwYIKwYBBQUHAQEEZzBlMCoGCCsGAQUFBzABhh5odHRw
# Oi8vdHMtb2NzcC53cy5zeW1hbnRlYy5jb20wNwYIKwYBBQUHMAKGK2h0dHA6Ly90
# cy1haWEud3Muc3ltYW50ZWMuY29tL3Rzcy1jYS1nMi5jZXIwPAYDVR0fBDUwMzAx
# oC+gLYYraHR0cDovL3RzLWNybC53cy5zeW1hbnRlYy5jb20vdHNzLWNhLWcyLmNy
# bDAoBgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMjAdBgNV
# HQ4EFgQURsZpow5KFB7VTNpSYxc/Xja8DeYwHwYDVR0jBBgwFoAUX5r1blzMzHSa
# 1N197z/b7EyALt0wDQYJKoZIhvcNAQEFBQADggEBAHg7tJEqAEzwj2IwN3ijhCcH
# bxiy3iXcoNSUA6qGTiWfmkADHN3O43nLIWgG2rYytG2/9CwmYzPkSWRtDebDZw73
# BaQ1bHyJFsbpst+y6d0gxnEPzZV03LZc3r03H0N45ni1zSgEIKOq8UvEiCmRDoDR
# EfzdXHZuT14ORUZBbg2w6jiasTraCXEQ/Bx5tIB7rGn0/Zy2DBYr8X9bCT2bW+IW
# yhOBbQAuOA2oKY8s4bL0WqkBrxWcLC9JG9siu8P+eJRRw4axgohd8D20UaF5Mysu
# e7ncIAkTcetqGVvP6KUwVyyJST+5z3/Jvz4iaGNTmr1pdKzFHTx/kuDDvBzYBHUw
# ggUmMIIEDqADAgECAhACXbrxBhFj1/jVxh2rtd9BMA0GCSqGSIb3DQEBCwUAMHIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJ
# RCBDb2RlIFNpZ25pbmcgQ0EwHhcNMTUwNTA0MDAwMDAwWhcNMTYwNTExMTIwMDAw
# WjBtMQswCQYDVQQGEwJVUzERMA8GA1UECBMITmV3IFlvcmsxFzAVBgNVBAcTDldl
# c3QgSGVucmlldHRhMRgwFgYDVQQKEw9Kb2VsIEguIEJlbm5ldHQxGDAWBgNVBAMT
# D0pvZWwgSC4gQmVubmV0dDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# AJfRKhfiDjMovUELYgagznWf+HFcDENk118Y/K6UkQDwKmVyVOvDyaVefjSmZZcV
# NZqqYpm9d/Iajf2dauyC3pg3oay8KfXAADLHgbmbvYDc5zGuUNsTzMUOKlp9h13c
# qsg898JwpRpI659xCQgJjZ6V83QJh+wnHvjA9ojjA4xkbwhGp4Eit6B/uGthEA11
# IHcFcXeNI3fIkbwWiAw7ZoFtSLm688NFhxwm+JH3Xwj0HxuezsmU0Yc/po31CoST
# nGPVN8wppHYZ0GfPwuNK4TwaI0FEXxwdwB+mEduxa5e4zB8DyUZByFW338XkGfc1
# qcJJ+WTyNKFN7saevhwp02cCAwEAAaOCAbswggG3MB8GA1UdIwQYMBaAFFrEuXsq
# CqOl6nEDwGD5LfZldQ5YMB0GA1UdDgQWBBQV0aryV1RTeVOG+wlr2Z2bOVFAbTAO
# BgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1
# oDOgMYYvaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1n
# MS5jcmwwNaAzoDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3Vy
# ZWQtY3MtZzEuY3JsMEIGA1UdIAQ7MDkwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUH
# AgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwgYQGCCsGAQUFBwEBBHgw
# djAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUF
# BzAChkJodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNz
# dXJlZElEQ29kZVNpZ25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0B
# AQsFAAOCAQEAIi5p+6eRu6bMOSwJt9HSBkGbaPZlqKkMd4e6AyKIqCRabyjLISwd
# i32p8AT7r2oOubFy+R1LmbBMaPXORLLO9N88qxmJfwFSd+ZzfALevANdbGNp9+6A
# khe3PiR0+eL8ZM5gPJv26OvpYaRebJTfU++T1sS5dYaPAztMNsDzY3krc92O27AS
# WjTjWeILSryqRHXyj8KQbYyWpnG2gWRibjXi5ofL+BHyJQRET5pZbERvl2l9Bo4Z
# st8CM9EQDrdG2vhELNiA6jwenxNPOa6tPkgf8cH8qpGRBVr9yuTMSHS1p9Rc+ybx
# FSKiZkOw8iCR6ZQIeKkSVdwFf8V+HHPrETCCBTAwggQYoAMCAQICEAQJGBtf1btm
# dVNDtW+VUAgwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoT
# DERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UE
# AxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTEzMTAyMjEyMDAwMFoX
# DTI4MTAyMjEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0
# IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNl
# cnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsxSRnP0PtFmbE620T1f+Wo
# ndsy13Hqdp0FLreP+pJDwKX5idQ3Gde2qvCchqXYJawOeSg6funRZ9PG+yknx9N7
# I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJRdQtoaPpiCwgla4cSocI
# 3wz14k1gGL6qxLKucDFmM3E+rHCiq85/6XzLkqHlOzEcz+ryCuRXu0q16XTmK/5s
# y350OTYNkO/ktU6kqepqCquE86xnTrXE94zRICUj6whkPlKWwfIPEvTFjg/Bougs
# UfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8lk9ECAwEAAaOCAc0wggHJ
# MBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4
# MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVk
# SURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURSb290Q0EuY3JsME8GA1UdIARIMEYwOAYKYIZIAYb9bAAC
# BDAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAoG
# CGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAfBgNVHSME
# GDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQsFAAOCAQEAPuwN
# WiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh134LYP3DPQ/Er4v97yrfIFU3sOH20ZJ1
# D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63XX0R58zYUBor3nEZOXP+QsRsHDpEV+7qv
# tVHCjSSuJMbHJyqhKSgaOnEoAjwukaPAJRHinBRHoXpoaK+bp1wgXNlxsQyPu6j4
# xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW6Fkd6fp0ZGuy62ZD2rOw
# jNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ92JuoVP6EpQYhS6Skepo
# bEQysmah5xikmmRR7zGCBDcwggQzAgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EAJduvEGEWPX+NXGHau130EwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFJkeJLEYRPoiFZyEhcVW
# 0L0/qIWxMA0GCSqGSIb3DQEBAQUABIIBAIFY3TRL4zJVMK9AO83Raf44JAWpN4Sp
# jhVmN7maachyl84rUFtqaMZ/vDj9SbTs+lvrwOkYVLpjlFU3EM5tohgWJXWtuHdy
# XWTLf1BsPUL9fgluvw3NGJdEkdbrc5pAFA47p5dmRhUViS9KFlYE/iz35sQpadQJ
# Eno/+sEk4WK7iwu7DsAyfJQ81GSEpr6amO5olzd5lobQFW6xIqmwLAcYzfikcDGc
# n89eB1V5Vk8vTUxLMvguXD0kYbigK9BvLObeN6kGFU2W8DA4fTe0xQuOZ4hKTfaS
# bvsWR3mwoz/Z2G8uN1vnCQqt4dfMmRFzWWh98S5hqcYFywspa/PTgkGhggILMIIC
# BwYJKoZIhvcNAQkGMYIB+DCCAfQCAQEwcjBeMQswCQYDVQQGEwJVUzEdMBsGA1UE
# ChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFudGVjIFRpbWUg
# U3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMgIQDs/0OMj+vzVuBNhqmBsaUDAJBgUr
# DgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMTUxMDA5MTg1MjQxWjAjBgkqhkiG9w0BCQQxFgQU5JFjuGoWYIPm6DchiaQB
# Z74Gn1gwDQYJKoZIhvcNAQEBBQAEggEACiXBvHdLOobRwj7AhjKuOLn6vhChSyiv
# iO202qbfWeMA26GxbU4XLWf4Odi7Iupyuf42S9I8H0vn+P8L5MElqubDbJ8bQHNR
# 1N96EDd5dOW7BK8yGuCNH9cqXGPwQRhP7atuIdTLQiim+Z8/YZmEn5kv0/SpOlRJ
# sRC0cR7I8r7GtIQOmlHD4y/epW/YMuulpLq/9hwoUtTYABhGT4rE53HA2IWAb/MD
# vV8I/ZR2JFPkjGUYVrJPnDQQZkX5FYN2UqoOvgwtYH+sdI46wTto+M5qeu6f+4q5
# YwbO+3VQrArlCoAh7TTq1pc8iTmZu84aCQ+zkLaG+qgP2ZkuSl0kYA==
# SIG # End signature block
