param (
$interval = 30,
$localPath = "Y:",
$remotePath = "/",
$FileMask = "*.plot",
$TlsFingerprint,
[Switch]
$ftpes,
$password,
$user,
$hostname
)

$length = 0

function HandleException ($e)
{
    Write-Host -ForegroundColor Red $_.Exception.Message
    Beep
}

function Beep()
{
    [System.Console]::Beep()
}

function SetConsoleTitle ($status)
{
    if ($sessionOptions)
    {
        $status = "$($sessionOptions.ToString()) - $status"
    }
    $Host.UI.RawUI.WindowTitle = $status
}

# Session.FileTransferProgress event handler

function FileTransferProgress
{
    param($e)
	$CPS = $e.CPS / 1000000
	$progress = $e.FileProgress * 100
	
	if ($e.CPS)
		{
			$remainingSeconds = $length * (1 - $e.FileProgress) / $e.CPS
			$ts =  [timespan]::fromseconds($remainingSeconds)
			$remainingTime = $ts.ToString("hh\h\ mm\m")
        }
	
    Write-Progress `
        -Activity $e.FileName -Status ("{1:p0} {0:n2} MB/s {2}" -f $CPS, $e.FileProgress, $remainingTime) `
        -PercentComplete ($progress)
	
	$TaskBarObject.SetProgressValue($progress,100)
}

try
{
	[Reflection.Assembly]::LoadFrom("C:\api\Microsoft.WindowsAPICodePack.Shell.dll") | Out-Null
	$TaskBarObject = [Microsoft.WindowsAPICodePack.TaskBar.TaskBarManager]::Instance
	
    # Load WinSCP .NET assembly
    $assemblyPath = if ($env:WINSCP_PATH) { $env:WINSCP_PATH } else { $PSScriptRoot }
    Add-Type -Path (Join-Path $assemblyPath "WinSCPnet.dll")

	# Session config
	$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
    Protocol = [WinSCP.Protocol]::Ftp
    HostName = $hostname
    UserName = $user
    Password = $password
	}
	if ($ftpes)
		{		
			$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
			Protocol = [WinSCP.Protocol]::Ftp
			HostName = $hostname
			UserName = $user
			Password = $password
			FtpSecure = [WinSCP.FtpSecure]::Explicit
			TlsHostCertificateFingerprint = $TlsFingerprint
			}
		}


    $session = New-Object WinSCP.Session
    
    try
    {
		# Will continuously report progress of transfer
        $session.add_FileTransferProgress( { FileTransferProgress($_) } )
        
		Write-Host "Connecting..."
        SetConsoleTitle "Connecting"
        $session.Open($sessionOptions)

        while ($True)
        {
            Write-Host "`n`n`n`n`n`nLooking for changes..."
            SetConsoleTitle "Looking for changes"
            try
            {
                $transferOptions = New-Object WinSCP.TransferOptions
				$transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
				$transferOptions.FileMask = $FileMask
				$transferOptions.ResumeSupport.State = [WinSCP.TransferResumeSupportState]::On
				
				$differences =
                    $session.CompareDirectories(
                        [WinSCP.SynchronizationMode]::Local, $localPath, $remotePath, $False, $False, [WinSCP.SynchronizationCriteria]::Size, $transferOptions)

                Write-Host
                if ($differences.Count -eq 0)
                {
                    Write-Host "No changes found."   
                }
                else
                {
                    Write-Host "Synchronizing $($differences.Count) change(s)..."
                    SetConsoleTitle "Synchronizing changes"

                    foreach ($difference in $differences)
                    {
                        $action = $difference.Action
                        if ($action -eq [WinSCP.SynchronizationAction]::DownloadNew)
                        {
                            $message = "Downloading new $($difference.Remote.FileName)..."
                        }
                        elseif ($action -eq [WinSCP.SynchronizationAction]::DownloadUpdate)
                        {
                            $message = "Downloading updated $($difference.Remote.FileName)..."
                        }
                        elseif ($action -eq [WinSCP.SynchronizationAction]::DeleteLocal)
                        {
                            $message = "Deleting $($difference.Local.FileName)..."
                        }
                        else
                        {
                            throw "Unexpected difference $action"
                        }

                        Write-Host $message
						$length = $difference.Remote.Length

                        try
                        {
                            $transfer = $difference.Resolve($session) | Out-Null
							Write-Host "Download succeeded, removing from source."
							$filename = [WinSCP.RemotePath]::EscapeFileMask($difference.Remote.FileName)
							$removalResult = $session.RemoveFiles($filename)
							if ($removalResult.IsSuccess)
								{
									Write-Host "Done."
								}
							else
								{
									Write-Host "Removing of file $($difference.Remote.FileName) failed."
								}
							Write-Host "Replacing extension."
							Get-ChildItem $transfer.Destination | Rename-Item -NewName { $_.name -Replace '\.plot$','.plot.2' }
                            Write-Host "Done."
							
                        }
                        catch
                        {
                            Write-Host
                            HandleException $_
                        }
                    }
                }
            }
            catch
            {
                Write-Host
                HandleException $_
            }

            SetConsoleTitle "Waiting"
            $wait = [int]$interval
            # Wait for 1 second in a loop, to make the waiting breakable
            while ($wait -gt 0)
            {
                Write-Host -NoNewLine "`rWaiting for $wait seconds, press Ctrl+C to abort... "
                Start-Sleep -Seconds 1
                $wait--
            }

            Write-Host
            Write-Host
        }
    }
    finally
    {
        Write-Host # to break after "Waiting..." status
        Write-Host "Disconnecting..."
        # Disconnect, clean up
        $session.Dispose()
    }
}
catch
{
    HandleException $_
    SetConsoleTitle "Error"
}


Write-Host "Press any key to exit..."
[System.Console]::ReadKey() | Out-Null


# Never exits cleanly
exit 1
