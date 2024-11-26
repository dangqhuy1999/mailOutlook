# Define computer names and paths
$localComputer = "MT-1417"
$remoteComputer = "MT-1418"
$iperfExePathLocal = "D:\IT-Only\perf\iperf3.exe"
$iperfExePathRemote = "D:\IT-Only\perf\iperf3.exe"
$remoteShare = "\\truenas\IT Data\Speedtest"
$logFile = "$remoteShare\speedtest_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Function to start iperf server on local computer
function Start-IperfServer {
    param($iperfExePath)
    if (Test-Path -Path $iperfExePath) {
        Write-Host "Starting iperf server at $iperfExePath"
        Start-Process -FilePath $iperfExePath -ArgumentList "-s" -NoNewWindow -PassThru
    } else {
        Write-Host "iperf3.exe not found at $iperfExePath, cannot start server"
    }
}

# Start iperf server
Start-IperfServer -iperfExePath $iperfExePathLocal

# Run iperf client on remote computer and save the result to a log file
$iperfResult = Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
    param($iperfExePath, $localComputer)
    if (Test-Path -Path $iperfExePath) {
        Write-Host "Running iperf client from $iperfExePath"
        & $iperfExePath -c $localComputer
    } else {
        Write-Host "iperf3.exe not found at $iperfExePath, cannot run client"
    }
} -ArgumentList $iperfExePathRemote, $localComputer
$iperfResult | Out-File -FilePath $logFile -Force
