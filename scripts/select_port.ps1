$ports = @(5000, 5001)

foreach ($port in $ports) {
    $listener = $null
    try {
        $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Parse("127.0.0.1"), $port)
        $listener.Start()
        $listener.Stop()
        Write-Output $port
        exit 0
    } catch {
        if ($listener) {
            $listener.Stop()
        }
    }
}

Write-Error "5000 and 5001 are both occupied."
exit 1
