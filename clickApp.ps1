param (
    [Parameter(Mandatory=$true, Position=0)]
    [int] $X,
    [Parameter(Mandatory=$true, Position=1)]
    [int] $Y
)


    $signature=@'
    [DllImport("user32.dll",CharSet=CharSet.Auto,CallingConvention=CallingConvention.StdCall)]
    public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@
$SendMouseClick = Add-Type -memberDefinition $signature -name "Win32MouseEventNew" -namespace Win32Functions -passThru

Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject wscript.shell


function Make-Move {
    Param (
        [Parameter(Mandatory=$true, Position=0)]
        [int] $X,
        [Parameter(Mandatory=$true, Position=1)]
        [int] $Y
    )

    $POSITION = [Windows.Forms.Cursor]::Position
    $POSITION.x = $X
    $POSITION.y = $Y+3

    [Windows.Forms.Cursor]::Position = $POSITION
    Start-Sleep -s 2
    $SendMouseClick::mouse_event(0x0002, 0, 0, 0, 0);
    $SendMouseClick::mouse_event(0x0004, 0, 0, 0, 0);
}

function ConnectVPN {
    if ($(Get-VpnConnection -Name "VPN").ConnectionStatus -eq "Disconnected") {
        echo "Disconnected"
        rasphone
        Start-Sleep 2
        $wshell.SendKeys('{ENTER}')
    }
    Start-Sleep 5
}

function Connect {
    if (Test-Connection -ComputerName jira -Quiet) {
        echo "Connected"
        Start-Sleep 4
    }
    else {
        echo "Connecting"
        rasphone
        Start-Sleep 3
        $wshell.SendKeys('{ENTER}')
    }
}


Connect
powershell.exe -Command "& 'C:\Program Files (x86)\work.jar'"
Start-Sleep -s 4
Make-Move -X $X -Y $Y

Start-Sleep 2
$wshell.SendKeys('{ENTER}')

#Start-Sleep -s 4
#Make-Move -X 465 -Y 75
