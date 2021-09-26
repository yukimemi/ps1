Write-Host "[info] Currently ordering a double shot of espresso..."

$sig = @"
[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern void SetThreadExecutionState(uint esFlags);
"@

$ES_CONTINUOUS = [uint32]"0x80000000"
$ES_DISPLAY_REQUIRED = [uint32]"0x00000002"
$ES_SYSTEM_REQUIRED = [uint32]"0x00000001"

$stes = Add-Type -MemberDefinition $sig -Name System -Namespace Win32 -PassThru

[void]$stes::SetThreadExecutionState($ES_SYSTEM_REQUIRED -bor $ES_DISPLAY_REQUIRED -bor $ES_CONTINUOUS)

Read-Host "[info] Enter if you want to exit ..."
Write-Host "[info] No more espressos left behind the counter."
