[CmdletBinding()]
param(
    [string]$Ip = "",
    [int]$Port = 8080,
    [string]$DllPath = "",
    [int]$DurationSec = 300,
    [int]$IntervalMs = 500,
    [string]$OutFile = ""
)

$ErrorActionPreference = "Stop"

function Find-DllPath {
    $candidatePaths = @(
        $env:FAIRINO_DLL_PATH,
        (Join-Path (Get-Location) "libfairino.dll"),
        (Join-Path (Get-Location) "Assets\\Plugins\\Fairino\\libfairino.dll"),
        (Join-Path $env:USERPROFILE "Desktop\\libfairino.dll"),
        (Join-Path $env:USERPROFILE "Documents\\libfairino.dll"),
        (Join-Path $env:USERPROFILE "Downloads\\libfairino.dll")
    )

    foreach ($candidate in $candidatePaths) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path $candidate -PathType Leaf)) {
            return [System.IO.Path]::GetFullPath($candidate)
        }
    }

    foreach ($root in @((Join-Path $env:USERPROFILE "Desktop"), (Join-Path $env:USERPROFILE "Documents"), (Join-Path $env:USERPROFILE "Downloads"))) {
        if (-not (Test-Path $root -PathType Container)) { continue }
        $found = Get-ChildItem -Path $root -Filter libfairino.dll -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $found) { return $found.FullName }
    }

    return $null
}

function Resolve-InputDefaults {
    if ([string]::IsNullOrWhiteSpace($script:Ip)) { $script:Ip = $env:FAIRINO_IP }
    if ([string]::IsNullOrWhiteSpace($script:Ip)) { $script:Ip = "192.168.58.2" }

    if (-not $PSBoundParameters.ContainsKey("Port") -and $env:FAIRINO_PORT) {
        $parsedPort = 0
        if ([int]::TryParse($env:FAIRINO_PORT, [ref]$parsedPort)) { $script:Port = $parsedPort }
    }

    if ([string]::IsNullOrWhiteSpace($script:DllPath)) { $script:DllPath = Find-DllPath }
    if ([string]::IsNullOrWhiteSpace($script:DllPath)) {
        throw "libfairino.dll not found automatically. Pass -DllPath, set FAIRINO_DLL_PATH, or place the DLL in Desktop/Documents/Downloads."
    }
    if ((Split-Path $script:DllPath -Leaf) -ne "libfairino.dll" -and (Test-Path $script:DllPath -PathType Container)) {
        $script:DllPath = Join-Path $script:DllPath "libfairino.dll"
    }
    $script:DllPath = [System.IO.Path]::GetFullPath($script:DllPath)
    if (-not (Test-Path $script:DllPath -PathType Leaf)) { throw "libfairino.dll not found: $script:DllPath" }

    $script:DurationSec = [Math]::Max(1, $script:DurationSec)
    $script:IntervalMs = [Math]::Max(100, $script:IntervalMs)

    if ([string]::IsNullOrWhiteSpace($script:OutFile)) {
        $timestamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
        $script:OutFile = Join-Path $env:USERPROFILE "Desktop\\fr5-record-$timestamp.jsonl"
    } else {
        $outDir = Split-Path $script:OutFile -Parent
        if (-not [string]::IsNullOrWhiteSpace($outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
    }

    $script:OutFile = [System.IO.Path]::GetFullPath($script:OutFile)
}

function Get-Method([Type]$type, [string]$name, [Type[]]$paramTypes) {
    $method = $type.GetMethod($name, $paramTypes)
    if ($null -eq $method) { throw "Method not found: $name" }
    return $method
}

function New-DescPose([Type]$descPoseType, [Type]$descTranType, [Type]$rpyType) {
    $descPose = [Activator]::CreateInstance($descPoseType)
    $tran = [Activator]::CreateInstance($descTranType)
    $rpy = [Activator]::CreateInstance($rpyType)
    $descPoseType.GetField("tran").SetValue($descPose, $tran)
    $descPoseType.GetField("rpy").SetValue($descPose, $rpy)
    return $descPose
}

function New-JointPos([Type]$jointPosType) {
    $jointPos = [Activator]::CreateInstance($jointPosType)
    $jointPosType.GetField("jPos").SetValue($jointPos, [double[]]@(0,0,0,0,0,0))
    return $jointPos
}

function Convert-DescPose($descPose) {
    if ($null -eq $descPose) { return $null }
    $type = $descPose.GetType()
    $tran = $type.GetField("tran").GetValue($descPose)
    $rpy = $type.GetField("rpy").GetValue($descPose)
    return [ordered]@{
        x = [double]$tran.GetType().GetField("x").GetValue($tran)
        y = [double]$tran.GetType().GetField("y").GetValue($tran)
        z = [double]$tran.GetType().GetField("z").GetValue($tran)
        rx = [double]$rpy.GetType().GetField("rx").GetValue($rpy)
        ry = [double]$rpy.GetType().GetField("ry").GetValue($rpy)
        rz = [double]$rpy.GetType().GetField("rz").GetValue($rpy)
    }
}

function Convert-RealtimeState($pkg) {
    if ($null -eq $pkg) { return $null }
    $type = $pkg.GetType()
    return [ordered]@{
        robotState = [int]$type.GetField("robot_state").GetValue($pkg)
        programState = [int]$type.GetField("program_state").GetValue($pkg)
        mainCode = [int]$type.GetField("main_code").GetValue($pkg)
        subCode = [int]$type.GetField("sub_code").GetValue($pkg)
        robotMode = [int]$type.GetField("robot_mode").GetValue($pkg)
        toolId = [int]$type.GetField("tool").GetValue($pkg)
        userId = [int]$type.GetField("user").GetValue($pkg)
        emergencyStop = ([int]$type.GetField("EmergencyStop").GetValue($pkg)) -ne 0
        collisionState = ([int]$type.GetField("collisionState").GetValue($pkg)) -ne 0
        safetyStop0 = ([int]$type.GetField("safety_stop0_state").GetValue($pkg)) -ne 0
        safetyStop1 = ([int]$type.GetField("safety_stop1_state").GetValue($pkg)) -ne 0
        enabledState = [int]$type.GetField("rbtEnableState").GetValue($pkg)
        motionQueueLength = [int]$type.GetField("mc_queue_len").GetValue($pkg)
        jointPosDeg = @($type.GetField("jt_cur_pos").GetValue($pkg))
        tcpPose = @($type.GetField("tl_cur_pos").GetValue($pkg))
        flangePose = @($type.GetField("flange_cur_pos").GetValue($pkg))
        toolCoord = @($type.GetField("toolCoord").GetValue($pkg))
        wobjCoord = @($type.GetField("wobjCoord").GetValue($pkg))
        gripperPosition = [int]$type.GetField("gripper_position").GetValue($pkg)
        gripperFault = [int]$type.GetField("gripper_fault").GetValue($pkg)
    }
}

function Write-JsonLine($obj) {
    $json = $obj | ConvertTo-Json -Depth 10 -Compress
    Add-Content -Path $script:OutFile -Value $json -Encoding UTF8
}
function Capture-Sample {
    param(
        [object]$robot,
        [Type]$jointPosType,
        [Type]$descPoseType,
        [Type]$descTranType,
        [Type]$rpyType,
        [Type]$realtimeStateType,
        $methods,
        [int]$sampleIndex
    )

    $sample = [ordered]@{
        type = "sample"
        sampleIndex = $sampleIndex
        capturedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
        ip = $script:Ip
        port = $script:Port
        realtimeState = $null
        actualJointPose = $null
        actualTcpPose = $null
        safetyCode = $null
        realtimeSamplePeriod = $null
        coordContext = $null
        controllerFault = $null
    }

    $pkg = [Activator]::CreateInstance($realtimeStateType)
    $realtimeArgs = [object[]]@($pkg)
    $realtimeCode = [int]$methods.GetRobotRealTimeState.Invoke($robot, $realtimeArgs)
    $realtimePayload = $null
    if ($realtimeCode -eq 0) { $realtimePayload = Convert-RealtimeState $realtimeArgs[0] }
    $sample.realtimeState = [ordered]@{ code = $realtimeCode; payload = $realtimePayload }

    $jointPos = New-JointPos $jointPosType
    $jointArgs = [object[]]@([byte]0, $jointPos)
    $jointCode = [int]$methods.GetActualJointPosDegree.Invoke($robot, $jointArgs)
    $jointPayload = $null
    if ($jointCode -eq 0) { $jointPayload = @($jointPosType.GetField("jPos").GetValue($jointArgs[1])) }
    $sample.actualJointPose = [ordered]@{ code = $jointCode; jointPosDeg = $jointPayload }

    $tcpPose = New-DescPose $descPoseType $descTranType $rpyType
    $tcpArgs = [object[]]@([byte]0, $tcpPose)
    $tcpCode = [int]$methods.GetActualTCPPose.Invoke($robot, $tcpArgs)
    $tcpPayload = $null
    if ($tcpCode -eq 0) { $tcpPayload = Convert-DescPose $tcpArgs[1] }
    $sample.actualTcpPose = [ordered]@{ code = $tcpCode; tcpPose = $tcpPayload }

    $sample.safetyCode = [ordered]@{ code = 0; safetyCode = [int]$methods.GetSafetyCode.Invoke($robot, [object[]]@()) }

    $periodArgs = [object[]]@(0)
    $periodCode = [int]$methods.GetRobotRealtimeStateSamplePeriod.Invoke($robot, $periodArgs)
    $periodValue = $null
    if ($periodCode -eq 0) { $periodValue = [int]$periodArgs[0] }
    $sample.realtimeSamplePeriod = [ordered]@{ code = $periodCode; periodMs = $periodValue }

    $toolNumArgs = [object[]]@([byte]0, 0)
    $toolNumCode = [int]$methods.GetActualTCPNum.Invoke($robot, $toolNumArgs)
    $wobjNumArgs = [object[]]@([byte]0, 0)
    $wobjNumCode = [int]$methods.GetActualWObjNum.Invoke($robot, $wobjNumArgs)

    $toolCoord = New-DescPose $descPoseType $descTranType $rpyType
    $toolCoordArgs = [object[]]@($toolCoord)
    $toolCoordCode = [int]$methods.GetCurToolCoord.Invoke($robot, $toolCoordArgs)

    $wobjCoord = New-DescPose $descPoseType $descTranType $rpyType
    $wobjCoordArgs = [object[]]@($wobjCoord)
    $wobjCoordCode = [int]$methods.GetCurWObjCoord.Invoke($robot, $wobjCoordArgs)

    $toolIdValue = $null
    if ($toolNumCode -eq 0) { $toolIdValue = [int]$toolNumArgs[1] }
    $userIdValue = $null
    if ($wobjNumCode -eq 0) { $userIdValue = [int]$wobjNumArgs[1] }
    $toolCoordValue = $null
    if ($toolCoordCode -eq 0) { $toolCoordValue = Convert-DescPose $toolCoordArgs[0] }
    $wobjCoordValue = $null
    if ($wobjCoordCode -eq 0) { $wobjCoordValue = Convert-DescPose $wobjCoordArgs[0] }

    $sample.coordContext = [ordered]@{
        toolIdCode = $toolNumCode
        toolId = $toolIdValue
        userIdCode = $wobjNumCode
        userId = $userIdValue
        toolCoordCode = $toolCoordCode
        toolCoord = $toolCoordValue
        wobjCoordCode = $wobjCoordCode
        wobjCoord = $wobjCoordValue
    }

    $faultArgs = [object[]]@(0, 0)
    $faultCode = [int]$methods.GetRobotErrorCode.Invoke($robot, $faultArgs)
    $safetyStopArgs = [object[]]@([byte]0, [byte]0)
    $safetyStopCode = [int]$methods.GetSafetyStopState.Invoke($robot, $safetyStopArgs)
    $dragArgs = [object[]]@([byte]0)
    $dragCode = [int]$methods.IsInDragTeach.Invoke($robot, $dragArgs)

    $mainCode = $null
    $subCode = $null
    if ($faultCode -eq 0) {
        $mainCode = [int]$faultArgs[0]
        $subCode = [int]$faultArgs[1]
    }
    $safetyStop0 = $null
    $safetyStop1 = $null
    if ($safetyStopCode -eq 0) {
        $safetyStop0 = ([int]$safetyStopArgs[0]) -ne 0
        $safetyStop1 = ([int]$safetyStopArgs[1]) -ne 0
    }
    $dragState = $null
    if ($dragCode -eq 0) {
        $dragState = ([int]$dragArgs[0]) -ne 0
    }

    $sample.controllerFault = [ordered]@{
        faultCode = $faultCode
        mainCode = $mainCode
        subCode = $subCode
        safetyStopCode = $safetyStopCode
        safetyStop0 = $safetyStop0
        safetyStop1 = $safetyStop1
        dragTeachCode = $dragCode
        isInDragTeach = $dragState
    }

    return $sample
}
Resolve-InputDefaults
$encoding = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($OutFile, "", $encoding)

$assembly = [System.Reflection.Assembly]::LoadFrom($DllPath)
$robotType = $assembly.GetType("fairino.Robot")
if ($null -eq $robotType) { throw "Type fairino.Robot not found in $DllPath" }

$jointPosType = $assembly.GetType("fairino.JointPos")
$descPoseType = $assembly.GetType("fairino.DescPose")
$descTranType = $assembly.GetType("fairino.DescTran")
$rpyType = $assembly.GetType("fairino.Rpy")
$realtimeStateType = $assembly.GetType("fairino.ROBOT_STATE_PKG")
$robot = [Activator]::CreateInstance($robotType)

$methods = [ordered]@{
    RPC = Get-Method $robotType "RPC" ([Type[]]@([string]))
    CloseRPC = Get-Method $robotType "CloseRPC" ([Type[]]@())
    GetRobotRealTimeState = Get-Method $robotType "GetRobotRealTimeState" ([Type[]]@($realtimeStateType.MakeByRefType()))
    GetActualJointPosDegree = Get-Method $robotType "GetActualJointPosDegree" ([Type[]]@([byte], $jointPosType.MakeByRefType()))
    GetActualTCPPose = Get-Method $robotType "GetActualTCPPose" ([Type[]]@([byte], $descPoseType.MakeByRefType()))
    GetSafetyCode = Get-Method $robotType "GetSafetyCode" ([Type[]]@())
    GetRobotRealtimeStateSamplePeriod = Get-Method $robotType "GetRobotRealtimeStateSamplePeriod" ([Type[]]@([int].MakeByRefType()))
    GetActualTCPNum = Get-Method $robotType "GetActualTCPNum" ([Type[]]@([byte], [int].MakeByRefType()))
    GetActualWObjNum = Get-Method $robotType "GetActualWObjNum" ([Type[]]@([byte], [int].MakeByRefType()))
    GetCurToolCoord = Get-Method $robotType "GetCurToolCoord" ([Type[]]@($descPoseType.MakeByRefType()))
    GetCurWObjCoord = Get-Method $robotType "GetCurWObjCoord" ([Type[]]@($descPoseType.MakeByRefType()))
    GetRobotErrorCode = Get-Method $robotType "GetRobotErrorCode" ([Type[]]@([int].MakeByRefType(), [int].MakeByRefType()))
    GetSafetyStopState = Get-Method $robotType "GetSafetyStopState" ([Type[]]@([byte].MakeByRefType(), [byte].MakeByRefType()))
    IsInDragTeach = Get-Method $robotType "IsInDragTeach" ([Type[]]@([byte].MakeByRefType()))
}

$connectCode = [int]$methods.RPC.Invoke($robot, [object[]]@($Ip))
Write-JsonLine ([ordered]@{
    type = "session_start"
    capturedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
    ip = $Ip
    port = $Port
    durationSec = $DurationSec
    intervalMs = $IntervalMs
    dllPath = $DllPath
    outFile = $OutFile
    connectCode = $connectCode
    connected = ($connectCode -eq 0)
})

try {
    if ($connectCode -eq 0) {
        $sampleIndex = 0
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        while ($stopwatch.Elapsed.TotalSeconds -lt $DurationSec) {
            $sample = Capture-Sample -robot $robot -jointPosType $jointPosType -descPoseType $descPoseType -descTranType $descTranType -rpyType $rpyType -realtimeStateType $realtimeStateType -methods $methods -sampleIndex $sampleIndex
            Write-JsonLine $sample
            $sampleIndex++
            Start-Sleep -Milliseconds $IntervalMs
        }
    }
}
finally {
    $disconnectCode = -1
    try { $disconnectCode = [int]$methods.CloseRPC.Invoke($robot, [object[]]@()) } catch { $disconnectCode = -1 }

    Write-JsonLine ([ordered]@{
        type = "session_end"
        capturedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
        connectCode = $connectCode
        disconnectCode = $disconnectCode
    })
}

Write-Host "Record saved to $OutFile"
if ($connectCode -eq 0) { exit 0 }
exit 1
