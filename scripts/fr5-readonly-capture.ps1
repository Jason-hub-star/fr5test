[CmdletBinding()]
param(
    [string]$Ip = "",
    [int]$Port = 8080,
    [string]$DllPath = "",
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

    $searchRoots = @(
        (Join-Path $env:USERPROFILE "Desktop"),
        (Join-Path $env:USERPROFILE "Documents"),
        (Join-Path $env:USERPROFILE "Downloads")
    )

    foreach ($root in $searchRoots) {
        if (-not (Test-Path $root -PathType Container)) {
            continue
        }

        $found = Get-ChildItem -Path $root -Filter libfairino.dll -Recurse -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($null -ne $found) {
            return $found.FullName
        }
    }

    return $null
}

function Resolve-InputDefaults {
    if ([string]::IsNullOrWhiteSpace($script:Ip)) {
        $script:Ip = $env:FAIRINO_IP
    }
    if ([string]::IsNullOrWhiteSpace($script:Ip)) {
        $script:Ip = "192.168.58.2"
    }

    if (-not $PSBoundParameters.ContainsKey("Port") -and $env:FAIRINO_PORT) {
        $parsedPort = 0
        if ([int]::TryParse($env:FAIRINO_PORT, [ref]$parsedPort)) {
            $script:Port = $parsedPort
        }
    }

    if ([string]::IsNullOrWhiteSpace($script:DllPath)) {
        $script:DllPath = Find-DllPath
    }
    if ([string]::IsNullOrWhiteSpace($script:DllPath)) {
        throw "libfairino.dll not found automatically. Pass -DllPath, set FAIRINO_DLL_PATH, or place the DLL in a common location such as Desktop/Documents/Downloads."
    }
    if ((Split-Path $script:DllPath -Leaf) -ne "libfairino.dll" -and (Test-Path $script:DllPath -PathType Container)) {
        $script:DllPath = Join-Path $script:DllPath "libfairino.dll"
    }

    $script:DllPath = [System.IO.Path]::GetFullPath($script:DllPath)
    if (-not (Test-Path $script:DllPath -PathType Leaf)) {
        throw "libfairino.dll not found: $script:DllPath"
    }

    if ([string]::IsNullOrWhiteSpace($script:OutFile)) {
        $capturesDir = Join-Path $env:USERPROFILE "Desktop"
        $timestamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
        $script:OutFile = Join-Path $capturesDir "fr5-capture-$timestamp.json"
    }
    else {
        $outDir = Split-Path $script:OutFile -Parent
        if (-not [string]::IsNullOrWhiteSpace($outDir)) {
            New-Item -ItemType Directory -Force -Path $outDir | Out-Null
        }
    }

    $script:OutFile = [System.IO.Path]::GetFullPath($script:OutFile)
}

function New-CaptureDocument {
    [ordered]@{
        captureVersion = 1
        capturedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
        host = [ordered]@{
            computerName = $env:COMPUTERNAME
            userName = $env:USERNAME
            powershellEdition = $PSVersionTable.PSEdition
            powershellVersion = $PSVersionTable.PSVersion.ToString()
        }
        input = [ordered]@{
            ip = $script:Ip
            port = $script:Port
            dllPath = $script:DllPath
            outFile = $script:OutFile
        }
        steps = New-Object System.Collections.ArrayList
        data = [ordered]@{}
        summary = [ordered]@{
            connected = $false
            connectCode = $null
            disconnectCode = $null
            hasVersion = $false
            hasRealtimeState = $false
            hasFallbackJointPose = $false
            hasFallbackTcpPose = $false
            hasCoordContext = $false
            hasFaultState = $false
            hasDragTeachState = $false
        }
    }
}

function Add-Step($doc, [string]$name, [int]$code, [string]$message, $payload) {
    $null = $doc.steps.Add([ordered]@{
        name = $name
        code = $code
        success = ($code -eq 0)
        message = $message
        payload = $payload
    })
}

function Get-Method([Type]$type, [string]$name, [Type[]]$paramTypes) {
    $method = $type.GetMethod($name, $paramTypes)
    if ($null -eq $method) {
        throw "Method not found: $name"
    }
    $method
}

function New-DescPose([Type]$descPoseType, [Type]$descTranType, [Type]$rpyType) {
    $descPose = [Activator]::CreateInstance($descPoseType)
    $tran = [Activator]::CreateInstance($descTranType)
    $rpy = [Activator]::CreateInstance($rpyType)
    $descPoseType.GetField("tran").SetValue($descPose, $tran)
    $descPoseType.GetField("rpy").SetValue($descPose, $rpy)
    $descPose
}

function New-JointPos([Type]$jointPosType) {
    $jointPos = [Activator]::CreateInstance($jointPosType)
    $jointPosType.GetField("jPos").SetValue($jointPos, [double[]]@(0,0,0,0,0,0))
    $jointPos
}

function Convert-DescPose($descPose) {
    if ($null -eq $descPose) { return $null }
    $type = $descPose.GetType()
    $tran = $type.GetField("tran").GetValue($descPose)
    $rpy = $type.GetField("rpy").GetValue($descPose)
    [ordered]@{
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
    [ordered]@{
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
        load = [double]$type.GetField("load").GetValue($pkg)
        loadCog = @($type.GetField("loadCog").GetValue($pkg))
        gripperMotionDone = ([int]$type.GetField("gripper_motiondone").GetValue($pkg)) -ne 0
        gripperFaultId = [int]$type.GetField("gripper_fault_id").GetValue($pkg)
        gripperFault = [int]$type.GetField("gripper_fault").GetValue($pkg)
        gripperActive = [int]$type.GetField("gripper_active").GetValue($pkg)
        gripperPosition = [int]$type.GetField("gripper_position").GetValue($pkg)
        gripperSpeed = [int]$type.GetField("gripper_speed").GetValue($pkg)
        gripperCurrent = [int]$type.GetField("gripper_current").GetValue($pkg)
    }
}
Resolve-InputDefaults
$capture = New-CaptureDocument

$assembly = [System.Reflection.Assembly]::LoadFrom($DllPath)
$robotType = $assembly.GetType("fairino.Robot")
if ($null -eq $robotType) {
    throw "Type fairino.Robot not found in $DllPath"
}

$jointPosType = $assembly.GetType("fairino.JointPos")
$descPoseType = $assembly.GetType("fairino.DescPose")
$descTranType = $assembly.GetType("fairino.DescTran")
$rpyType = $assembly.GetType("fairino.Rpy")
$realtimeStateType = $assembly.GetType("fairino.ROBOT_STATE_PKG")
$robot = [Activator]::CreateInstance($robotType)

$connectMethod = Get-Method $robotType "RPC" ([Type[]]@([string]))
$disconnectMethod = Get-Method $robotType "CloseRPC" ([Type[]]@())
$getSdkVersionMethod = Get-Method $robotType "GetSDKVersion" ([Type[]]@([string].MakeByRefType()))
$getSoftwareVersionMethod = Get-Method $robotType "GetSoftwareVersion" ([Type[]]@([string].MakeByRefType(), [string].MakeByRefType(), [string].MakeByRefType()))
$getFirmwareVersionMethod = Get-Method $robotType "GetFirmwareVersion" ([Type[]]@([string].MakeByRefType(), [string].MakeByRefType(), [string].MakeByRefType(), [string].MakeByRefType(), [string].MakeByRefType(), [string].MakeByRefType(), [string].MakeByRefType(), [string].MakeByRefType()))
$getRealtimeStateMethod = Get-Method $robotType "GetRobotRealTimeState" ([Type[]]@($realtimeStateType.MakeByRefType()))
$getActualJointPosMethod = Get-Method $robotType "GetActualJointPosDegree" ([Type[]]@([byte], $jointPosType.MakeByRefType()))
$getActualTcpPoseMethod = Get-Method $robotType "GetActualTCPPose" ([Type[]]@([byte], $descPoseType.MakeByRefType()))
$getSafetyCodeMethod = Get-Method $robotType "GetSafetyCode" ([Type[]]@())
$getRealtimePeriodMethod = Get-Method $robotType "GetRobotRealtimeStateSamplePeriod" ([Type[]]@([int].MakeByRefType()))
$getActualTcpNumMethod = Get-Method $robotType "GetActualTCPNum" ([Type[]]@([byte], [int].MakeByRefType()))
$getActualWObjNumMethod = Get-Method $robotType "GetActualWObjNum" ([Type[]]@([byte], [int].MakeByRefType()))
$getCurToolCoordMethod = Get-Method $robotType "GetCurToolCoord" ([Type[]]@($descPoseType.MakeByRefType()))
$getCurWObjCoordMethod = Get-Method $robotType "GetCurWObjCoord" ([Type[]]@($descPoseType.MakeByRefType()))
$getRobotErrorCodeMethod = Get-Method $robotType "GetRobotErrorCode" ([Type[]]@([int].MakeByRefType(), [int].MakeByRefType()))
$getSafetyStopStateMethod = Get-Method $robotType "GetSafetyStopState" ([Type[]]@([byte].MakeByRefType(), [byte].MakeByRefType()))
$isInDragTeachMethod = Get-Method $robotType "IsInDragTeach" ([Type[]]@([byte].MakeByRefType()))

$connectCode = [int]$connectMethod.Invoke($robot, [object[]]@($Ip))
$capture.summary.connectCode = $connectCode
$capture.summary.connected = ($connectCode -eq 0)
$connectMessage = "CONNECT_FAIL"
if ($connectCode -eq 0) { $connectMessage = "OK" }
Add-Step -doc $capture -name "Connect" -code $connectCode -message $connectMessage -payload ([ordered]@{ ip = $Ip; port = $Port })

try {
    if ($connectCode -eq 0) {
        $sdkArgs = [object[]]@('')
        $sdkCode = [int]$getSdkVersionMethod.Invoke($robot, $sdkArgs)
        $swArgs = [object[]]@('', '', '')
        $swCode = [int]$getSoftwareVersionMethod.Invoke($robot, $swArgs)
        $fwArgs = [object[]]@('', '', '', '', '', '', '', '')
        $fwCode = [int]$getFirmwareVersionMethod.Invoke($robot, $fwArgs)
        $capture.data.version = [ordered]@{
            sdkCode = $sdkCode
            sdkVersion = [string]$sdkArgs[0]
            softwareCode = $swCode
            robotModel = [string]$swArgs[0]
            webVersion = [string]$swArgs[1]
            controllerVersion = [string]$swArgs[2]
            firmwareCode = $fwCode
            firmware = @($fwArgs | ForEach-Object { [string]$_ })
        }
        $capture.summary.hasVersion = ($sdkCode -eq 0)
        Add-Step -doc $capture -name "GetVersion" -code $sdkCode -message "Version captured" -payload $capture.data.version

        $pkg = [Activator]::CreateInstance($realtimeStateType)
        $realtimeArgs = [object[]]@($pkg)
        $realtimeCode = [int]$getRealtimeStateMethod.Invoke($robot, $realtimeArgs)
        $realtimePayload = $null
        if ($realtimeCode -eq 0) { $realtimePayload = Convert-RealtimeState $realtimeArgs[0] }
        $capture.data.realtimeState = $realtimePayload
        $capture.summary.hasRealtimeState = ($realtimeCode -eq 0)
        Add-Step -doc $capture -name "GetRobotRealTimeState" -code $realtimeCode -message "Realtime state read" -payload $realtimePayload

        $jointPos = New-JointPos $jointPosType
        $jointArgs = [object[]]@([byte]0, $jointPos)
        $jointCode = [int]$getActualJointPosMethod.Invoke($robot, $jointArgs)
        $jointPayload = $null
        if ($jointCode -eq 0) { $jointPayload = @($jointPosType.GetField("jPos").GetValue($jointArgs[1])) }
        $capture.data.actualJointPose = [ordered]@{ code = $jointCode; jointPosDeg = $jointPayload }
        $capture.summary.hasFallbackJointPose = ($jointCode -eq 0)
        Add-Step -doc $capture -name "GetActualJointPosDegree" -code $jointCode -message "Fallback joint pose read" -payload $capture.data.actualJointPose

        $tcpPose = New-DescPose $descPoseType $descTranType $rpyType
        $tcpArgs = [object[]]@([byte]0, $tcpPose)
        $tcpCode = [int]$getActualTcpPoseMethod.Invoke($robot, $tcpArgs)
        $tcpPayload = $null
        if ($tcpCode -eq 0) { $tcpPayload = Convert-DescPose $tcpArgs[1] }
        $capture.data.actualTcpPose = [ordered]@{ code = $tcpCode; tcpPose = $tcpPayload }
        $capture.summary.hasFallbackTcpPose = ($tcpCode -eq 0)
        Add-Step -doc $capture -name "GetActualTCPPose" -code $tcpCode -message "Fallback TCP pose read" -payload $capture.data.actualTcpPose
        $safetyCode = [int]$getSafetyCodeMethod.Invoke($robot, [object[]]@())
        $capture.data.safetyCode = [ordered]@{ code = 0; safetyCode = $safetyCode }
        Add-Step -doc $capture -name "GetSafetyCode" -code 0 -message "Safety code read" -payload $capture.data.safetyCode

        $periodArgs = [object[]]@(0)
        $periodCode = [int]$getRealtimePeriodMethod.Invoke($robot, $periodArgs)
        $periodValue = $null
        if ($periodCode -eq 0) { $periodValue = [int]$periodArgs[0] }
        $capture.data.realtimeSamplePeriod = [ordered]@{ code = $periodCode; periodMs = $periodValue }
        Add-Step -doc $capture -name "GetRobotRealtimeStateSamplePeriod" -code $periodCode -message "Realtime period read" -payload $capture.data.realtimeSamplePeriod

        $toolNumArgs = [object[]]@([byte]0, 0)
        $toolNumCode = [int]$getActualTcpNumMethod.Invoke($robot, $toolNumArgs)
        $wobjNumArgs = [object[]]@([byte]0, 0)
        $wobjNumCode = [int]$getActualWObjNumMethod.Invoke($robot, $wobjNumArgs)

        $toolCoord = New-DescPose $descPoseType $descTranType $rpyType
        $toolCoordArgs = [object[]]@($toolCoord)
        $toolCoordCode = [int]$getCurToolCoordMethod.Invoke($robot, $toolCoordArgs)

        $wobjCoord = New-DescPose $descPoseType $descTranType $rpyType
        $wobjCoordArgs = [object[]]@($wobjCoord)
        $wobjCoordCode = [int]$getCurWObjCoordMethod.Invoke($robot, $wobjCoordArgs)

        $toolIdValue = $null
        if ($toolNumCode -eq 0) { $toolIdValue = [int]$toolNumArgs[1] }
        $userIdValue = $null
        if ($wobjNumCode -eq 0) { $userIdValue = [int]$wobjNumArgs[1] }
        $toolCoordValue = $null
        if ($toolCoordCode -eq 0) { $toolCoordValue = Convert-DescPose $toolCoordArgs[0] }
        $wobjCoordValue = $null
        if ($wobjCoordCode -eq 0) { $wobjCoordValue = Convert-DescPose $wobjCoordArgs[0] }

        $capture.data.coordContext = [ordered]@{
            toolIdCode = $toolNumCode
            toolId = $toolIdValue
            userIdCode = $wobjNumCode
            userId = $userIdValue
            toolCoordCode = $toolCoordCode
            toolCoord = $toolCoordValue
            wobjCoordCode = $wobjCoordCode
            wobjCoord = $wobjCoordValue
        }
        $capture.summary.hasCoordContext = ($toolNumCode -eq 0 -and $toolCoordCode -eq 0)
        $coordStatusCode = 0
        foreach ($candidate in @($toolNumCode, $toolCoordCode, $wobjNumCode, $wobjCoordCode)) {
            if ($candidate -ne 0) { $coordStatusCode = $candidate; break }
        }
        Add-Step -doc $capture -name "ReadCoordContext" -code $coordStatusCode -message "Tool/User context read" -payload $capture.data.coordContext

        $faultArgs = [object[]]@(0, 0)
        $faultCode = [int]$getRobotErrorCodeMethod.Invoke($robot, $faultArgs)
        $safetyStopArgs = [object[]]@([byte]0, [byte]0)
        $safetyStopCode = [int]$getSafetyStopStateMethod.Invoke($robot, $safetyStopArgs)
        $dragArgs = [object[]]@([byte]0)
        $dragCode = [int]$isInDragTeachMethod.Invoke($robot, $dragArgs)

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

        $capture.data.controllerFault = [ordered]@{
            faultCode = $faultCode
            mainCode = $mainCode
            subCode = $subCode
            safetyStopCode = $safetyStopCode
            safetyStop0 = $safetyStop0
            safetyStop1 = $safetyStop1
            dragTeachCode = $dragCode
            isInDragTeach = $dragState
        }
        $capture.summary.hasFaultState = ($faultCode -eq 0)
        $capture.summary.hasDragTeachState = ($dragCode -eq 0)
        $faultStatusCode = 0
        foreach ($candidate in @($faultCode, $safetyStopCode, $dragCode)) {
            if ($candidate -ne 0) { $faultStatusCode = $candidate; break }
        }
        Add-Step -doc $capture -name "ReadFaultAndDragState" -code $faultStatusCode -message "Fault / safety / drag state read" -payload $capture.data.controllerFault
    }
}
finally {
    try {
        $disconnectCode = [int]$disconnectMethod.Invoke($robot, [object[]]@())
        $capture.summary.disconnectCode = $disconnectCode
        Add-Step -doc $capture -name "Disconnect" -code $disconnectCode -message "Disconnect attempted" -payload $null
    }
    catch {
        Add-Step -doc $capture -name "Disconnect" -code -1 -message $_.Exception.Message -payload $null
    }
}

$encoding = New-Object System.Text.UTF8Encoding($false)
$json = $capture | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText($OutFile, $json, $encoding)

Write-Host "Capture saved to $OutFile"
if ($capture.summary.connected) {
    exit 0
}
exit 1
