# FR5 Read-Only Capture

FAIRINO FR5 controller state를 `read-only`로 캡처해 JSON으로 저장하는 PowerShell 유틸입니다.

이 저장소의 목표는 아래 순서만 안전하게 수행하는 것입니다.

1. `Connect`
2. `GetVersion`
3. `GetRobotRealTimeState`
4. `GetActualJointPosDegree`
5. `GetActualTCPPose`
6. `GetSafetyCode`
7. `GetRobotRealtimeStateSamplePeriod`
8. `GetActualTCPNum`
9. `GetActualWObjNum`
10. `GetCurToolCoord`
11. `GetCurWObjCoord`
12. `GetRobotErrorCode`
13. `GetSafetyStopState`
14. `IsInDragTeach`
15. `Disconnect`

`MoveJ`, `MoveL`, jog 계열 명령은 포함하지 않습니다.

## Safety

- 첫 세션은 반드시 `read-only capture`만 수행하세요.
- pendant / e-stop / 현장 승인 없이 모션 명령을 보내지 마세요.
- 유선 LAN 연결을 권장합니다.

## Requirements

- `Windows PowerShell 5.1` 권장
- 동료 PC에 접근 가능한 `libfairino.dll`
- 로봇 컨트롤러와 같은 subnet

PowerShell 7/.NET Core에서는 SDK 호환성 문제가 있을 수 있습니다.

## Usage

가장 쉬운 사용법은 아래 한 줄입니다.

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\fr5-readonly-capture.ps1 -Ip 192.168.58.2
```

이 경우 스크립트는 아래를 자동으로 시도합니다.

- `libfairino.dll` 자동 탐색
  - 환경변수 `FAIRINO_DLL_PATH`
  - 현재 폴더
  - 현재 폴더의 `Assets\Plugins\Fairino\`
  - 바탕화면 / 문서 / 다운로드
- 저장 경로 자동 설정
  - 바탕화면 `fr5-capture-YYYYMMDD-HHMMSS.json`

5분 동안 연속 기록하려면 아래 한 줄을 쓰세요.

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\fr5-readonly-record.ps1 -Ip 192.168.58.2
```

이 경우 스크립트는:

- `libfairino.dll` 자동 탐색
- 바탕화면에 `fr5-record-YYYYMMDD-HHMMSS.jsonl` 생성
- 5분 동안 기록
- 5분 후 자동 종료

즉, 네. 5분이 지나면 스크립트는 자동으로 끝나고, 저장 파일은 바탕화면에 보입니다.
기록 중에도 파일은 이미 생성되어 있고 계속 커집니다.

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\fr5-readonly-capture.ps1 `
  -Ip 192.168.58.2 `
  -Port 8080 `
  -DllPath "C:\path\to\libfairino.dll"
```

환경 변수로도 줄 수 있습니다.

```powershell
$env:FAIRINO_IP = "192.168.58.2"
$env:FAIRINO_PORT = "8080"
$env:FAIRINO_DLL_PATH = "C:\path\to\libfairino.dll"
powershell.exe -ExecutionPolicy Bypass -File .\scripts\fr5-readonly-capture.ps1
```

결과는 기본적으로 바탕화면에 JSON으로 저장됩니다.

바탕화면에 바로 저장하려면 `-OutFile`을 지정하세요.

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\fr5-readonly-capture.ps1 `
  -Ip 192.168.58.2 `
  -Port 8080 `
  -DllPath "C:\path\to\libfairino.dll" `
  -OutFile "$env:USERPROFILE\Desktop\fr5-capture.json"
```

## Output

JSON에는 아래가 포함됩니다.

- 연결 성공/실패
- SDK / software / firmware version
- realtime joint / tcp / tool / user / enable / mode / emergency stop
- fallback joint / tcp getter 값
- tool coordinate / wobj coordinate
- fault / safety
- drag teach 상태

`fr5-readonly-record.ps1`는 `jsonl` 형식으로 저장합니다.

- 한 줄당 한 샘플
- 기록 중에도 파일이 보임
- 중간에 끊겨도 앞부분 데이터 보존

## Notes

- `RPC`는 공식 SDK 기준으로 `ip`만 받습니다.
- `port` 인자는 메타데이터로만 저장되며 SDK 호출에는 직접 쓰지 않습니다.
- 캡처 결과는 원본 evidence입니다. 이 값을 곧바로 SSOT로 취급하지 말고, 검토 후 정규화 문서에 반영하세요.
