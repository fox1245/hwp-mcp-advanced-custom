# WSL2-Windows SSE 연동 설정

## 개요

Windows에서 HWP COM 객체를 제어하는 MCP 서버를 SSE 트랜스포트로 실행하고,
WSL2의 Claude Code에서 네트워크로 연결하는 구성입니다.

## 구조

```
WSL2 (Claude Code)
    | SSE (http://<windows-host-ip>:8765/sse)
Windows (advanced_hwp_server.py — SSE 모드, 0.0.0.0:8765)
    | COM
한글 프로그램
```

## 변경 사항

### advanced_hwp_server.py

- `FastMCP` 생성자에 `host="0.0.0.0"`, `port=8765` 추가
- `mcp.run()` → `mcp.run(transport="sse")` 변경

### requirements.txt

- `pythoncom` 주석 처리 (`pywin32`에 포함되어 별도 설치 불필요)

## Windows 측 실행

```powershell
cd C:\pickcare\mcp\hwp-mcp-advanced-custom
.\venv\Scripts\python.exe advanced_hwp_server.py
```

## WSL2 측 MCP 등록

```bash
# Windows 호스트 IP 확인
cat /etc/resolv.conf | grep nameserver

# MCP 서버 등록
claude mcp add advanced-hwp --transport sse http://<windows-host-ip>:8765/sse
```

## 필수 조건

- Windows에 한글(HWP) 프로그램 설치
- Windows Python 3.11+ 가상환경 (`venv/`)에 종속성 설치 완료
- Windows 방화벽에서 포트 8765 허용
- 한글 사용 시에만 서버 실행 필요
