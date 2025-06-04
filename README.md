# Windows TTS MCP Server

Claude Desktop용 Windows PowerShell 기반 Text-to-Speech 서버

## 🎯 개요

Windows TTS MCP Server는 Claude Desktop에서 텍스트를 음성으로 변환해주는 MCP(Model Context Protocol) 서버입니다. Windows의 내장 Speech API를 PowerShell을 통해 활용하여 안정적이고 빠른 TTS 기능을 제공합니다.

## ✨ 주요 특징

- 🎵 **안정적인 TTS**: Windows PowerShell 기반으로 스레드 안전성 확보
- ⚡ **빠른 응답**: 백그라운드 재생으로 즉시 응답
- 🎛️ **다양한 옵션**: 속도, 볼륨 조절 가능
- ⏹️ **완전한 제어**: 재생 중지, 강제 종료, 상태 확인
- 🚨 **긴급 기능**: 시스템 전체 음소거
- 📦 **간편한 설치**: uvx로 원클릭 설치 및 실행

## 🚀 설치 및 실행

### uvx로 실행 (권장)

```bash
# 패키지 실행
uvx windows-tts-mcp

# 개발 모드
uvx --from . tts-dev

# GitHub에서 직접 실행 (향후)
uvx --from git+https://github.com/yourusername/windows-tts-mcp windows-tts-mcp
```

### 로컬 개발

```bash
# 저장소 클론
git clone https://github.com/yourusername/windows-tts-mcp
cd windows-tts-mcp

# 로컬에서 실행
uvx --from . windows-tts-mcp

# 개발 모드
uvx --from . tts-dev
```

## 🎛️ Claude Desktop 설정

Claude Desktop의 설정 파일에 다음을 추가하세요:

```json
{
  "mcpServers": {
    "windows-tts": {
      "command": "uvx",
      "args": ["--from", "D:\\MCP\\windows_tts_mcp", "windows-tts-mcp"]
    }
  }
}
```

설정 파일 위치:
- **Windows**: `%APPDATA%\\Claude\\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Linux**: `~/.config/claude/claude_desktop_config.json`

## 🎵 사용법

Claude Desktop에서 다음과 같이 요청하면 됩니다:

### 📢 기본 음성 재생
- "이 텍스트를 읽어줘"
- "다음 내용을 음성으로 들려줘"
- "답변을 소리로 들려줘"

### ⚡ 속도 조절
- "빠르게 읽어줘: [텍스트]"
- "천천히 읽어줘: [텍스트]"
- "작은 소리로 읽어줘: [텍스트]"

### ⏹️ 재생 제어
- "음성 중지해줘"
- "TTS 상태 확인해줘"
- "모든 음성 강제 종료해줘"
- "긴급 음소거해줘" (시스템 전체 음소거)

## 🛠️ 도구 목록

| 도구 | 설명 | 사용 예시 |
|------|------|----------|
| `speak` | 기본 음성 재생 | "이 텍스트 읽어줘" |
| `speak_fast` | 빠른 속도 재생 | "빠르게 읽어줘" |
| `speak_slow` | 천천히 재생 | "천천히 읽어줘" |
| `speak_quiet` | 작은 볼륨 재생 | "작게 읽어줘" |
| `speak_short` | 짧은 텍스트 (100자 이하) | "짧게 말해줘" |
| `stop_speech` | 음성 재생 중지 | "음성 멈춰줘" |
| `kill_all_tts` | 모든 TTS 강제 종료 | "모든 음성 종료" |
| `get_tts_status` | 재생 상태 확인 | "음성 상태 확인" |
| `emergency_silence` | 긴급 음소거 | "긴급 음소거" |
| `test_tts` | 시스템 테스트 | "TTS 테스트" |

## 🔧 기술적 세부사항

### 시스템 요구사항
- **운영체제**: Windows 10/11
- **Python**: 3.11 이상
- **PowerShell**: 5.0 이상 (Windows 기본 포함)
- **기타**: Windows Speech Platform (기본 포함)

### 아키텍처
```
Claude Desktop
    ↓ (MCP Protocol)
Windows TTS MCP Server
    ↓ (subprocess)
PowerShell
    ↓ (System.Speech)
Windows Speech API
    ↓
Audio Output
```

### 특징
- **스레드 안전성**: 각 TTS 요청을 별도 스레드에서 처리
- **프로세스 관리**: 실행 중인 모든 TTS 프로세스 추적 및 제어
- **오류 처리**: 타임아웃, 예외 상황에 대한 견고한 처리
- **리소스 정리**: 완료된 프로세스 자동 정리

## 🐛 문제 해결

### 음성이 나오지 않는 경우
1. **Windows 오디오 확인**
   ```bash
   # PowerShell에서 직접 테스트
   Add-Type -AssemblyName System.Speech
   $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
   $synth.Speak("테스트")
   ```

2. **볼륨 설정 확인**
   - Windows 볼륨 믹서에서 PowerShell 볼륨 확인
   - 시스템 기본 오디오 장치 설정 확인

3. **Speech Platform 확인**
   - Windows 설정 → 시간 및 언어 → 음성
   - 음성 인식 언어 설정 확인

### 프로세스가 남아있는 경우
```bash
# 강제 정리
uvx windows-tts-mcp
# Claude에서 "모든 TTS 강제 종료해줘" 실행
```

## 📝 개발

### 프로젝트 구조
```
windows-tts-mcp/
├── pyproject.toml          # 패키지 설정
├── README.md               # 이 파일
├── windows_tts_mcp/        # 메인 패키지
│   ├── __init__.py
│   └── main.py            # TTS 서버 코드
├── .gitignore
└── uv.lock
```

### 개발 환경 설정
```bash
# 저장소 클론
git clone https://github.com/yourusername/windows-tts-mcp
cd windows-tts-mcp

# uv로 의존성 설치
uv sync

# 개발 모드 실행
uvx --from . tts-dev
```

### 기여하기
1. 이 저장소를 포크하세요
2. 새 브랜치를 만드세요 (`git checkout -b feature/amazing-feature`)
3. 변경사항을 커밋하세요 (`git commit -m 'Add amazing feature'`)
4. 브랜치에 푸시하세요 (`git push origin feature/amazing-feature`)
5. Pull Request를 열어주세요

## 📄 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 `LICENSE` 파일을 참조하세요.

## 🙏 감사의 말

- [Anthropic](https://www.anthropic.com/) - Claude와 MCP 개발
- [Model Context Protocol](https://modelcontextprotocol.io/) - MCP 표준 제공
- [uv](https://github.com/astral-sh/uv) - 현대적인 Python 패키지 관리

## 🔗 관련 링크

- [Model Context Protocol 공식 문서](https://modelcontextprotocol.io/)
- [Claude Desktop 다운로드](https://claude.ai/download)
- [uv 공식 문서](https://docs.astral.sh/uv/)

---

**Made with ❤️ for Claude Desktop users**
