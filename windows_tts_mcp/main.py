"""
Windows TTS MCP Server
PowerShell 기반 Text-to-Speech 서버 for Claude Desktop
"""
import os
import sys
import io

# Windows 콘솔 UTF-8 인코딩 강제 설정
if sys.platform == "win32":
    os.environ['PYTHONIOENCODING'] = 'utf-8'
    # UTF-8 출력 강제 설정
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        # Python 3.6 이하 버전 호환성
        try:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
        except AttributeError:
            pass

from mcp.server.fastmcp import FastMCP
import subprocess
import platform
import threading
import time
from typing import Optional

# MCP 서버 생성
mcp = FastMCP("Windows TTS")

# 실행 중인 TTS 프로세스 관리
running_processes = []
process_lock = threading.Lock()

def safe_print(message: str):
    """안전한 print 함수 - 인코딩 오류 방지"""
    try:
        # 먼저 이모지와 특수 문자 제거
        import re
        # 이모지 패턴 제거
        emoji_pattern = re.compile("["
                                  u"\U0001F600-\U0001F64F"  # 감정 표현
                                  u"\U0001F300-\U0001F5FF"  # 기호와 그림 문자
                                  u"\U0001F680-\U0001F6FF"  # 교통과 지도 기호
                                  u"\U0001F1E0-\U0001F1FF"  # 국기
                                  u"\U0001F900-\U0001F9FF"  # 추가 기호
                                  u"\U0001FA70-\U0001FAFF"  # 추가 기호
                                  u"\u2600-\u26FF"          # 기타 기호
                                  u"\u2700-\u27BF"          # 장식 문자
                                  "]+", re.UNICODE)
        clean_message = emoji_pattern.sub('', message)
        
        print(clean_message)
    except UnicodeEncodeError:
        # 유니코드 문자를 안전한 형태로 변환
        try:
            safe_message = message.encode('ascii', errors='replace').decode('ascii')
            print(safe_message)
        except:
            print("[ENCODING ERROR] Unable to display message")
    except Exception as e:
        print(f"[PRINT ERROR] {str(e)}")

def powershell_tts(text: str, rate: int = 0, volume: int = 100) -> bool:
    """PowerShell을 사용한 TTS 실행"""
    process = None
    try:
        if platform.system() != "Windows":
            safe_print("[ERROR] Windows가 아닙니다")
            return False
        
        # 텍스트에서 작은따옴표 이스케이프 처리
        escaped_text = text.replace("'", "''")
        
        # PowerShell TTS 명령어
        cmd = [
            "powershell", "-Command",
            f"Add-Type -AssemblyName System.Speech; "
            f"$synth = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
            f"$synth.Rate = {rate}; "
            f"$synth.Volume = {volume}; "
            f"$synth.Speak('{escaped_text}'); "
            f"$synth.Dispose()"
        ]
        
        # 프로세스 시작
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        
        # 실행 중인 프로세스 목록에 추가
        with process_lock:
            running_processes.append(process)
        
        # 프로세스 완료 대기
        stdout, stderr = process.communicate(timeout=60)
        
        # 완료된 프로세스 목록에서 제거
        with process_lock:
            if process in running_processes:
                running_processes.remove(process)
        
        if process.returncode == 0:
            safe_print(f"[SUCCESS] TTS 완료: {text[:30]}...")
            return True
        else:
            safe_print(f"[ERROR] TTS 오류: {stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        safe_print("[WARNING] TTS 시간 초과")
        if process:
            process.kill()
            with process_lock:
                if process in running_processes:
                    running_processes.remove(process)
        return False
    except Exception as e:
        safe_print(f"[ERROR] TTS 예외: {e}")
        if process:
            try:
                process.kill()
                with process_lock:
                    if process in running_processes:
                        running_processes.remove(process)
            except:
                pass
        return False

def split_text_for_tts(text: str, max_length: int = 500) -> list:
    """텍스트를 TTS용으로 적절히 분할"""
    if len(text) <= max_length:
        return [text]
    
    # 문장 단위로 분할 시도
    import re
    sentences = re.split(r'[.!?。！？]\s*', text)
    
    chunks = []
    current_chunk = ""
    
    for sentence in sentences:
        # 문장이 너무 긴 경우 더 작게 분할
        if len(sentence) > max_length:
            # 쉼표나 기타 구두점으로 분할
            sub_parts = re.split(r'[,;:\s]\s*', sentence)
            for part in sub_parts:
                if len(current_chunk + part) <= max_length:
                    current_chunk += part + " "
                else:
                    if current_chunk.strip():
                        chunks.append(current_chunk.strip())
                    current_chunk = part + " "
        else:
            # 현재 청크에 문장을 추가할 수 있는지 확인
            if len(current_chunk + sentence) <= max_length:
                current_chunk += sentence + ". "
            else:
                if current_chunk.strip():
                    chunks.append(current_chunk.strip())
                current_chunk = sentence + ". "
    
    # 마지막 청크 추가
    if current_chunk.strip():
        chunks.append(current_chunk.strip())
    
    return chunks

@mcp.tool()
def speak(text: str) -> str:
    """텍스트를 음성으로 읽어줍니다"""
    try:
        # 텍스트 분할
        text_chunks = split_text_for_tts(text, 500)
        total_chunks = len(text_chunks)
        
        def _speak_thread():
            for i, chunk in enumerate(text_chunks, 1):
                safe_print(f"[TTS] {i}/{total_chunks} 부분 재생 중: {chunk[:50]}...")
                success = powershell_tts(chunk)
                if not success:
                    safe_print(f"[RETRY] TTS 재시도: {chunk[:30]}...")
                    # 한 번 더 시도
                    powershell_tts(chunk)
                
                # 각 청크 사이에 짧은 간격
                if i < total_chunks:
                    time.sleep(0.5)
        
        # 백그라운드에서 실행
        thread = threading.Thread(target=_speak_thread, daemon=True)
        thread.start()
        
        if total_chunks > 1:
            return f"[START] 음성 재생 시작 ({total_chunks}개 부분으로 분할): '{text[:50]}...'"
        else:
            return f"[START] 음성 재생 시작: '{text[:50]}...'"
        
    except Exception as e:
        return f"[ERROR] 음성 재생 오류: {str(e)}"

@mcp.tool()
def speak_fast(text: str) -> str:
    """텍스트를 빠른 속도로 읽어줍니다"""
    try:
        # 텍스트 분할 (빠른 재생은 조금 더 짧게)
        text_chunks = split_text_for_tts(text, 400)
        total_chunks = len(text_chunks)
        
        def _speak_fast():
            for i, chunk in enumerate(text_chunks, 1):
                safe_print(f"[FAST TTS] {i}/{total_chunks} 부분 재생 중: {chunk[:50]}...")
                powershell_tts(chunk, rate=3, volume=100)  # 빠른 속도
                if i < total_chunks:
                    time.sleep(0.3)  # 빠른 재생은 간격도 짧게
        
        thread = threading.Thread(target=_speak_fast, daemon=True)
        thread.start()
        
        if total_chunks > 1:
            return f"[FAST] 빠른 재생 시작 ({total_chunks}개 부분): '{text[:50]}...'"
        else:
            return f"[FAST] 빠른 재생 시작: '{text[:50]}...'"
        
    except Exception as e:
        return f"[ERROR] 빠른 재생 오류: {str(e)}"

@mcp.tool()
def speak_slow(text: str) -> str:
    """텍스트를 천천히 읽어줍니다"""
    try:
        # 텍스트 분할
        text_chunks = split_text_for_tts(text, 400)
        total_chunks = len(text_chunks)
        
        def _speak_slow():
            for i, chunk in enumerate(text_chunks, 1):
                safe_print(f"[SLOW TTS] {i}/{total_chunks} 부분 재생 중: {chunk[:50]}...")
                powershell_tts(chunk, rate=-3, volume=100)  # 느린 속도
                if i < total_chunks:
                    time.sleep(0.8)  # 느린 재생은 간격을 더 길게
        
        thread = threading.Thread(target=_speak_slow, daemon=True)
        thread.start()
        
        if total_chunks > 1:
            return f"[SLOW] 천천히 재생 시작 ({total_chunks}개 부분): '{text[:50]}...'"
        else:
            return f"[SLOW] 천천히 재생 시작: '{text[:50]}...'"
        
    except Exception as e:
        return f"[ERROR] 천천히 재생 오류: {str(e)}"

@mcp.tool()
def speak_quiet(text: str) -> str:
    """텍스트를 작은 볼륨으로 읽어줍니다"""
    try:
        # 텍스트 분할
        text_chunks = split_text_for_tts(text, 400)
        total_chunks = len(text_chunks)
        
        def _speak_quiet():
            for i, chunk in enumerate(text_chunks, 1):
                safe_print(f"[QUIET TTS] {i}/{total_chunks} 부분 재생 중: {chunk[:50]}...")
                powershell_tts(chunk, rate=0, volume=50)  # 작은 볼륨
                if i < total_chunks:
                    time.sleep(0.5)
        
        thread = threading.Thread(target=_speak_quiet, daemon=True)
        thread.start()
        
        if total_chunks > 1:
            return f"[QUIET] 작은 볼륨 재생 시작 ({total_chunks}개 부분): '{text[:50]}...'"
        else:
            return f"[QUIET] 작은 볼륨 재생 시작: '{text[:50]}...'"
        
    except Exception as e:
        return f"[ERROR] 작은 볼륨 재생 오류: {str(e)}"

@mcp.tool()
def speak_short(text: str) -> str:
    """짧은 텍스트를 즉시 읽어줍니다 (100자 이하)"""
    try:
        if len(text) > 100:
            return "[ERROR] 텍스트가 너무 깁니다. speak를 사용하세요."
        
        def _speak_short():
            powershell_tts(text)
        
        thread = threading.Thread(target=_speak_short, daemon=True)
        thread.start()
        
        return f"[SHORT] 짧은 텍스트 재생: '{text}'"
        
    except Exception as e:
        return f"[ERROR] 짧은 텍스트 재생 오류: {str(e)}"

@mcp.tool()
def stop_speech() -> str:
    """현재 재생 중인 모든 음성을 중지합니다"""
    try:
        stopped_count = 0
        
        with process_lock:
            # 실행 중인 모든 TTS 프로세스 종료
            for process in running_processes[:]:  # 복사본으로 순회
                try:
                    if process.poll() is None:  # 아직 실행 중인 프로세스
                        process.terminate()
                        time.sleep(0.1)
                        if process.poll() is None:  # 여전히 실행 중이면 강제 종료
                            process.kill()
                        stopped_count += 1
                    running_processes.remove(process)
                except Exception as e:
                    safe_print(f"프로세스 종료 오류: {e}")
            
            running_processes.clear()
        
        # PowerShell 프로세스도 강제 종료
        try:
            # Windows에서 모든 PowerShell TTS 프로세스 찾아서 종료
            subprocess.run([
                "powershell", "-Command",
                "Get-Process | Where-Object {$_.ProcessName -eq 'powershell' -and $_.CommandLine -like '*Speech*'} | Stop-Process -Force"
            ], capture_output=True, timeout=5)
        except:
            pass
        
        if stopped_count > 0:
            return f"[STOP] {stopped_count}개의 음성 재생을 중지했습니다"
        else:
            return "[INFO] 현재 재생 중인 음성이 없습니다"
            
    except Exception as e:
        return f"[ERROR] 음성 중지 오류: {str(e)}"

@mcp.tool()
def kill_all_tts() -> str:
    """모든 TTS 관련 프로세스를 강제 종료합니다"""
    try:
        # 1. 관리 중인 프로세스 종료
        with process_lock:
            for process in running_processes[:]:
                try:
                    process.kill()
                    running_processes.remove(process)
                except:
                    pass
            running_processes.clear()
        
        # 2. 시스템의 모든 PowerShell TTS 프로세스 강제 종료
        try:
            subprocess.run([
                "taskkill", "/F", "/IM", "powershell.exe"
            ], capture_output=True, timeout=10)
        except:
            pass
        
        # 3. Speech 관련 프로세스 정리
        try:
            subprocess.run([
                "powershell", "-Command",
                "Get-Process | Where-Object {$_.ProcessName -like '*speech*' -or $_.CommandLine -like '*Speech*'} | Stop-Process -Force"
            ], capture_output=True, timeout=5)
        except:
            pass
        
        return "[KILL] 모든 TTS 프로세스를 강제 종료했습니다"
        
    except Exception as e:
        return f"[ERROR] 강제 종료 오류: {str(e)}"

@mcp.tool()
def get_tts_status() -> str:
    """현재 TTS 상태를 확인합니다"""
    try:
        active_count = 0
        
        with process_lock:
            # 실행 중인 프로세스 확인
            for process in running_processes[:]:
                if process.poll() is None:  # 아직 실행 중
                    active_count += 1
                else:  # 완료된 프로세스는 목록에서 제거
                    running_processes.remove(process)
        
        if active_count > 0:
            return f"[ACTIVE] 현재 {active_count}개의 음성이 재생 중입니다"
        else:
            return "[IDLE] 현재 재생 중인 음성이 없습니다"
            
    except Exception as e:
        return f"[ERROR] 상태 확인 오류: {str(e)}"

@mcp.tool()
def emergency_silence() -> str:
    """긴급 음소거 - 모든 오디오 중지 + 시스템 음소거"""
    try:
        # 1. TTS 프로세스 모두 종료
        kill_all_tts()
        
        # 2. 시스템 음소거
        try:
            subprocess.run([
                "powershell", "-Command",
                "(New-Object -ComObject WScript.Shell).SendKeys([char]173)"  # 음소거 키
            ], capture_output=True, timeout=3)
        except:
            pass
        
        # 3. 대안: nircmd 사용 (설치되어 있다면)
        try:
            subprocess.run(["nircmd", "mutesysvolume", "1"], capture_output=True, timeout=3)
        except:
            pass
        
        return "[EMERGENCY] 긴급 음소거 실행 완료 (TTS 중지 + 시스템 음소거)"
        
    except Exception as e:
        return f"[ERROR] 긴급 음소거 오류: {str(e)}"

@mcp.tool()
def test_tts() -> str:
    """TTS 시스템 테스트"""
    try:
        if platform.system() != "Windows":
            return "[ERROR] 이 TTS 서버는 Windows에서만 작동합니다"
        
        test_text = "Windows TTS MCP 서버 테스트입니다"
        
        def _test():
            success = powershell_tts(test_text)
            if success:
                safe_print("[SUCCESS] TTS 테스트 성공")
            else:
                safe_print("[ERROR] TTS 테스트 실패")
        
        thread = threading.Thread(target=_test, daemon=True)
        thread.start()
        
        return "[TEST] TTS 테스트를 시작했습니다"
        
    except Exception as e:
        return f"[ERROR] TTS 테스트 오류: {str(e)}"

@mcp.resource("tts://help")
def get_help() -> str:
    """TTS 서버 사용 방법"""
    return """
Windows TTS MCP Server (PowerShell 기반)

음성 재생:
- "이 텍스트를 읽어줘" → speak 사용 (긴 텍스트 자동 분할)
- "빠르게 읽어줘" → speak_fast 사용
- "천천히 읽어줘" → speak_slow 사용
- "작게 읽어줘" → speak_quiet 사용

음성 제어:
- "음성 중지해줘" → stop_speech 사용
- "모든 TTS 강제 종료" → kill_all_tts 사용
- "TTS 상태 확인" → get_tts_status 사용
- "긴급 음소거" → emergency_silence 사용

전체 도구 목록:
재생: speak, speak_fast, speak_slow, speak_quiet, speak_short
제어: stop_speech, kill_all_tts, get_tts_status, emergency_silence
기타: test_tts

특징:
- Windows PowerShell 기반으로 안정적
- 백그라운드 재생으로 빠른 응답
- 긴 텍스트 자동 분할 재생 (500자 단위)
- 문장 단위 지능형 분할로 자연스러운 재생
- 실행 중인 음성 추적 및 제어 가능
- 강제 중지 및 긴급 음소거 지원
- 길이 제한 없이 긴 텍스트 지원

주의사항:
- Windows에서만 작동
- emergency_silence는 시스템 전체 음소거
- 긴 텍스트는 자동으로 여러 부분으로 나누어 재생

패키지 정보:
- 실행: uvx windows-tts-mcp
- 개발: uvx --from . tts-dev
"""

def main():
    """메인 서버 실행"""
    # 콘솔 인코딩 문제 해결을 위해 이모지 대신 텍스트 사용
    safe_print("Windows TTS MCP Server 시작...")
    safe_print("패키지 버전: v1.0.0")
    
    # 시작 테스트
    if platform.system() == "Windows":
        def _startup_test():
            time.sleep(1)  # 서버 시작 대기
            powershell_tts("Windows TTS MCP 서버가 시작되었습니다")
        
        startup_thread = threading.Thread(target=_startup_test, daemon=True)
        startup_thread.start()
    else:
        safe_print("[WARNING] Windows가 아닙니다. 일부 기능이 제한됩니다.")
    
    # MCP 서버 실행
    mcp.run()

def dev_main():
    """개발 모드 실행"""
    safe_print("[DEV] Windows TTS MCP 개발 모드 시작...")
    safe_print("[DEV] FastMCP 개발 서버로 실행합니다...")
    
    try:
        # 개발 모드에서는 디버그 정보 출력
        safe_print(f"[DEV] 현재 디렉터리: {os.getcwd()}")
        safe_print(f"[DEV] Python 경로: {sys.executable}")
        safe_print(f"[DEV] 스크립트 파일: {__file__}")
        
        # FastMCP dev 모드 활성화
        os.environ["FASTMCP_DEV"] = "1"
        safe_print("[DEV] FastMCP 개발 모드 활성화됨")
        
        # 개발 모드로 메인 서버 실행 (디버그 출력 포함)
        main()
        
    except KeyboardInterrupt:
        safe_print("\n[EXIT] 개발 모드 종료")
    except Exception as e:
        safe_print(f"[ERROR] 개발 모드 오류: {e}")
        # 스택 트레이스 출력
        import traceback
        safe_print(traceback.format_exc())

if __name__ == "__main__":
    main()
