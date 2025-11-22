# 빠른 설치 가이드

## 현재 상태
- 서버 실행 중: https://localhost:3000
- Manifest 위치: `C:\woo-work\workflow\Office-Add-in-samples\Samples\hello-world\powerpoint-hello-world\manifest-localhost.xml`

## 설치 방법 (추천)

### 방법 1: PowerPoint에서 직접 업로드

1. **PowerPoint 실행**

2. **새 프레젠테이션 만들기** (또는 기존 파일 열기)

3. **삽입 탭으로 이동**

4. **추가 기능** 클릭 (리본 메뉴 오른쪽에 있음)

5. **내 추가 기능** 클릭

6. **내 추가 기능 관리** (오른쪽 상단 드롭다운) > **내 추가 기능 업로드** 클릭

7. **파일 선택** 대화상자에서 다음 경로로 이동:
   ```
   C:\woo-work\workflow\Office-Add-in-samples\Samples\hello-world\powerpoint-hello-world
   ```

8. **manifest-localhost.xml** 파일 선택 > **업로드** 클릭

9. **홈 탭**으로 이동하면 **대체 텍스트 관리자** 버튼이 표시됩니다

## 사용 방법

1. **대체 텍스트 관리자** 버튼 클릭
2. 작업 창이 오른쪽에 열립니다
3. **슬라이드 요소 분석** 버튼을 클릭하여 시작

## 문제 해결

### "애드인을 로드할 수 없습니다" 오류가 뜨는 경우

1. 브라우저에서 https://localhost:3000/taskpane.html 을 열어보세요
2. 인증서 경고가 나오면 "고급" > "localhost로 계속" 클릭
3. 페이지가 보이면 PowerPoint를 완전히 종료하고 다시 시작
4. 애드인을 다시 업로드

### 서버가 실행되지 않는 경우

새 터미널/명령 프롬프트에서:
```bash
cd C:\woo-work\workflow\Office-Add-in-samples\Samples\hello-world\powerpoint-hello-world
http-server -S -C localhost.crt -K localhost.key --cors . -p 3000
```

### 애드인 제거 방법

1. PowerPoint 실행
2. 삽입 > 추가 기능 > 내 추가 기능
3. "대체 텍스트 관리자" 오른쪽 상단의 "..." 메뉴 > 제거

## 팁

- PowerPoint를 재시작하면 애드인이 자동으로 로드됩니다 (최초 1회 업로드 후)
- 코드를 수정하면 작업 창만 새로고침하면 됩니다
- Manifest를 수정한 경우에만 PowerPoint를 재시작해야 합니다
