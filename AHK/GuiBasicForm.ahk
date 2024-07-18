#Persistent  ; 스크립트를 지속적으로 실행
#NoEnv  ; 환경변수 사용 안 함
#Warn  ; 경고 활성화
#SingleInstance force  ; 중복 실행 방지
SendMode Input  ; 입력 모드 설정


;; GUI 프로그램 시작 화면 창 구성 -----------------------------------------------------------------------------------

Gui, +LastFound
Gui, +MinimizeBox  ; 최소화 버튼 활성화

hGui := WinExist()

manual =
(
단축키 목록

무엇이든 상상할 수 있는 사람은
무엇이든 만들어 낼 수 있다
				- 엘런 튜링





■ F12 : 종료

                                    2024.07.18 version 1.0.0
)

Gui, Add, Text, , %manual%
Gui, Add, Button, x136 y150 w70 h40 gMinimizeBtn Default, 확인  ; 최소화 버튼 추가. 확인 버튼에 기본 포커스
Gui, Add, Button, x219 y150 w70 h40 gExitBtn, 종료
Gui, Show, w310 h200, 꼬북이A

Menu, Tray, Add, 첫 화면, RestoreGui
Menu, Tray, Add, 종료 (F12), ExitScript
Menu, Tray, Default, 첫 화면
Menu, Tray, Tip, 꼬북이  ; 트레이 아이콘에 툴팁 설정

return

;; GUI 기본 버튼 및 명령----------------------------------------------------------------------------------
MinimizeBtn:  ; 최소화 버튼 라벨
    WinHide, ahk_id %hGui%
return

RestoreGui: ; 첫 화면 라벨
    WinShow, ahk_id %hGui%
    WinActivate, ahk_id %hGui%
	SendInput {Esc}
return

ExitScript:
    ExitApp
return

ExitBtn:
    ExitApp  ; 종료 버튼 클릭 시 종료
return

GuiClose:
    ExitApp  ; GUI 창 닫기 버튼 클릭 시 종료
return


;; HOT KEY 명령 설정 -----------------------------------------------------------------------------------

; 종료
F12::
	MsgBox, 종료되었습니다
	ExitApp
