# ed_excel

# 현재 적용된 개발방식
매크로포함 엑셀로 개발, 확장자 .xlsm 방식
.bas 파일이 엑셀 내부에 포함됨으로 github 에서 안보임
## 불러오기
1. 해당 파일을 불러오면, 매크로 모듈까지 같이 불러옴
## 버전 관리 github
1. vba module 제작
1. 엑셀의 저장버튼 누름
1. git add, commit, push 
#. git add 시에 엑셀 파일이 열려 있으면 저장 안됨(임시파일 접근 에러 발생)



- 아래는 모듈식 개발 방식임, xlsm 방식아닌 경우
# 모듈식 개발방식 
## 모듈 설치
1. copy this on local pc
1. open Excel
1. open VBA 
1. 가져오기 모듈 선택
1. *.bas 선택
## 모듈 저장 on github
1. vba module 제작
1. save 
1. 내보내기 모듈
1. *.bas  로 저장
1. git add, commit, push on local pc 
#. git add 시에 엑셀 파일이 열려 있으면 저장 안됨(임시파일 접근 에러 발생)

