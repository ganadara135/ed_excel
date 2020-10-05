# 사용법 설명
1. 다운받은 엑셀파일(ed_excel.xlsm)과 원본데이터(rawData.csv)을 같은 폴더에 위치 
1. 엑셀파일을 실행시킨다
1. 실행된 화면에서 업데이트 버튼을 누른다
![업데이트](https://user-images.githubusercontent.com/24896007/95046985-34bbc180-0720-11eb-8be7-e2bf994038bf.PNG)
1. 엑셀시트는 모두 4개가 있음
![시트4개](https://user-images.githubusercontent.com/24896007/95046958-2f5e7700-0720-11eb-8703-c89ceaa0307e.PNG)
1. 상단의 보기 -> 매크로 실행
![매크로](https://user-images.githubusercontent.com/24896007/95046959-2ff70d80-0720-11eb-8fc5-435ddac892f2.PNG)
1. 매크로중에 main 실행
![main](https://user-images.githubusercontent.com/24896007/95046960-2ff70d80-0720-11eb-99f1-b36338c35832.PNG)
1. rawData 를 불러오기
![불러오기](https://user-images.githubusercontent.com/24896007/95046961-308fa400-0720-11eb-9a3f-9c0d251df966.PNG)
1. 엑셀화면에 사용자 컨트롤 창이 뜸
![컨트롤창](https://user-images.githubusercontent.com/24896007/95046963-31283a80-0720-11eb-95c7-476ca9c62424.PNG)
1. 진행 메시지창이 뜨면 '예'를 선택
![메시지창](https://user-images.githubusercontent.com/24896007/95046965-31283a80-0720-11eb-8632-ea6126642465.PNG)
1. 차트에서 시작점과 종료점 선택하라는 메시지창
![메시지창2](https://user-images.githubusercontent.com/24896007/95046967-31c0d100-0720-11eb-937d-0fdc53cd1445.PNG)
1. 화면상 나타난 차트에서 시작점과 종료점을 더블클릭해서 선택
![더블클릭](https://user-images.githubusercontent.com/24896007/95046968-31c0d100-0720-11eb-879b-b2bd08b6ad96.PNG)
1. 엑셀시트 상단에 시작점과 종료점이 나타남
![시작종료점](https://user-images.githubusercontent.com/24896007/95046970-32596780-0720-11eb-9291-c988e2537754.PNG)
1. 사용자 컨트롤 창에서 변곡점 계산하기를 선택하면 메시지창 나타남
![변곡점계산](https://user-images.githubusercontent.com/24896007/95046972-32596780-0720-11eb-8cb1-9688051fc86e.PNG)
1. 엑셀시트에 변곡점 메시지가 나타남
![메시지창3](https://user-images.githubusercontent.com/24896007/95046973-32f1fe00-0720-11eb-9e2e-21d303632c1d.PNG)
1. 사용자 컨트롤 창에서 플마고저계산하기 선택하면 다음과 같은 이미지 나타남
![로딩이미지](https://user-images.githubusercontent.com/24896007/95046974-32f1fe00-0720-11eb-81c8-9663a2171204.PNG)
1. 사용자 컨트롤 창에서 로딩률계산 선택하면 산출서 서식이 나타남
![로딩률계산](https://user-images.githubusercontent.com/24896007/95046975-338a9480-0720-11eb-96a3-683e3b2dd2e5.PNG)
1. 사용자 컨트롤 창에서 초기화 버튼을 누르면 시트 삭제 메시지 나타남 '삭제' 선택
![초기화](https://user-images.githubusercontent.com/24896007/95046977-338a9480-0720-11eb-9aa2-c9b0dacc7a97.PNG)
![메시지창4](https://user-images.githubusercontent.com/24896007/95046980-34232b00-0720-11eb-9126-f9fc2c5c523f.PNG)


# 변곡점 구하기 알고리즘
조건1 - 연속선이나 꺽인선이 많아 미분처리 불가
1차: 현재값과 이전값의 차이가 사용자과 정한 기준값보다 크면 변곡점으로 함
1차 단점: 사용자가 기준값을 선택함
2차: 현재값과 이전값의 기울기 차이가 45도 이상이면 변곡점으로 선택
2차 단점: 변곡점이 너무 많아짐, 값이 떨리는 부분에서 45도 이상이 많아짐
3차: 현재값과 이전값의 차이가 최고높이점값의 절반 이상 일때 변곡점으로 함
3차 단점: 이상치 값들이 최고높이점이 될 수 있음
4차: 3차의 최고높이점값을 구할때 상위 2%의 평균치를 사용

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

