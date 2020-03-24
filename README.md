# 엠앤와이즈 알림톡 템플릿 관리

## 필요한 기능

- 알림톡 템플릿 자동으로 신청 (register.py)
  [{"name": "참여하기",
  "type": "WL", "url_mobile": "http://plus.kakao.com/talk/bot/@#{포켓서베이}/참여하기#{참여코드}"
  }
  ][{"name": "참여하기","type": "wl", "url_mobile": "http://plus.kakao.com/talk/bot/@#{pocketsurvey}/참여하기#{참여코드}"}]
- 알림톡 템플릿 날짜 범위 지정으로 신청 결과 조회 (result.py)
- 알림톡 템플릿 검수된것들 엑셀이나 csv, json 등 중 한 포멧으로 내보내기 (result.py)

## 사용한 패키지

- requests
- selenium (driver로 실행하면 실제 구동되는 걸 볼 수 있음. browser를 사용하면 보여주지 않음.)
- beautifulsoup
- openpyxl

## 템플릿 등록하기

- python register.py 엑셀파일제목
- 엑셀 파일은 버튼 열을 제외한 아래의 내용을 사전 편집 후 저장
  "검색용아이디" : 챗봇 아이디(포켓서베이의 경우 pocketsurvey)
  "템플릿코드(최대 30자)" : custom
  "템플릿명(최대 200자)" : custom
  "탬플릿유형" : BA
  "템플릿내용(최대 1000자)" : custom
  "부가정보(최대 500자/변수불가)" : custom
  "광고성메시지(최대 80자/url,변수불가)" : custom
  "PC노출여부": Y

## 템플릿 검수 현황 조회하기

- python result.py 조회시작일 조회종료일
- 조회일은 연월일로 작성 ex) 2020.03.20 or 2020/03/20 상관없음
