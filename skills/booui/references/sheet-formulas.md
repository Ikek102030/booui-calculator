# 시트 수식 모음

실제 장례 정산에서 쓰고 검증한 수식이다. 탭 이름이 다르면 맞춰 바꾼다.
돈 서식은 전부 `#,##0"원"`.

## 부의금 입력

| 칸 | 수식 |
|---|---|
| F1 확인 필요 | `=ARRAYFORMULA(IF(ROW(A1:A)=1,"확인 필요",IF((LEN(A1:A&B1:B&C1:C&D1:D&E1:E)=0)+REGEXMATCH(B1:B&"","소계\|합계"),"",TRIM(IF(B1:B="","이름 없음 ","")&IF(A1:A="","받은 가족 없음 ","")&IF(C1:C="","금액 없음 ",IF(ISNUMBER(C1:C),IF(C1:C<1000,"금액 확인 ",""),"금액 확인 "))&IF(B1:B="","",IF(COUNTIFS(A1:A,A1:A,B1:B,B1:B,C1:C,C1:C)>1,"중복 의심",""))))))` |
| G1 일차(자동) | `=ARRAYFORMULA(IF(ROW(B1:B)=1,"일차 (자동)",IF((LEN(A1:A&B1:B&C1:C&D1:D&E1:E)=0)+REGEXMATCH(B1:B&"","소계\|합계"),"",1+COUNTIFS(B1:B,"*소계*",ROW(B1:B),"<"&ROW(B1:B)))))` |

(표 안의 `\|` 는 실제 수식에서 `|` 다.)

- A2:A 데이터 확인: 목록 범위 `=요약!$A$6:$A$200`, strict=false(새 이름 입력 허용), 드롭다운 표시.
- C2:C 데이터 확인: 1000 이상 숫자, strict=false.
- 조건부 서식: `=REGEXMATCH($B2&"","소계|합계")` → 굵게 + 연파랑(구분선), `=$F2<>""` → 연노랑.
- G 는 '소계' 줄만 세서 일차를 나눈다. '총합계' 줄은 일차를 늘리지 않는다.

## 근조 화환

| 칸 | 수식 |
|---|---|
| F1 확인 필요 | `=ARRAYFORMULA(IF(ROW(A1:A)=1,"확인 필요",IF(LEN(A1:A&B1:B&C1:C&D1:D&E1:E)=0,"",TRIM(IF(A1:A="","받은 가족 없음 ","")&IF(B1:B="","보낸 분 없음 ","")&IF(B1:B="","",IF(COUNTIFS(B1:B,B1:B,C1:C,C1:C)>1,"중복 의심",""))))))` |

- C 종류 드롭다운: 근조화환 · 근조 바구니 · 근조기 · 쌀화환 · 기타 (strict=false).
- 중복은 보낸 분 + 종류가 둘 다 같을 때만.

## 요약

위쪽 숫자:

| 칸 | 수식 |
|---|---|
| B1 총 합계 | `=SUMIF('부의금 입력'!G2:G,">0",'부의금 입력'!C2:C)` (일차가 매겨진 줄 = 구분선 뺀 줄) |
| B2 총 건수 | `=COUNTIF('부의금 입력'!G2:G,">0")` |
| E1 조화·물품 | `=SUMPRODUCT(LEN('근조 화환'!A2:A&'근조 화환'!B2:B&'근조 화환'!C2:C)>0)` |
| E3 장부 총합계 | `=IFERROR(INDEX(FILTER('부의금 입력'!C2:C,REGEXMATCH('부의금 입력'!B2:B&"","총\s*합계")),1),"")` |
| E2 장부와 차이 | `=IF(ISNUMBER(E3),B1-E3,"")` |
| I1 지출 총액 / I2 남는 돈 | `=지출!H4` / `=지출!H5` |

가족 표(5행 머리글, 6행부터). **펼침 범위는 6~24행으로 막는다.**

| 칸 | 수식 |
|---|---|
| A6 받은 가족 | `=IFERROR(UNIQUE(FILTER({'부의금 입력'!A2:A;'근조 화환'!A2:A},{'부의금 입력'!A2:A;'근조 화환'!A2:A}<>"")),"")` |
| B6 건수 | `=ARRAYFORMULA(IF(A6:A24="","",COUNTIF('부의금 입력'!A2:A,A6:A24)))` |
| C6 합계 | `=ARRAYFORMULA(IF(A6:A24="","",SUMIF('부의금 입력'!A2:A,A6:A24,'부의금 입력'!C2:C)))` |
| D6 비중 | `=ARRAYFORMULA(IF(A6:A24="","",IFERROR(C6:C24/$B$1,0)))` |
| E6 조화·물품 | `=ARRAYFORMULA(IF(A6:A24="","",COUNTIF('근조 화환'!A2:A,A6:A24)))` |

F 는 좁은 여백. G 부터는 사용자가 만드는 정산 열 자리(예: `=C6*77%`). `deliver.py` 는 여백 다음 열들을 정산 열로 읽어 PDF 카드에 넣는다.

일차별 소계(K4 제목, K5 머리글):

| 칸 | 수식 |
|---|---|
| K6 일차 | `=IFERROR(SORT(UNIQUE(FILTER('부의금 입력'!G2:G,'부의금 입력'!G2:G<>""))),"")` |
| L6 건수 | `=ARRAYFORMULA(IF(ISNUMBER(K6:K24),COUNTIF('부의금 입력'!G2:G,K6:K24),""))` |
| M6 합계 | `=ARRAYFORMULA(IF(ISNUMBER(K6:K24),SUMIF('부의금 입력'!G2:G,K6:K24,'부의금 입력'!C2:C),""))` |
| N6 장부에 적은 소계 | `=MAP(K6:K24,LAMBDA(d,IF(NOT(ISNUMBER(d)),"",IFERROR(INDEX(FILTER('부의금 입력'!C2:C,REGEXMATCH('부의금 입력'!B2:B&"","소계")),d),"(아직 없음)"))))` |
| O6 차이 | `=MAP(M6:M24,N6:N24,LAMBDA(j,k,IF(OR(NOT(ISNUMBER(j)),NOT(ISNUMBER(k))),"",j-k)))` |

## 지출

| 칸 | 수식 |
|---|---|
| E1 확인 필요 | `=ARRAYFORMULA(IF(ROW(A1:A)=1,"확인 필요",IF((LEN(A1:A&B1:B&C1:C&D1:D)=0)+REGEXMATCH(A1:A&"","합계\|소계"),"",TRIM(IF(A1:A="","항목 없음 ","")&IF(B1:B="","금액 없음 ",IF(ISNUMBER(B1:B),"","금액 확인 "))&IF(D1:D="","낸 사람 없음","")))))` |
| H3 부의금 총액 | `=요약!B1` |
| H4 지출 총액 | `=IFERROR(SUM(FILTER(B2:B,ISNUMBER(B2:B),NOT(REGEXMATCH(A2:A&"","합계\|소계")))),0)` |
| H5 남는 돈 | `=H3-H4` |
| G9 낸 사람별 | `=IFERROR(ARRAYFORMULA(LET(item_col,A2:A,amount_col,B2:B,payer_col,D2:D,keep_row,(LEN(item_col&amount_col)>0)*NOT(REGEXMATCH(item_col&"","합계\|소계")),payers,FILTER(IF(payer_col="","(미기재)",payer_col),keep_row),money,FILTER(IF(ISNUMBER(amount_col),amount_col,0),keep_row),payer_list,UNIQUE(payers),paid_sum,MAP(payer_list,LAMBDA(x,SUM((payers=x)*money))),paid_cnt,MAP(payer_list,LAMBDA(x,SUM(--(payers=x)))),SORT(HSTACK(payer_list,paid_sum,paid_cnt),2,FALSE))),"")` |
| N3 공금 시작 | 사용자 값 |
| N4 공금에서 나간 돈 | `=SUMIF(D2:D,"공금",B2:B)` |
| N5 공금 잔액 | `=IF(ISNUMBER(N3),N3-N4,"")` |
| Z2 (숨김) 낸 사람 목록 | `=IFERROR(VSTACK("공금",FILTER(요약!A6:A200,요약!A6:A200<>"")),"공금")` |

## 물품 통계

| 칸 | 수식 |
|---|---|
| A4 종류별 | `=ARRAYFORMULA(LET(sender,'근조 화환'!B2:B,kind_raw,'근조 화환'!C2:C,has_row,LEN(sender&kind_raw)>0,kinds,FILTER(IF(kind_raw="","(종류 없음)",kind_raw),has_row),kind_list,UNIQUE(kinds),cnt,MAP(kind_list,LAMBDA(x,SUM(--(kinds=x)))),VSTACK({"종류","개수","비중"},SORT(HSTACK(kind_list,cnt,cnt/ROWS(kinds)),2,FALSE),HSTACK("합계",ROWS(kinds),1))))` |
| E4 가족 × 종류 | `=ARRAYFORMULA(LET(sender,'근조 화환'!B2:B,fam_raw,'근조 화환'!A2:A,kind_raw,'근조 화환'!C2:C,has_row,LEN(sender&kind_raw)>0,fams,FILTER(IF(fam_raw="","(미지정)",fam_raw),has_row),kinds,FILTER(IF(kind_raw="","(종류 없음)",kind_raw),has_row),fam_list,UNIQUE(fams),kind_uniq,UNIQUE(kinds),kind_list,TRANSPOSE(SORT(kind_uniq,MAP(kind_uniq,LAMBDA(x,SUM(--(kinds=x)))),FALSE)),grid,MAKEARRAY(ROWS(fam_list),COLUMNS(kind_list),LAMBDA(r,c,SUM((fams=INDEX(fam_list,r,1))*(kinds=INDEX(kind_list,1,c))))),row_tot,BYROW(grid,LAMBDA(x,SUM(x))),body,SORT(HSTACK(fam_list,grid,row_tot),COLUMNS(kind_list)+2,FALSE),VSTACK(HSTACK("받은 가족",kind_list,"합계"),body,HSTACK("합계",BYCOL(grid,LAMBDA(x,SUM(x))),SUM(grid)))))` |

## 가족별(드롭다운)

| 칸 | 수식 |
|---|---|
| B1 | 드롭다운 `=요약!$A$6:$A$200` (strict) |
| D1 / F1 / I1 | `=SUMIF('부의금 입력'!A2:A,B1,'부의금 입력'!C2:C)` / `=COUNTIF('부의금 입력'!A2:A,B1)` / `=COUNTIF('근조 화환'!A2:A,B1)` |
| A4 부의금 명단 | `=IFERROR(LET(r,FILTER({'부의금 입력'!B2:B,'부의금 입력'!C2:C,'부의금 입력'!D2:D,'부의금 입력'!E2:E},'부의금 입력'!A2:A=$B$1),{SEQUENCE(ROWS(r)),r}),"")` |
| H4 조화·물품 | `=IFERROR(FILTER({'근조 화환'!B2:B,'근조 화환'!C2:C,'근조 화환'!D2:D,'근조 화환'!E2:E},'근조 화환'!A2:A=$B$1),"")` |

A4:K 에 `warningOnly` 보호 범위를 건다. 2행에 "보기 전용, 메모는 원본 탭에" 안내.

## 사람별 탭(7단계용)

가족별 탭과 같지만 B1 이 드롭다운이 아니라 이름 글자다. A1 은 반드시 `받은 가족`, B1 은 탭 이름과 같게(`deliver.py` 가 이걸로 사람 탭을 찾는다).
한 탭을 만든 뒤 `duplicateSheet`(newSheetName = 이름) → `updateCells` 로 B1 만 바꾼다.

## 정산 방법 고르기 (요약)

사용자가 시트에서 방법을 바꿀 수 있게 드롭다운으로 둔다. 가족 순서는 UNIQUE 로 바뀔 수 있으니, 빠질 가족은 줄 위치가 아니라 **이름 목록**으로 받는다.
수식은 검증한 엑셀 양식과 같은 꼴(한 줄씩, 6~24행에 채워 내림)이고 구글 시트·엑셀 공통 함수만 쓴다.

| 칸 | 수식 |
|---|---|
| G4 / H4 | `정산 방법 ▶` / 드롭다운 `1. 똑같이 나누기,2. 받은 부의금 비율,3. 손님 수 비율` (처음엔 `▼ 정산 방법을 골라 주세요`) |
| Q4 (숨김) 방법 번호 | `=IFERROR(VALUE(LEFT($H$4,1)),0)` |
| T5 / T6:T15 | `나눠 내지 않을 가족` / 사용자가 이름을 적는 칸 |
| R6 (숨김) 가중치 | `=IF(OR($A6="",COUNTIF($T$6:$T$15,$A6)>0,$Q$4=0),0,CHOOSE($Q$4,1,$C6,$B6))` |
| G6 지출 분담 | `=IF(OR($A6="",$Q$4=0),"",IF(SUM($R$6:$R$24)=0,0,지출!$H$4*R6/SUM($R$6:$R$24)))` |
| H6 남는 몫 | `=IF(G6="","",C6-G6)` |

R6·G6·H6 은 24행까지 채워 내린다. 엑셀 양식(make_template.py)은 같은 계산에서 참여 여부만 '가족 목록' B열(X 면 빠짐)로 받는다.
