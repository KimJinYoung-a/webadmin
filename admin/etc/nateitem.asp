<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/yahooitemcls.asp"-->
<%
'' dbget.close()	:	response.End
'' 마이마진(가격비교 사이트)에서 긁어감, (네이트와 제휴되어있음;양식참고:http://www.mm.co.kr/shop_admin/reg/shop_plist.asp?menu=03)

dim nowdate
dim adate,bdate
dim fso, FileName,tFile,appPath
dim readtextfile

appPath = server.mappath("/admin/etc/nate/") + "\"
FileName = "nateitem.txt"

nowdate = now()
adate = CDate(Left(nowdate,10) + " 09:00:00")
bdate = CDate(Left(nowdate,10) + " 17:00:00")

dim sqlStr,ref
ref = Left(request.ServerVariables("REMOTE_ADDR"),250)

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "NAT1-" + ref + "')"
dbget.execute sqlStr

if ((nowdate>adate) and (nowdate<bdate)) then
    '오전 9시 ~ 오후 6시에는 다시 파일을 만들지말고 기존의 자료 그대로 출력
    '업데이트 주기 08, 12, 15, 18시(총 4번)
    response.redirect "/admin/etc/nate/" & FileName
    dbget.close()	:	response.End
end if

dim oNate, buf
dim totalpage, totalcount
dim ix, j

'// 파일 컨트롤 호출
Set fso = CreateObject("Scripting.FileSystemObject")
Set tFile = fso.CreateTextFile(appPath & FileName )

'// 전체 페이지수와 총상품수 접수
set oNate = new CYahooItemList
oNate.FPageSize = 300
oNate.FScrollCount = 100
oNate.GetNateItemCountDB3
	totalpage = oNate.FTotalPage
	totalcount = oNate.FTotalCount
set oNate = Nothing

buf = "<p>TOTAL:" & totalcount
tFile.WriteLine buf

'// 도돌이 저장
for j=0 to totalpage - 1
	set oNate = new CYahooItemList
	oNate.FPageSize = 300
	oNate.FScrollCount = 100
	oNate.FTotalCount = totalpage
	oNate.FTotalPage = totalcount
	oNate.FCurrPage = j+1
	oNate.GetNateItemDB3  

	for ix=0 to oNate.FResultCount-1
		buf = "<p>"
		buf = buf & "tenbyten" & oNate.FItemList(ix).FItemId & "[^]"			'상품코드(쇼핑몰ID+상품코드)
		buf = buf & oNate.FItemList(ix).GetModelname & "[^]"					'상품명
		buf = buf & oNate.FItemList(ix).GetItemUrl & "[^]"						'상품링크
		buf = buf & oNate.FItemList(ix).GetPrice & "[^]"						'판매가
		buf = buf & oNate.FItemList(ix).getNatePath & "[^]"						'제품분류(카테고리)
		buf = buf & oNate.FItemList(ix).Getmakername & "[^]"					'제조사(브랜드)
		buf = buf & oNate.FItemList(ix).GetImageUrl & "[^]"						'상품이미지
		buf = buf & oNate.FItemList(ix).GetDeliverPay & "원[^]"					'배송료
		buf = buf & oNate.FItemList(ix).GetMMCouponStr & "[^]"					'할인금액/쿠폰 (현재는 없음)
		buf = buf & "[^][^]"

		if buf<>"" then
			tFile.WriteLine buf
		end if
	next

	set oNate = Nothing
next

tFile.Close

Set tFile = Nothing
Set fso = Nothing

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "NAT2-" + ref + "')"
dbget.execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<%
response.redirect "/admin/etc/nate/" & FileName
%>