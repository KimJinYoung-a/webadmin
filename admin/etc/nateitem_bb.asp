<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/nateitemcls.asp"-->
<%
'' dbget.close()	:	response.End
'' BB(가격비교 사이트)에서 긁어감 (2011-09-20 마이마진에서 BB로 전환)

dim page
page = request("page")
if page="" then page=1

dim sqlStr,ref
ref = Left(request.ServerVariables("REMOTE_ADDR"),250)

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "NAT3-" + ref + "')"
dbget.execute sqlStr

dim oNate, buf
dim totalpage, totalcount
dim ix

set oNate = new CNateItemList
oNate.FPageSize = 500
oNate.FScrollCount = 100
oNate.FTotalCount = totalpage
oNate.FTotalPage = totalcount
oNate.FCurrPage = page
oNate.GetNateItemDB3  

totalpage = oNate.FTotalPage
totalcount = oNate.FTotalCount

buf = "<<<total>>>" & vbCrLf
buf = buf & "	<<<총상품수>>>" & formatNumber(totalcount,0) & vbCrLf
buf = buf & "	<<<최종갱신일>>>" & left(GetCurrentTimeFormat,12) & vbCrLf
'''buf = buf & "	<<<수정/추가상품수>>>0" & vbCrLf
buf = buf & "<<</total>>>" & vbCrLf & vbCrLf

Response.Write buf

for ix=0 to oNate.FResultCount-1
	buf = "<<<product>>>" & vbCrLf
	buf = buf & "	<<<상품아이디>>>" & oNate.FItemList(ix).FItemId & vbCrLf
	buf = buf & "	<<<상품명>>>" & oNate.FItemList(ix).GetModelname & vbCrLf
	buf = buf & "	<<<상품분류명>>>" & oNate.FItemList(ix).getNateBBPath & vbCrLf							'제품분류(카테고리)
	buf = buf & "	<<<제조사>>>" & oNate.FItemList(ix).FitemMaker & vbCrLf
	buf = buf & "	<<<출시일>>>" & vbCrLf
	buf = buf & "	<<<브랜드>>>" & oNate.FItemList(ix).Getmakername & vbCrLf
	buf = buf & "	<<<원산지>>>" & oNate.FItemList(ix).FsourceArea & vbCrLf
	buf = buf & "	<<<상품URL>>>" & Replace(oNate.FItemList(ix).GetItemUrl,"http://","") & vbCrLf			'상품링크
	buf = buf & "	<<<상품이미지URL>>>" & Replace(oNate.FItemList(ix).GetListImageUrl,"http://","") & vbCrLf	'상품이미지(목록:120)
	buf = buf & "	<<<상품큰이미지URL>>>" & Replace(oNate.FItemList(ix).GetBasicImageUrl,"http://","") & vbCrLf	'상품이미지(기본:400)
	buf = buf & "	<<<판매가>>>" & formatNumber(oNate.FItemList(ix).GetPrice,0) & vbCrLf
	buf = buf & "	<<<쿠폰가>>>" & vbCrLf
	buf = buf & "	<<<배송료>>>" & formatNumber(oNate.FItemList(ix).GetDeliverPay,0) & vbCrLf
	buf = buf & "	<<<배송기간>>>" & vbCrLf
	buf = buf& "	<<<할인쿠폰>>>" & oNate.FItemList(ix).GetMMCouponStr & vbCrLf			'할인금액/쿠폰 (현재는 없음)
	buf = buf & "	<<<적립금>>>" & formatNumber(oNate.FItemList(ix).Fmileage,0) & vbCrLf
	buf = buf & "	<<<무이자할부>>>" & vbCrLf																'전상품무이자할부 지원시는 공란
	buf = buf & "	<<<이벤트>>>" & vbCrLf

	buf = buf & "<<</product>>>" & vbCrLf & vbCrLf

	Response.Write buf
next

buf = ""
for ix=0 + oNate.StarScrollPage to oNate.FScrollCount + oNate.StarScrollPage - 1
	if ix > oNate.FTotalpage then Exit for
	buf = buf & "<a href='http://webadmin.10x10.co.kr/admin/etc/nateitem_bb.asp?page=" & ix & "'>" & ix & "</a>"
next

Response.Write buf

set oNate = Nothing

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "NAT4-" + ref + "')"
dbget.execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
