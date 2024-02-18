<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  [OFF]오프_상품관리>>신상품관리
' History : 2008.04.17 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim designer, page, itemid, locationid, datesearch, sdate, edate, itemname, IsOnlineItem, imageList, offmain
Dim sort, i, offlist, offsmall, iTotCnt, itemgubun, isusing, inc3pl
dim yyyy1, mm1, dd1, yyyy2, mm2, dd2, datefg, fromDate, toDate, cdl, cdm, cds
	designer = requestCheckVar(request("designer"),32)
	page = requestCheckVar(request("page"),10)
	itemid = requestCheckVar(request("itemid"),10)
	datesearch = requestCheckVar(request("datesearch"),10)
	sdate = requestCheckVar(request("sdate"),10)
	edate = requestCheckVar(request("edate"),10)
	itemname = requestCheckVar(request("itemname"),124)
	itemgubun = requestCheckVar(request("itemgubun"),2)
	isusing = requestCheckVar(request("isusing"),1)
	sort = requestCheckVar(request("sort"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	datefg = requestCheckVar(request("datefg"),32)
	cdl     = requestCheckVar(request("cdl"),3)
	cdm     = requestCheckVar(request("cdm"),3)
	cds     = requestCheckVar(request("cds"),3)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
If page = "" Then page = 1
If sort = "" Then sort = "itemregdate"

if datefg = "" then datefg = "maechul"

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'/매장일경우 본인 매장만 사용가능
if (C_IS_SHOP) then
	
	'/어드민권한 점장 미만
	'if getlevel_sn("",session("ssBctId")) > 6 then
		locationid = C_STREETSHOPID
	'end if

else
	if (C_IS_Maker_Upche) then
		locationid = session("ssBctID")
	else
		locationid = request("locationid")
	end if
end if

dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 5000
	ioffitem.FCurrPage = page
	ioffitem.FRectShopid = locationid
	ioffitem.FRectDesigner = designer
	ioffitem.FRectDateSearch = datesearch
	ioffitem.FRectSDate = sdate
	ioffitem.FRectEDate = edate
	ioffitem.FRectItemgubun = itemgubun
	ioffitem.FRectItemName = itemname
	ioffitem.FRectItemId = itemid
	ioffitem.FRectIsusing = isusing
	ioffitem.FRectSorting = sort
	ioffitem.frectdatefg = datefg	
	ioffitem.FRectStartDay = fromDate
	ioffitem.FRectEndDay = toDate
	ioffitem.FRectCDL = cdl
	ioffitem.FRectCDM = cdm
	ioffitem.FRectCDS = cds
	ioffitem.FRectInc3pl = inc3pl
	
	If locationid <> "" Then
		ioffitem.GetOffLineNewItemList_xls
	End If

iTotCnt = ioffitem.FTotalCount

Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=신상품관리_검색리스트.xls"
%>

<html>
<head></head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>ITEMID</td>
	<td>BRANDID</td>
	<td>상품명[옵션명]</td>
	<td>사용<br>여부</td>
	<td>상품등록일</td>
	<td>최종업데이트일</td>
	<td>브랜드<Br>최초입고일</td>
	
	<% if locationid <> "" then %>
		<td>상품<Br>최초입고일</td>
	<% end if %>

	<td>
		매출액
	</td>
	
	<% if not(C_IS_SHOP) then %>
		<td>
			매입가
		</td>
	<% end if %>
	
	<td>판매<br>수량</td>
</tr>
<%
If isarray(ioffitem.frectgetrows) Then
'
'		sqlStr = sqlStr & " s.itemgubun, s.shopitemid, s.itemoption, s.makerid, s.shopitemname, s.shopitemoptionname"
'		sqlStr = sqlStr & " , s.isusing ,i.smallimage, IsNULL(s.offimgmain,'') as offimgmain, IsNULL(s.offimglist,'') as offimglist"
'		sqlStr = sqlStr & " , IsNULL(s.offimgsmall,'') as offimgsmall, s.regdate, s.updt, d.firstipgodate, d.shopid "
'
'		If FRectShopid <> "" Then
'			sqlStr = sqlStr & " ,ss.regdate as stockregdate"
'		end if
'
'		sqlStr = sqlStr & " ,IsNULL(t.itemcnt,0) as itemcnt, IsNULL(t.sellsum,0) as sellsum, IsNULL(t.suplyprice,0) as suplyprice"
'		
For i=0 To ubound(ioffitem.frectgetrows,2)
%>
<tr bgcolor="#FFFFFF">
	<td width=80><%=ioffitem.frectgetrows(0,i)%><%=ioffitem.frectgetrows(1,i)%><%=ioffitem.frectgetrows(2,i)%></td>	
	<td><%=ioffitem.frectgetrows(3,i)%></td>
	<td align="left">
		<%=ioffitem.frectgetrows(4,i)%>
		
		<% if ioffitem.frectgetrows(5,i) <> "" then %>
		[<%= ioffitem.frectgetrows(5,i) %>]
		<% end if %>
	</td>
	<td width=30><%=ioffitem.frectgetrows(6,i)%></td>
	<td width=140><%=ioffitem.frectgetrows(11,i)%></td>
	<td width=140><%=ioffitem.frectgetrows(12,i)%></td>
	<td width=80><%=ioffitem.frectgetrows(13,i)%></td>

	<% if locationid <> "" then %>
		<td width=140>
			<%=ioffitem.frectgetrows(15,i)%>
			
			<% if ioffitem.frectgetrows(15,i)<>"" then %>
				<!--<Br><a href="javascript:pop_ipgomaechul('<%= locationid %>','<%=ioffitem.frectgetrows(0,i) & Format00(6,ioffitem.frectgetrows(1,i)) & ioffitem.frectgetrows(2,i)%>','<%= left(left(ioffitem.frectgetrows(15,i),10),4) %>','<%= mid(left(ioffitem.frectgetrows(15,i),10),6,2) %>','<%= right(left(ioffitem.frectgetrows(15,i),10),2) %>','','','');" onfocus="this.blur()">
				날짜별상세매출보기</a>-->
			<% end if %>
		</td>
	<% end if %>

	<td align="right" bgcolor="#E6B9B8" width=80><%= FormatNumber(ioffitem.frectgetrows(17,i),0) %></td>
	
	<% if not(C_IS_SHOP) then %>
		<td align="right" width=80><%= FormatNumber(ioffitem.frectgetrows(18,i),0) %></td>
	<% end if %>
	
	<td align="right" width=60><%= FormatNumber(ioffitem.frectgetrows(16,i),0) %></td>
</tr>
<%
Next

Else
%>
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td align="center" colspan="20">검색된 상품이 없습니다.</td>
	</tr>
<% End If %>
</table>
<%
set ioffitem  = Nothing
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->