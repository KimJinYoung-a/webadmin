<%@ language=vbscript %>
<% option explicit
	session.codePage = 949
%>
<%
'###########################################################
' Description : 출고리스트 엑셀다운
' Hieditor : 2016.12.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
dim idx, menupos, oipchul, oipchuldetail,executedt, storeid,storemarginrate, sqlStr, i
dim sellcashtotal, suplycashtotal, buycashtotal, ttlitemno, BasicMonth, IsExpireEdit

	idx = request("idx")
	menupos = request("menupos")
	storemarginrate = request("storemarginrate")

sellcashtotal  = 0
suplycashtotal = 0
buycashtotal = 0
BasicMonth = CStr(DateSerial(Year(now()),Month(now())-1,1))

set oipchul = new CIpChulStorage
	oipchul.FRectId = idx
	oipchul.GetIpChulMaster

executedt = oipchul.FOneItem.Fexecutedt

if (Left(oipchul.FOneItem.Fcode,2) <> "SO") then
	response.write "<script type='text/javascript'>alert('에러 : 출고코드가 아닙니다.');</script>"
	response.write "<br><br>에러 : 출고코드가 아닙니다." & oipchul.FOneItem.Fcode
	response.end
end if

set oipchuldetail = new CIpChulStorage
	oipchuldetail.FRectStoragecode = oipchul.FOneItem.Fcode
	oipchuldetail.GetIpChulDetail

if IsNULL(oipchul.FOneItem.Fexecutedt) then
	IsExpireEdit = Lcase(CStr(false))
else
	IsExpireEdit = Lcase(CStr(CDate(oipchul.FOneItem.Fexecutedt)<Cdate(BasicMonth)))
end if

if (  (storemarginrate = "") or (storemarginrate = "0") ) then
	sqlStr = "select IsNull(a.marginrate, 0) as marginrate "
	sqlStr = sqlStr + " from [db_storage].[dbo].vw_acount_user_delivery a "
	sqlStr = sqlStr + " where a.userid = '" +  oipchul.FOneItem.Fsocid  + "' "
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		storemarginrate = rsget("marginrate")
	else
		storemarginrate = "0"
	end if
	rsget.close
elseif (storemarginrate = "") then
	storemarginrate = "0"
end if

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>

<html>
<head>
<meta http-equiv='Content-Type' content='text/html;charset=euc-kr' />
<title>출고 상세 내역 엑셀 다운로드</title>

<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>

</head>
<body>
<table width="100%" border="1" align="center" cellpadding="1" cellspacing="1" bgcolor="black">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>상품코드</td>
	<td>브랜드ID</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td>수량</td>
	<td>판매가</td>
	<td>출고가</td>
	<td>매입가</td>
	<td>출고<br>할인율</td>
	<td>매입<br>마진</td>
	<td>출고<br>구분</td>
	<td>매입<br>구분</td>
	<td>센터<br>매입<br>구분</td>
</tr>
<% for i=0 to oipchuldetail.FResultCount -1 %>
<%
ttlitemno = ttlitemno + oipchuldetail.FItemList(i).Fitemno
sellcashtotal = sellcashtotal + oipchuldetail.FItemList(i).Fitemno * oipchuldetail.FItemList(i).Fsellcash
suplycashtotal = suplycashtotal + oipchuldetail.FItemList(i).Fitemno * oipchuldetail.FItemList(i).Fsuplycash
buycashtotal = buycashtotal + oipchuldetail.FItemList(i).Fitemno * oipchuldetail.FItemList(i).Fbuycash
%>
<tr bgcolor="#FFFFFF">
	<td align="center">
		<a href="javascript:popViewCurrentStock('<%= oipchuldetail.FItemList(i).Fiitemgubun %>', '<%= oipchuldetail.FItemList(i).FItemId %>', '<%= oipchuldetail.FItemList(i).FItemOption %>');">
			<%= oipchuldetail.FItemList(i).Fiitemgubun %><%= Format00(8,oipchuldetail.FItemList(i).FItemId) %><%= oipchuldetail.FItemList(i).FItemOption %>
		</a>
	</td>
	<td align="center" class='txt'><%= oipchuldetail.FItemList(i).Fimakerid %></td>
	<td class='txt'><%= oipchuldetail.FItemList(i).FIItemName %></td>
	<td align=center><%= oipchuldetail.FItemList(i).FIItemoptionName %></td>
	<td align=center>
		<%= oipchuldetail.FItemList(i).Fitemno %>
	</td>
	<td align=right>
		<%= oipchuldetail.FItemList(i).Fsellcash %>
	</td>
	<td align=right>
		<%= oipchuldetail.FItemList(i).Fsuplycash %>
	</td>
	<td align=right>
		<%= oipchuldetail.FItemList(i).Fbuycash %>
	</td>
	<td align=center>
	<% if oipchuldetail.FItemList(i).Fsellcash<>0 then %>
	<%= 100-CLng(oipchuldetail.FItemList(i).Fsuplycash/oipchuldetail.FItemList(i).Fsellcash*100*100)/100 %>%
	<% end if %>
	</td>
	<td align=center>
	<% if oipchuldetail.FItemList(i).Fsellcash<>0 then %>
	<%= 100-CLng(oipchuldetail.FItemList(i).Fbuycash/oipchuldetail.FItemList(i).Fsellcash*100*100)/100 %>%
	<% end if %>
	</td>
	<td align="center"><%= oipchuldetail.FItemList(i).FMWgubun %></td>
	<% if (C_ADMIN_AUTH) and ((oipchuldetail.FItemList(i).FOnlineMwdiv="W") and (oipchuldetail.FItemList(i).FMWgubun<>"C")) or (oipchuldetail.FItemList(i).FOnlineMwdiv="U") then %>
	<td align="center"><font color="<%= oipchuldetail.FItemList(i).getOnlineMwdivColor %>"><%= oipchuldetail.FItemList(i).FOnlineMwdiv %></font></td>
	<td align="center"><font color="<%= oipchuldetail.FItemList(i).getOnlineMwdivColor %>"><%= oipchuldetail.FItemList(i).FCenterMwdiv %></font></td>
	<% else %>
	<td align="center"><font color="<%= oipchuldetail.FItemList(i).getOnlineMwdivColor %>"><%= oipchuldetail.FItemList(i).FOnlineMwdiv %></font></td>
	<td align="center"><font color="<%= oipchuldetail.FItemList(i).getOnlineMwdivColor %>"><%= oipchuldetail.FItemList(i).FCenterMwdiv %></font></td>
	<% end if %>
	<input type="hidden" name="didx" value="<%= oipchuldetail.FItemList(i).Fid %>">
</tr>
<% next %>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td colspan=4 align="center">총계</td>
	<td align="center"><%= FormatNumber(ttlitemno,0) %></td>
	<td align="right"><b><%= FormatNumber(sellcashtotal,0) %></b></td>
	<td align="right"><b><%= FormatNumber(suplycashtotal,0) %></b></td>
	<td align="right"><b><%= FormatNumber(buycashtotal,0) %></b></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>
</table>
</body>
</html>

<%
set oipchuldetail = Nothing
set oipchul = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
