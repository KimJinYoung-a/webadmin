<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 주문
' History : 이상구 생성
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim page, shopid
page = requestCheckVar(request("page"),10)
shopid = session("ssBctBigo")

if page="" then page=1

if (left(shopid,Len("streetshop"))<>"streetshop") then
	response.write "<script type='text/javascript'>alert('샾 구분이 지정되지 않았습니다. - 관리자 문의요망');</script>"
	dbget.close()	:	response.End
end if

dim osheet
set osheet = new COrderSheet
osheet.FCurrPage = page
osheet.Fpagesize=20
osheet.FRectBaljuId = shopid
osheet.GetOrderSheetList

dim i
dim totaljumunsuply, totalfixsuply, totaljumunsellcash
%>
<script type='text/javascript'>
//function PopIpgoSheet(v,itype){
//	var popwin;
//	popwin = window.open('popshopjumunsheet.asp?idx=' + v + '&itype=' + itype,'shopjumunsheet','width=680,height=600,scrollbars=yes,status=no');
//	popwin.focus();
//}

function PopIpgoSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/popshopjumunsheet2.asp?idx=' + v ,'shopjumunsheet','width=740,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function ExcelSheet(v){
	window.open('/admin/fran/popshopjumunsheet2.asp?idx=' + v + '&xl=on');
}

</script>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
	<td width=80>주문코드</td>
	<td width=110>공급처</td>
	<td width=80>주문일</td>
	<td width=80>입고요청일</td>
	<td width=100>주문상태</td>
	<td width=80>총주문액<br>(소비자가)</td>
	<td width=80>총주문액<br>(공급가)</td>
	<td width=80>확정금액<br>(공급가)</td>
	<td width=80>발송일</td>
	<td width=100>송장번호</td>
	<td width=70>내역서</td>
</tr>
<% if osheet.FResultCount >0 then %>
<% for i=0 to osheet.FResultcount-1 %>
<%
totaljumunsellcash = totaljumunsellcash + osheet.FItemList(i).Fjumunsellcash
totaljumunsuply = totaljumunsuply + osheet.FItemList(i).Fjumunsuplycash
totalfixsuply   = totalfixsuply + osheet.FItemList(i).Ftotalsuplycash
%>
<tr bgcolor="#FFFFFF">
	<td align=center><a href="offjumuninputedit.asp?idx=<%= osheet.FItemList(i).Fidx %>"><%= osheet.FItemList(i).Fbaljucode %></a></td>
	<td align=center><%= osheet.FItemList(i).Ftargetid %><br>(<%= osheet.FItemList(i).Ftargetname %>)</td>
	<td align=center><%= Left(osheet.FItemList(i).FRegdate,10) %><br>(<%= osheet.FItemList(i).Fregname %>)</td>
	<td align=center><%= Left(osheet.FItemList(i).Fscheduledate,10) %></td>
	<td align=center><font color="<%= osheet.FItemList(i).GetStateColor %>"><%= osheet.FItemList(i).GetStateName %></font></td>
	<td align=right><%= FormatNumber(osheet.FItemList(i).Fjumunsellcash,0) %></td>
	<td align=right><%= FormatNumber(osheet.FItemList(i).Fjumunsuplycash,0) %></td>
	<td align=right><%= FormatNumber(osheet.FItemList(i).Ftotalsuplycash,0) %></td>
	<td align=center><%= Left(osheet.FItemList(i).Fbeasongdate,10) %></td>
	<td align=center><%= Left(osheet.FItemList(i).Fsongjangno,10) %></td>
	<td align=center><a href="javascript:PopIpgoSheet('<%= osheet.FItemList(i).FIdx %>');"><img src="/images/iexplorer.gif" width=21 border=0></a> <a href="javascript:ExcelSheet('<%= osheet.FItemList(i).FIdx %>');"><img src="/images/iexcel.gif" width=21 border=0></a></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td>총계</td>
	<td colspan=4></td>
	<td align=right><%= formatNumber(totaljumunsellcash,0) %></td>
	<td align=right><%= formatNumber(totaljumunsuply,0) %></td>
	<td align=right><%= formatNumber(totalfixsuply,0) %></td>
	<td colspan=3></td>
</tr>
<tr bgcolor="#FFFFFF" height=20>
	<td colspan=11 align=center>
	<% if osheet.HasPreScroll then %>
		<a href="?page=<%= osheet.StartScrollPage-1 %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + osheet.StartScrollPage to osheet.FScrollCount + osheet.StartScrollPage - 1 %>
		<% if i>osheet.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if osheet.HasNextScroll then %>
		<a href="?page=<%= i %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=11 align=center>[ 검색결과가 없습니다. ]</td>
</tr>
<% end if %>
</table>
<%
set osheet = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->