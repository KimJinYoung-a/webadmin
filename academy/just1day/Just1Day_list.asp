<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/sitemaster/just1DayCls.asp"-->
<%
'###############################################
' PageName : 핑거스 Just1Day_list.asp
' Discription : 저스트 원데이 목록
' History : 2016.08.01 유태욱 : 핑거스 just1day
'###############################################

dim page, sDt, eDt, itemid, i, lp, dispCate

page = RequestCheckvar(request("page"),10)
if page = "" then page=1
sDt = RequestCheckvar(request("sDt"),10)
eDt = RequestCheckvar(request("eDt"),10)
itemid = RequestCheckvar(request("itemid"),10)
dispCate = requestCheckvar(request("disp"),16)

dim oJust
set oJust = New Cjust1Day
oJust.FCurrPage = page
oJust.FPageSize=20
oJust.FRectSdt = sDt
oJust.FRectEdt = eDt
oJust.FRectItemId = itemid
oJust.FRectDispCate		= dispCate
oJust.Getjust1DayList

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
// 페이지 이동
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="Just1Day_list.asp";
	document.refreshFrm.submit();
}

</script>
<!-- 상단 검색폼 시작 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" action="Just1Day_list.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="fidx">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		기간 
		<input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> /
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "sDt", trigger    : "sDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "eDt", trigger    : "eDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		상품코드 <input type="text" name="itemid" class="text" size="12" value="<%=itemid%>">
		&nbsp;
		전시카테고리: <!-- #include virtual="/academy/comm/dispCateSelectBox.asp"--> 
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">

<tr>
	<!--
	<td><a href="" onclick="AssignXMLReal(refreshFrm,2); return false;"><img src="/images/refreshcpage.gif" border="0"> XML Real 적용</a></td>
	-->
	<td align="right"><input type="button" value="아이템 추가" onclick="self.location='Just1Day_write.asp?mode=add&menupos=<%= menupos %>'" class="button"></td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		검색결과 : <b><%=oJust.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oJust.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<!--	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>-->
<!--		<td>idx</td>-->
	<td>날짜</td>
	<td>Image</td>
	<td>제품명</td>
	<td>전시카테고리</td>
	<td>할인률</td>
	<td>품절</td>
	<td>등록일</td>
</tr>
<%	if oJust.FResultCount < 1 then %>
<tr>
	<td colspan="9" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 아이템이 없습니다.</td>
</tr>
<%
	else
		for i=0 to oJust.FResultCount-1
%>
<!--	<form action="" name="frmBuyPrc<%=i%>" method="get">-->
	<a href="Just1Day_write.asp?mode=edit&menupos=<%= menupos %>&justdate=<%= oJust.FItemList(i).FjustDate %>">
	<tr <% if cstr(oJust.FItemList(i).FjustDate) < cstr(date()) then %>bgcolor="<%= adminColor("dgray") %>"<% else %>bgcolor="#FFFFFF" style="cursor:pointer;" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='ffffff';<% end if %>>
<!--			<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>		-->
<!--			<td align="center"><%= oJust.FItemList(i).Fidx %><input type="hidden" name="idx" value="<%= oJust.FItemList(i).Fidx %>"></td>-->
		<td align="center"><a href="Just1Day_write.asp?mode=edit&menupos=<%= menupos %>&justdate=<%= oJust.FItemList(i).FjustDate %>"><%= oJust.FItemList(i).FjustDate %></a></td>
		<td align="center"><img src="<%= oJust.FItemList(i).FsmallImage %>" height="40" border="0"></td>
		<td align="center"><%= "[" & oJust.FItemList(i).FItemID & "] " & oJust.FItemList(i).FItemname %></td>
		<td align="center"><%=fnCateCodeNameSplit(oJust.FItemList(i).FCateName,oJust.FItemList(i).FItemID)%></span></td>
		<td align="center"><%= formatPercent(1-oJust.FItemList(i).FjustSalePrice/oJust.FItemList(i).ForgPrice,0) %></td>
		<td align="center"><% if oJust.FItemList(i).FsellYn<>"Y" then Response.Write "품절" %></td>
		<td align="center"><%= left(oJust.FItemList(i).Fregdate,10) %></td>
	</tr>
	</a>
<!--	</form>-->
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<tr bgcolor="#FFFFFF">
	<td colspan="9" align="center">
	<!-- 페이지 시작 -->
	<%
		if oJust.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oJust.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oJust.StartScrollPage to oJust.FScrollCount + oJust.StartScrollPage - 1

			if lp>oJust.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if oJust.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- 페이지 끝 -->
	</td>
</tr>

</table>
<%
set oJust = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->