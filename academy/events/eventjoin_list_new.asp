<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/academy/eventPrizeCls.asp"-->
<%
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, fromDate ,toDate, tmpDate
dim userid, eventCode, eventName, research, page, i
	menupos = requestCheckVar(Request("menupos"),32)
	research = requestCheckVar(Request("research"),32)
	page = requestCheckVar(Request("page"),32)
	yyyy1   = RequestCheckvar(request("yyyy1"),4)
	mm1     = RequestCheckvar(request("mm1"),2)
	dd1     = RequestCheckvar(request("dd1"),2)
	yyyy2   = RequestCheckvar(request("yyyy2"),4)
	mm2     = RequestCheckvar(request("mm2"),2)
	dd2     = RequestCheckvar(request("dd2"),2)
	userid = requestCheckVar(Request("userid"),32)
	eventCode = requestCheckVar(Request("eventCode"),32)
	eventName = requestCheckVar(Request("eventName"),32)

if (page = "") then
	page = "1"
end if

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) + 1), 1)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, toDate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if

dim oCEventPrize
set oCEventPrize = new CEventPrize
	oCEventPrize.FRectUserid = userid
	oCEventPrize.FRectEventCode = eventCode
	oCEventPrize.FRectEventName = eventName
	oCEventPrize.FRectStartdate = fromDate
	oCEventPrize.FRectEndDate = toDate
	oCEventPrize.FPageSize = 20
	oCEventPrize.FCurrPage = page
	oCEventPrize.GetEventJoinListNew
%>

<script language='javascript'>

function fnGotoPage(page) {
	document.frm.page.value = page;
	document.frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		&nbsp;
		아이디 :
		<input type="text" class="text" size="16" maxlength="32" name="userid" value="<%= userid %>">
		&nbsp;
		이벤트코드 :
		<input type="text" class="text" size="8" maxlength="32" name="eventCode" value="<%= eventCode %>">
		&nbsp;
		이벤트명 :
		<input type="text" class="text" size="20" maxlength="32" name="eventName" value="<%= eventName %>">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		&nbsp;
		참여일자 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
</tr>
</form>
</table>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oCEventPrize.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oCEventPrize.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center" width="100">구분</td>
	<td align="center" width="60">이벤트<br>코드</td>
	<td align="center" width="400">이벤트명</td>
	<td align="center" width="80">시작일</td>
	<td align="center" width="80">종료일</td>
	<td align="center" width="80">당첨자발표</td>
	<td align="center" width="100">아이디</td>
	<td align="center">코멘트</td>
	<td align="center" width="150">참여일</td>
	<td align="center" width="50">상태</td>
	<td align="center">비고</td>
</tr>
<% if oCEventPrize.FresultCount>0 then %>
<% for i=0 to oCEventPrize.FresultCount-1 %>
	<% if oCEventPrize.FItemList(i).finvaliduserid<>"" then %>
		<tr align="center" bgcolor="#e1e1e1" onmouseover=this.style.background="<%= adminColor("tabletop") %>"; onmouseout=this.style.background='#e1e1e1';>
	<% else %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="<%= adminColor("tabletop") %>"; onmouseout=this.style.background='#FFFFFF';>
	<% end if %>

	<td align="center" height="25">
		<%= oCEventPrize.FItemList(i).GetEventGubunName %>
	</td>
	<td align="center" height="25">
		<%= oCEventPrize.FItemList(i).Fevt_code %>
	</td>
	<td align="left">
		<%= oCEventPrize.FItemList(i).Fevt_name %>
	</td>
	<td align="center">
		<%= oCEventPrize.FItemList(i).Fevt_startdate %>
	</td>
	<td align="center">
		<%= oCEventPrize.FItemList(i).Fevt_enddate %>
	</td>
	<td align="center">
		<%= oCEventPrize.FItemList(i).Fevt_prizedate %>
	</td>
	<td align="left">
		<%= oCEventPrize.FItemList(i).Fuserid %>
	</td>
	<td align="left">
		<%= oCEventPrize.FItemList(i).Fcomment %>
	</td>
	<td align="center">
		<%= oCEventPrize.FItemList(i).Fregdate %>
	</td>
	<td align="center">
		<%= oCEventPrize.FItemList(i).GetIsUsingStr %>
	</td>
	<td align="center">
		<% if oCEventPrize.FItemList(i).finvaliduserid<>"" then %>
			불량고객
		<% end if %>
	</td>
</tr>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if oCEventPrize.HasPreScroll then %>
			<span class="list_link"><a href="javascript:fnGotoPage(<%= oCEventPrize.StartScrollPage-1 %>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oCEventPrize.StartScrollPage to oCEventPrize.StartScrollPage + oCEventPrize.FScrollCount - 1 %>
			<% if (i > oCEventPrize.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oCEventPrize.FCurrPage) then %>
			<span class="page_link"><font color="red"><b>[<%= i %>]</b></font></span>
			<% else %>
			<a href="javascript:fnGotoPage(<%= i %>)" class="list_link"><font color="#000000">[<%= i %>]</font></a>
			<% end if %>
		<% next %>
		<% if oCEventPrize.HasNextScroll then %>
			<span class="list_link"><a href="javascript:fnGotoPage(<%= i %>)">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<%
set oCEventPrize = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
