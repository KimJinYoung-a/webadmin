<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/tenbyten/companyCalendarCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%

dim i, j, k
dim page, reload
dim act_yyyy1, act_yyyy2, act_mm1, act_mm2, act_dd1, act_dd2
dim tmpDate
dim act_fromDate, act_toDate
dim department_id
dim useOnly, myCalOnly


page = requestcheckvar(request("page"),10)
reload = requestcheckvar(request("reload"),10)
department_id = requestcheckvar(request("department_id"),32)
useOnly = requestcheckvar(request("useOnly"),1)
myCalOnly = requestcheckvar(request("myCalOnly"),1)

act_yyyy1   = requestcheckvar(request("act_yyyy1"),4)
act_mm1     = requestcheckvar(request("act_mm1"),4)
act_dd1     = requestcheckvar(request("act_dd1"),4)
act_yyyy2   = requestcheckvar(request("act_yyyy2"),4)
act_mm2     = requestcheckvar(request("act_mm2"),4)
act_dd2     = requestcheckvar(request("act_dd2"),4)

if (reload = "") then
	useOnly = "Y"
	myCalOnly = "Y"
end if

if (act_yyyy1="") then
	act_fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)
	act_toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) + 1), 1)

	act_yyyy1 = Cstr(Year(act_fromDate))
	act_mm1 = Cstr(Month(act_fromDate))
	act_dd1 = Cstr(day(act_fromDate))

	tmpDate = DateAdd("d", -1, act_toDate)
	act_yyyy2 = Cstr(Year(tmpDate))
	act_mm2 = Cstr(Month(tmpDate))
	act_dd2 = Cstr(day(tmpDate))
else
	act_fromDate = DateSerial(act_yyyy1, act_mm1, act_dd1)
	act_toDate = DateSerial(act_yyyy2, act_mm2, act_dd2+1)
end if

''adminuserid=session("ssBctId")

if page = "" then page = 1

dim oCompanyCalendar
set oCompanyCalendar = new CCompanyCalendar
	oCompanyCalendar.FPageSize = 20
	oCompanyCalendar.FCurrPage = page

	if (myCalOnly = "Y") then
		'// 사번
		oCompanyCalendar.FRectEmpNO = session("ssBctSn")
	end if

	oCompanyCalendar.FRectUseYN = useOnly
	oCompanyCalendar.FRectStartDate = act_fromDate
	oCompanyCalendar.FRectEndDate = act_toDate
	oCompanyCalendar.FRectDepartmentID = department_id

	oCompanyCalendar.getCompanyCalendarList()

%>

<script type="text/javascript">

function frmGoto(page) {
	frm.page.value = page;
	frm.submit();
}

function popOpenCalendarItem(idx) {
	var pop = window.open('popCalendarItem.asp?idx='+idx,'popOpenCalendarItem','width=800,height=450,scrollbars=yes,resizable=yes');
	pop.focus();
}

function popOpenCalendar() {
	var pop = window.open('popCompCalendar.asp','popOpenCalendar','width=1200,height=800,scrollbars=yes,resizable=yes');
	pop.focus();
}


</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left" height="30">
		부서 : <%= drawSelectBoxDepartment("department_id", department_id) %>
		&nbsp;
		기간 : <% DrawDateBoxdynamic act_yyyy1, "act_yyyy1", act_yyyy2, "act_yyyy2", act_mm1, "act_mm1", act_mm2, "act_mm2", act_dd1, "act_dd1", act_dd2, "act_dd2" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="30">
		<input type="checkbox" name="myCalOnly" value="Y" <% if (myCalOnly = "Y") then %>checked<% end if %> > 내 일정만 표시
		&nbsp;
		<input type="checkbox" name="useOnly" value="Y" <% if (useOnly = "Y") then %>checked<% end if %> > 사용일정만 표시
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="사내달력보기" onclick="popOpenCalendar();">
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onclick="popOpenCalendarItem('');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oCompanyCalendar.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oCompanyCalendar.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">IDX</td>
	<td>제목</td>
	<td width="150">기간</td>
	<td width="80">우선순위</td>
	<td width="40">사용</td>
	<td width="100">등록자</td>
	<td width="80">등록일</td>
	<td width="80">최종수정</td>
	<td>비고</td>
</tr>
<% if oCompanyCalendar.FresultCount > 0 then %>
	<% for i=0 to oCompanyCalendar.FresultCount-1 %>

	<% if oCompanyCalendar.FItemList(i).FuseYN = "Y" then %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
	<% else %>
		<tr align="center" bgcolor="#E1E1E1" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#E1E1E1';>
	<% end if %>

		<td>
			<%= oCompanyCalendar.FItemList(i).Fidx %>
		</td>
		<td>
			<%= Replace(oCompanyCalendar.FItemList(i).Ftitle, "<", "&lt;") %>
		</td>
		<td>
			<%= Left(oCompanyCalendar.FItemList(i).FstartDate,10) %> ~ <%= Left(oCompanyCalendar.FItemList(i).FendDate, 10) %>
		</td>
		<td>
			<%= oCompanyCalendar.FItemList(i).GetImportantLevelName %>
		</td>
		<td>
			<%= oCompanyCalendar.FItemList(i).FuseYN %>
		</td>
		<td>
			<%= oCompanyCalendar.FItemList(i).Freguserid %>
		</td>
		<td>
			<acronym title="<%= oCompanyCalendar.FItemList(i).Fregdate %>"><%= Left(oCompanyCalendar.FItemList(i).Fregdate,10) %></acronym>
		</td>
		<td>
			<acronym title="<%= oCompanyCalendar.FItemList(i).Flastupdate %>"><%= Left(oCompanyCalendar.FItemList(i).Flastupdate,10) %></acronym>
		</td>

		<td>
			<input type="button" onclick="popOpenCalendarItem('<%= oCompanyCalendar.FItemList(i).Fidx %>'); return false;" value="수정" class="button">
		</td>
	</tr>
	<% next %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oCompanyCalendar.HasPreScroll then %>
				<span class="list_link"><a href="javascript:frmGoto(<%= oCompanyCalendar.StartScrollPage-1 %>);">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oCompanyCalendar.StartScrollPage to oCompanyCalendar.StartScrollPage + oCompanyCalendar.FScrollCount - 1 %>
				<% if (i > oCompanyCalendar.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oCompanyCalendar.FCurrPage) then %>
				<span class="page_link"><font color="red"><b>[<%= i %>]</b></font></span>
				<% else %>
				<a href="javascript:frmGoto(<%= i %>);" class="list_link"><font color="#000000">[<%= i %>]</font></a>
				<% end if %>
			<% next %>
			<% if oCompanyCalendar.HasNextScroll then %>
				<span class="list_link"><a href="javascript:frmGoto(<%= i %>);">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set oCompanyCalendar = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
