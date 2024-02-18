<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<%
	Dim page, SearchKey, SearchString, part_sn, research
	Dim deleteyn
	dim lp
	dim needapprove

	page = Request("page")
	deleteyn = Request("deleteyn")
	needapprove = Request("needapprove")
	'SearchKey = Request("SearchKey")
	'SearchString = Request("SearchString")
	part_sn = Request("part_sn")
	research = Request("research")
	'deleteyn="N"
	if deleteyn="" and research="" then deleteyn="N"
	if needapprove="" and research="" then needapprove="Y"
	if page="" then page=1

	'session("ssAdminPOsn") 직책
	'// 직책 팀장이상. 또는 시스템팀을 제외하고 소속팀 지정
	if Not((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or session("ssAdminPsn")=7) then
		part_sn = session("ssAdminPsn")
	end if

	dim oVacation
	Set oVacation = new CTenByTenVacation

	oVacation.FPagesize = 20
	oVacation.FCurrPage = page
	'oVacation.FRectsearchKey = " t.userid "
	'oVacation.FRectsearchString = session("ssBctId")
	oVacation.FRectIsDelete = deleteyn
	oVacation.FRectpart_sn = part_sn
	oVacation.FRectNeedApprove = needapprove

	oVacation.GetMasterList





%>
<!-- 검색 시작 -->
<script language="javascript">

function ViewDetail(masteridx, part_sn)
{
	location.href = "/admin/member/tenbyten/pop_tenbyten_vacation_detail_list_admin.asp?masteridx=" + masteridx + "&part_sn=" + part_sn;
}

function ViewCalendar(part_sn)
{
	window.open("/admin/member/tenbyten/pop_vacation_calendar.asp?part_sn=" + part_sn,"popAddIem","width=800,height=650,scrollbars=yes");
}

// 페이지 이동
function goPage(pg)
{
	document.frm.page.value=pg;
	document.frm.submit();
}

</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			부서:
			<%=printPartOption("part_sn", part_sn)%>&nbsp;
			삭제여부:
			<select name="deleteyn" class="select">
				<option value="">전체</option>
				<option value="N" <% if (deleteyn = "N") then %>selected<% end if %>>정상</option>
				<option value="Y" <% if (deleteyn = "Y") then %>selected<% end if %>>삭제</option>
			</select>
			&nbsp;
			검색범위:
			<select name="needapprove" class="select">
				<option value="">전체</option>
				<option value="Y" <% if (needapprove = "Y") then %>selected<% end if %>>승인대기만</option>
			</select>
			&nbsp;
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
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
			<input type="button" class="button" value="휴가달력보기" onClick="javascript:ViewCalendar('<%= part_sn %>');">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 상단 띠 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%=oVacation.FtotalCount%></b>
			&nbsp;
			페이지 : <b><%= page %> / <%=oVacation.FtotalPage%></b>
		</td>
	</tr>
	<tr height=30 align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">idx</td>
		<td width="50">이름</td>
		<!--<td width="50">직책</td>-->
		<td width="50">구분</td>
		<td>사용가능기간</td>
		<td width="60">총일수</td>
		<td width="60">사용일수</td>
		<td width="60"><b>승인대기</b></td>
		<td width="60">잔여일수</td>
		<td width="60">사용가능</td>
		<!-- td width="60">만료일수</td -->
		<td width="60">등록자</td>
    </tr>
	<% if oVacation.FResultCount=0 then %>
	<tr height=30>
		<td colspan="15" align="center" bgcolor="#FFFFFF">등록(검색)된 내용이 없습니다.</td>
	</tr>
	<% else %>

	<% for lp=0 to oVacation.FResultCount - 1 %>
	<tr height=30 align="center" bgcolor="<% if (oVacation.FitemList(lp).Fdeleteyn="N") and (not oVacation.FitemList(lp).IsExpiredVacation) then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
		<td><%= oVacation.FitemList(lp).Fidx %></td>
		<td><%= oVacation.FitemList(lp).Fusername %></td>
		<!--<td><%= oVacation.FitemList(lp).Fposit_name %></td>-->
		<td><%= oVacation.FitemList(lp).GetDivCDStr %></td>
		<td><a href="javascript:ViewDetail(<%=oVacation.FitemList(lp).Fidx%>, '<%= part_sn %>')"><%= Left(oVacation.FitemList(lp).Fstartday,10) %>-<%= Left(oVacation.FitemList(lp).Fendday,10) %></a></td>
		<td><%= oVacation.FitemList(lp).Ftotalvacationday %></td>
		<td><%= oVacation.FitemList(lp).Fusedvacationday %></td>
		<td><b><%= oVacation.FitemList(lp).Frequestedday %></b></td>
		<td>
			<% if (oVacation.FitemList(lp).IsAvailableVacation = "Y") then %>
			<%= (oVacation.FitemList(lp).Ftotalvacationday - (oVacation.FitemList(lp).Fusedvacationday + oVacation.FitemList(lp).Frequestedday)) %>
			<% else %>
			0
			<% end if %>


		</td>
		<td>
		    <% if (oVacation.FitemList(lp).Fdeleteyn="N") and (not oVacation.FitemList(lp).IsExpiredVacation) then  %>
		    Y
		    <% else %>
		    N
		    <% end if %>
		    <!--
		    <%= oVacation.FitemList(lp).IsAvailableVacation %>
		    -->        
		</td>
		<!-- td>
			<% if (oVacation.FitemList(lp).IsAvailableVacation = "Y") then %>
			0
			<% else %>
			<%= (oVacation.FitemList(lp).Ftotalvacationday - oVacation.FitemList(lp).Fusedvacationday) %>
			<% end if %>
		</td-->
		<td><%= oVacation.FitemList(lp).Fregisterid %></td>
	</tr>
	<% next %>

	<% end if %>
<!-- 메인 목록 끝 -->

<!-- 페이지 시작 -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<%
				if oVacation.HasPreScroll then
					Response.Write "<a href='javascript:goPage(" & oVacation.StartScrollPage-1 & ")'>[pre]</a>"
				else
					Response.Write "[pre]"
				end if

				for lp=0 + oVacation.StartScrollPage to oVacation.FScrollCount + oVacation.StartScrollPage - 1

					if lp>oVacation.FTotalpage then Exit for

					if CStr(page)=CStr(lp) then
						Response.Write " <font color='red'>[" & lp & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
					end if

				next

				if oVacation.HasNextScroll then
					Response.Write "<a href='javascript:goPage(" & lp & ")'>[next]</a>"
				else
					Response.Write "[next]"
				end if
			%>
		</td>
	</tr>
</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->