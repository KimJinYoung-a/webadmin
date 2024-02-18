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

	page = Request("page")
	deleteyn = Request("deleteyn")
	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")
	part_sn = Request("part_sn")
	research = Request("research")
	if deleteyn="" and research="" then deleteyn="N"
	if page="" then page=1

	'// 로그인정보(등급)에 따라 기본 부서 설정(마스터 이상:2 및 시스템팀:7 제외)
	if Not(session("ssAdminLsn")<=2 or session("ssAdminPsn")=7) then
		part_sn = session("ssAdminPsn")
	end if

	dim oVacation
	Set oVacation = new CTenByTenVacation

	oVacation.FPagesize = 20
	oVacation.FCurrPage = page
	oVacation.FRectsearchKey = searchKey
	oVacation.FRectsearchString = searchString
	oVacation.FRectIsDelete = deleteyn
	oVacation.FRectpart_sn = part_sn

	oVacation.GetMasterList





%>
<!-- 검색 시작 -->
<script language="javascript">

function ViewDetail(masteridx)
{
	location.href = "/admin/member/tenbyten/tenbyten_vacation_detail_list.asp?masteridx=" + masteridx;
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
			<% if session("ssAdminLsn")<=2 then %>
			부서:
			<%=printPartOption("part_sn", part_sn)%>&nbsp;
			<% end if %>
			사용여부:
			<select name="deleteyn" class="select">
				<option value="">전체</option>
				<option value="Y">사용</option>
				<option value="N">삭제</option>
			</select>
			&nbsp;
			검색:
			<select name="SearchKey" class="select">
				<option value="">::구분::</option>
				<option value="t.userid">아이디</option>
				<option value="t.username">사용자명</option>
			</select>
			<script language="javascript">
				document.frm.deleteyn.value="<%= deleteyn %>";
				document.frm.SearchKey.value="<%= SearchKey %>";
			</script>
			<input type="text" class="text" name="SearchString" size="12" value="<%=SearchString%>">

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
			<input type="button" class="button" value="신규등록" onClick="javascript:AddItem('');">
		</td>
		<td align="right">
			<img src="/images/icon_excel.gif" border="0">
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
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">idx</td>
		<td width="50">구분</td>
    	<td width="100">아이디</td>
		<td width="60">이름</td>
		<td width="190">부서</td>
		<td width="50">직급</td>
		<td width="70">직책</td>
		<td width="150">사용가능기간</td>
		<td width="60">총일수</td>
		<td width="60">사용일수</td>
		<td width="60">승인대기</td>
		<td width="60">잔여일수</td>
		<td width="60">사용가능</td>
		<td width="60">삭제여부</td>
    </tr>
	<% if oVacation.FResultCount=0 then %>
	<tr>
		<td colspan="15" align="center" bgcolor="#FFFFFF">등록(검색)된 내용이 없습니다.</td>
	</tr>
	<% else %>

	<% for lp=0 to oVacation.FResultCount - 1 %>
	<tr align="center" bgcolor="<% if (oVacation.FitemList(lp).Fdeleteyn="N") and (oVacation.FitemList(lp).IsAvailableVacation="Y") then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
		<td><%= oVacation.FitemList(lp).Fidx %></td>
		<td><%= oVacation.FitemList(lp).GetDivCDStr %></td>
		<td><a href="javascript:ViewDetail(<%=oVacation.FitemList(lp).Fidx%>)"><%=oVacation.FitemList(lp).Fuserid%></a></td>
		<td><%= oVacation.FitemList(lp).Fusername %></td>
		<td><%= oVacation.FitemList(lp).Fpart_name %></td>
		<td><%= oVacation.FitemList(lp).Fposit_name %></td>
		<td><%= oVacation.FitemList(lp).Fjob_name %></td>
		<td><%= Left(oVacation.FitemList(lp).Fstartday,10) %>-<%= Left(oVacation.FitemList(lp).Fendday,10) %></td>
		<td><%= oVacation.FitemList(lp).Ftotalvacationday %></td>
		<td><%= oVacation.FitemList(lp).Fusedvacationday %></td>
		<td><%= oVacation.FitemList(lp).Frequestedday %></td>
		<td><%= (oVacation.FitemList(lp).Ftotalvacationday - (oVacation.FitemList(lp).Fusedvacationday + oVacation.FitemList(lp).Frequestedday)) %></td>
		<td><%= oVacation.FitemList(lp).IsAvailableVacation %></td>
		<td><%= oVacation.FitemList(lp).Fdeleteyn %></td>
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