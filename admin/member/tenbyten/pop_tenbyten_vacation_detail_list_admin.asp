<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 휴가신청
' History : 서동석 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<%
	Dim page, masteridx
	dim i
	dim part_sn
	dim userid
dim useridDb, usernameDb, part_nameDb, posit_nameDb, GetDivCDStrDb, totalvacationdayDb, startdayDb, enddayDb, usedvacationday
dim requesteddayDb, IsAvailableVacationDb, deleteynDb

	page = Request("page")
	masteridx = Request("masteridx")
	part_sn = Request("part_sn")

	if page="" then page=1

	userid = session("ssBctId")

	'// 직책 팀장이상. 또는 시스템팀을 제외하고 본인것만
	if Not((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or (session("ssAdminPOsn") = "5") or session("ssAdminPsn")=7 or C_ADMIN_AUTH) then
		userid = session("ssBctId")
		part_sn = session("ssAdminPsn")
	end if

	'// 로그인정보(등급)에 따라 기본 부서 설정(마스터 이상:2 및 시스템팀:7 제외)
	'if Not (session("ssAdminLsn")<=2 or session("ssAdminPsn")=7) then
	'	part_sn = session("ssAdminPsn")
	'end if

	dim oVacation
	Set oVacation = new CTenByTenVacation

	oVacation.FRectMasterIdx = masteridx
	'oVacation.FRectpart_sn = part_sn

	'// 직책 파트장이상. 또는 시스템팀을 제외하고 본인것만
	if Not((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or (session("ssAdminPOsn") = "5") or session("ssAdminPsn")=7 or C_ADMIN_AUTH) then
		oVacation.FRectsearchKey = " t.userid "
		oVacation.FRectsearchString = session("ssBctId")
	end if

	if masteridx<>"" and not(isnull(masteridx)) then
		oVacation.GetMasterOne
	end if

	if masteridx<>"" and not(isnull(masteridx)) then
		oVacation.GetDetailList
	end if

if oVacation.FResultCount>0 then
	useridDb=oVacation.FItemOne.Fuserid
	usernameDb=oVacation.FItemOne.Fusername
	part_nameDb=oVacation.FItemOne.Fpart_name
	posit_nameDb=oVacation.FItemOne.Fposit_name
	GetDivCDStrDb=oVacation.FItemOne.GetDivCDStr
	totalvacationdayDb=oVacation.FItemOne.Ftotalvacationday
	startdayDb=oVacation.FItemOne.Fstartday
	enddayDb=oVacation.FItemOne.Fendday
	usedvacationday=oVacation.FItemOne.Fusedvacationday
	requesteddayDb=oVacation.FItemOne.Frequestedday
	IsAvailableVacationDb=oVacation.FItemOne.IsAvailableVacation
	deleteynDb=oVacation.FItemOne.Fdeleteyn
end if

%>
<!-- 검색 시작 -->
<script language="javascript">
<!--
	function AddItem()
	{
<% if (useridDb = userid) then %>
		window.open("pop_vacation_detail_modify.asp?masteridx=<%= masteridx %>","popAddIem","width=500,height=600,scrollbars=yes");
<% else %>
		alert("휴가는 자기것만 신청할 수 있습니다.");
<% end if %>
	}

	function ViewList(part_sn)
	{
		location.href = "/admin/member/tenbyten/pop_tenbyten_vacation_list_admin.asp?part_sn=" + part_sn;
	}
	function ViewCalendar()
	{
		window.open("/admin/member/tenbyten/pop_vacation_calendar.asp","popAddIem","width=800,height=650,scrollbars=yes");
	}
	function SubmitAllow(masteridx, detailidx)
	{
		var frm = document.frmmodify;

		if (confirm("승인하시겠습니까?") == true) {
			frm.mode.value = "allowdetail";
			frm.masteridx.value = masteridx;
			frm.detailidx.value = detailidx;

			frm.submit();
		}
	}

	function SubmitDeny(masteridx, detailidx)
	{
		var frm = document.frmmodify;

		if (confirm("거절하시겠습니까?") == true) {
			frm.mode.value = "denydetail";
			frm.masteridx.value = masteridx;
			frm.detailidx.value = detailidx;

			frm.submit();
		}
	}

	function SubmitDelete(masteridx, detailidx)
	{
		var frm = document.frmmodify;

		if (confirm("삭제하시겠습니까?") == true) {
			frm.mode.value = "deletedetail";
			frm.masteridx.value = masteridx;
			frm.detailidx.value = detailidx;

			frm.submit();
		}
	}




	// 페이지 이동
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}

	//상세 내용보기
	function jsDetailView(idx){
		var winDetail = window.open("pop_vacation_detail_view.asp?detailidx="+idx,"popDetail","width=500,height=300,scrollbars=yes");
		winDetail.focus();
	}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="part_sn" value="<%= part_sn %>">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">이름(아이디)</td>
		<td align="left">
			<%= usernameDb %>(<%= useridDb %>)
		</td>
		<!--
		<td width="100" bgcolor="<%= adminColor("gray") %>">부서 / 직급</td>
		<td align="left">
			<%= part_nameDb %> / <%= posit_nameDb %>
		</td>
		-->
		<td width="100" bgcolor="<%= adminColor("gray") %>">부서</td>
		<td align="left">
			<%= part_nameDb %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">구분</td>
		<td align="left">
			<%= GetDivCDStrDb %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">총일수</td>
		<td align="left">
			<%= totalvacationdayDb %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">사용가능기간</td>
		<td align="left">
			<%= Left(startdayDb,10) %> - <%= Left(enddayDb,10) %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">사용일수</td>
		<td align="left">
			<%= usedvacationday %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">승인대기</td>
		<td align="left">
			<%= requesteddayDb %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">잔여일수</td>
		<td align="left">
			<b><%= (totalvacationdayDb - (usedvacationday + requesteddayDb)) %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">사용가능</td>
		<td align="left">
			<b><%= IsAvailableVacationDb %></b>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">삭제여부</td>
		<td align="left">
			<%= deleteynDb %>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

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
		<td width="50">상태</td>
		<td width="150">기간</td>
		<td width="60">신청일수</td>
		<td width="100">등록자</td>
		<td width="100">처리자</td>
		<td>비고</td>
    </tr>
	<% if oVacation.FResultCount=0 then %>
	<tr height=30>
		<td colspan="15" align="center" bgcolor="#FFFFFF">등록(검색)된 내용이 없습니다.</td>
	</tr>
	<% else %>
		<% for i=0 to oVacation.FResultCount - 1 %>
	<tr height=30 align="center" bgcolor="<% if (oVacation.FitemList(i).Fdeleteyn="N") then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
		<td><%=oVacation.FitemList(i).Fidx%></td>
		<td><%= oVacation.FitemList(i).GetStateDivCDStr %></td>
		<td><%= Left(oVacation.FitemList(i).Fstartday,10) %> - <%= Left(oVacation.FitemList(i).Fendday,10) %></td>
		<td><%= oVacation.FitemList(i).Ftotalday %><% If oVacation.FitemList(i).Ftotalday = "0.5" Then Response.Write CHKIIF(oVacation.FItemList(i).Fhalfgubun="am","[오전]","[오후]") End If %></td>
		<td><%= oVacation.FitemList(i).Fregistername %></td>
		<td><%= oVacation.FitemList(i).Fapprovername %></td>
		<td>
<%
'// '// 직책 팀장이상. 또는 시스템팀
if ((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or (session("ssAdminPOsn") = "5") or session("ssAdminPsn")=7 or C_ADMIN_AUTH) then
%>
			<% if (oVacation.FitemList(i).Fdeleteyn="N") and (oVacation.FitemList(i).Fstatedivcd="R") then %>
			<input type=button value=" 승 인 " class="button" onclick="SubmitAllow(<%= masteridx %>, <%=oVacation.FitemList(i).Fidx%>)"> <input type=button value=" 거 절 " class="button" onclick="SubmitDeny(<%= masteridx %>, <%=oVacation.FitemList(i).Fidx%>)">
			<% end if %>
			<% if (oVacation.FitemList(i).Fdeleteyn="N") then %>
			<input type=button value=" 삭 제 " class="button" onclick="SubmitDelete(<%= masteridx %>, <%=oVacation.FitemList(i).Fidx%>)">
			<% end if %>
<% end if %>
		</td>
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

				for i=0 + oVacation.StartScrollPage to oVacation.FScrollCount + oVacation.StartScrollPage - 1

					if i>oVacation.FTotalpage then Exit for

					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
					end if

				next

				if oVacation.HasNextScroll then
					Response.Write "<a href='javascript:goPage(" & i & ")'>[next]</a>"
				else
					Response.Write "[next]"
				end if
			%>
		</td>
	</tr>
</table>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="리스트" onClick="ViewList('<%= part_sn %>');">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- 액션 끝 -->


<form name=frmmodify method=post action="domodifyvacation.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="masteridx" value="">
	<input type="hidden" name="detailidx" value="">
</form>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
