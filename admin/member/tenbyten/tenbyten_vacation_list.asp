<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
'
'if (session("ssAdminPsn") = "10") and (session("ssBctId") <> "bseo") and (session("ssBctId") <> "boyishP") then
'	'// CS팀장님 요청사항, 2015-04-08
'	response.write  "권한이 없습니다. - 시스템팀 문의 " ''eastone
'	dbget.close() : response.end
'end if

Dim page, SearchKey, SearchString, part_sn, posit_sn, research
Dim deleteyn
dim lp
dim divcd,statediv

dim userid
dim showonlyavail, iPageSize
dim department_id, inc_subdepartment

page 			= Request("page")
deleteyn 		= Request("deleteyn")
SearchKey 		= Request("SearchKey")
SearchString 	= Request("SearchString")
part_sn 		= Request("part_sn")
posit_sn 		= Request("posit_sn")
research 		= Request("research")
divcd 			= Request("divcd")
statediv		= Request("statediv")
showonlyavail	= Request("showonlyavail")
department_id 	= requestCheckvar(Request("department_id"),10)
inc_subdepartment 	= requestCheckvar(Request("inc_subdepartment"),1)
iPageSize 	= requestCheckvar(Request("pagesize"),10)

if (iPageSize = "") then
	iPageSize = 20
end if
if deleteyn="" and research="" then deleteyn="N"
if page="" then page=1

if (SearchKey = "t.userid") then
	userid = SearchString
end if

'// 로그인정보(등급)에 따라 기본 부서 설정(마스터 이상:2 및 시스템팀:7 제외)
'if Not(session("ssAdminLsn")<=2 or session("ssAdminPsn")=7) then
'	part_sn = session("ssAdminPsn")
'end if

dim oVacation
Set oVacation = new CTenByTenVacation

oVacation.FPagesize = iPageSize
oVacation.FCurrPage = page
oVacation.FRectsearchKey = searchKey
oVacation.FRectsearchString = searchString
oVacation.FRectIsDelete = deleteyn
oVacation.FRectpart_sn = part_sn
oVacation.FRectposit_sn = posit_sn
oVacation.FRectDivCd = divcd
oVacation.FRectStateDiv = statediv
oVacation.FRectShowOnlyAvail = showonlyavail
oVacation.Fdepartment_id 		= department_id
oVacation.Finc_subdepartment 	= inc_subdepartment

oVacation.GetMasterList

%>
<!-- 검색 시작 -->
<script language="javascript">

function ViewDetail(masteridx)
{
	var pop = window.open("/admin/member/tenbyten/tenbyten_vacation_detail_list.asp?masteridx=" + masteridx,"ViewDetail","width=900,height=600,scrollbars=yes");
	pop.focus();
}

function AddItem(userid)
{
	window.open("pop_vacation_modify.asp?userid=" + userid,"popAddIem","width=500,height=600,scrollbars=yes");
}

function AddYearVacationItem(userid)
{
	window.open("pop_vacation_modify.asp?userid=" + userid + "&isyearvacation=Y","popAddIem","width=500,height=600,scrollbars=yes");
}

function AddAllYearVacation(insDivcode)
{
	var strMsg = "";
	if (insDivcode == "R"){
		strMsg = "내년도 연차가 생성됩니다.";
	}

	if (confirm(strMsg + "전체 연차 휴가를 등록하시겠습니까?") == true) {
		document.frmupdate.mode.value="addallyearvacationNew";
		document.frmupdate.insDivcode.value = insDivcode;

		document.frmupdate.submit();
	}
}

function AddAllLongYearVacation()
{
	if (confirm("전체 장기근속 휴가를 등록하시겠습니까?") == true) {
		document.frmupdate.mode.value="addalllongmonthvacation";
		document.frmupdate.submit();
	}
}


function AddReCalVacation()
{
	if (confirm("시급계약직 퇴직자 재정산  연차를 등록하시겠습니까?") == true) {
		document.frmupdate.mode.value="addrecalvacation";
		document.frmupdate.submit();
	}
}
// 페이지 이동
function goPage(pg)
{
	document.frm.page.value=pg;
	document.frm.submit();
}

function jsPartList(empno){
	var winPL = window.open("tenbyten_vacation_part_list.asp?empno=" + empno,"popPL","width=800,height=600,scrollbars=yes");
	winPL.focus();
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
			부서NEW:
			<%= drawSelectBoxDepartment("department_id", department_id) %>
			<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > 하위 부서직원 제외
			&nbsp;
			삭제여부:
			<select name="deleteyn" class="select">
				<option value="">전체</option>
				<option value="N">사용</option>
				<option value="Y">삭제</option>
			</select>
			&nbsp;
			재직여부:
			<select name="statediv" class="select">
				<option value="">전체</option>
				<option value="Y">재직</option>
				<option value="N">퇴사</option>
			</select>
			&nbsp;
			<% if C_ADMIN_AUTH or C_PSMngPart then %>
			직급:
			<%=printPositOptionIN90("posit_sn", posit_sn)%>&nbsp;
			&nbsp;
			<% end if %>
			휴가구분 :
			<select name=divcd class="select">
				<option value="">전체</option>
				<option value="1" <% if (divcd = "1") then %>selected<% end if %>>연차</option>
				<!--
				<option value="2">월차</option>
				-->
				<option value="3" <% if (divcd = "3") then %>selected<% end if %>>포상</option>
				<option value="4" <% if (divcd = "4") then %>selected<% end if %>>위로</option>
				<option value="6" <% if (divcd = "6") then %>selected<% end if %>>경조사</option>
				<option value="5" <% if (divcd = "5") then %>selected<% end if %>>장기</option>
				<option value="7" <% if (divcd = "7") then %>selected<% end if %>>휴일대체</option>
				<option value="8" <% if (divcd = "8") then %>selected<% end if %>>기타휴가</option>
				<option value="9" <% if (divcd = "9") then %>selected<% end if %>>보상휴가</option>
				<option value="A" <% if (divcd = "A") then %>selected<% end if %>>생일휴가</option>
			</select>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			검색:
			<select name="SearchKey" class="select">
				<option value="">::구분::</option>
				<option value="t.userid">아이디</option>
				<option value="t.username">사용자명</option>
				<option value="t.empno">사번</option>
			</select>
			<script language="javascript">
				document.frm.deleteyn.value="<%= deleteyn %>";
				document.frm.SearchKey.value="<%= SearchKey %>";
				document.frm.statediv.value="<%= statediv %>";
			</script>
			<input type="text" class="text" name="SearchString" size="20" value="<%=SearchString%>">	&nbsp;
			<input type="checkbox" name="showonlyavail" value="Y" <% if (showonlyavail = "Y") then %>checked<% end if %> >
			사용가능 휴가만
			&nbsp;
			표시갯수:
			<select class="select" name="pagesize">
				<option value="20" <% if (iPageSize = "20") then %>selected<% end if %> >20 개</option>
				<option value="50" <% if (iPageSize = "50") then %>selected<% end if %> >50 개</option>
				<option value="100" <% if (iPageSize = "100") then %>selected<% end if %> >100 개</option>
				<option value="500" <% if (iPageSize = "500") then %>selected<% end if %> >500 개</option>
			</select>
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
<%
'// 로그인정보(등급)에 따라 기본 부서 설정(관리자 이상:1 및 시스템팀:7 인사총무팀:20 제외)
if (session("ssAdminLsn")<=1 or session("ssAdminPsn")=7 or C_PSMngPart or C_ADMIN_AUTH) then
%>
			[
			관리자 :
			<input type="button" class="button" value="휴가 등록" onClick="javascript:AddItem('<%= userid %>');">
			<input type="button" class="button" value="연차 등록" onClick="javascript:AddYearVacationItem('<%= userid %>');">
			&nbsp;
			<input type="button" class="button" value="전체연차 등록(정규직,년1회)" onClick="javascript:AddAllYearVacation('R');">
			<input type="button" class="button" value="전체연차 등록(계약직,월1회)" onClick="javascript:AddAllYearVacation('P');">
			<input type="button" class="button" value="전체장기근속 등록(정규직,월1회)" onClick="javascript:AddAllLongYearVacation();">
			&nbsp;
			<input type="button" class="button" value="시급/월급계약직 퇴직자 연차재정산" onClick="javascript:AddReCalVacation();">
			]
<% end if %>
<%
'// 시스템팀:7, 30
if (session("ssAdminPsn") = 7) or (session("ssAdminPsn") = 30) then
%>
		<!--	[
			시스템팀 전용 :
			<input type="button" class="button" value="전체연차 등록" onClick="javascript:AddAllYearVacation('');">
			<input type="button" class="button" value="장기근속 등록" onClick="javascript:AddAllLongYearVacation('');">
			]-->
<% end if %>

		</td>
		<td align="right">
			<!-- <img src="/images/icon_excel.gif" border="0"> -->
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 상단 띠 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21">
			검색결과 : <b><%=oVacation.FtotalCount%></b>
			&nbsp;
			페이지 : <b><%= page %> / <%=oVacation.FtotalPage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>idx</td>
		<td>구분</td>
    	<td>사번</td>
		<td width="40">이름</td>
		<td width="70">입사일<br>(정규직)</td>
		<td width="70">실제<br>입사일</td>
		<td width="70">퇴사<br>(예정)일</td>
		<td>부서</td>
		<% if C_ADMIN_AUTH or C_PSMngPart then %><td>직급</td><% end if %>
		<td>직책</td>
		<td width="80">사용가능기간</td>
		<td width="40">총<br>일수</td>
		<td width="40">사용<br>일수</td>
		<td width="40">승인<br>대기</td>

		<td width="50">촉진<br>일수</td>
		<td width="50">정산<br>일수</td>
		<td width="50">퇴사<br>정산</td>

		<td width="50">잔여<br>일수</td>
		<td width="30">사용<br>가능</td>
		<td width="30">삭제<br>여부</td>
		<td>등록자</td>
    </tr>
	<% if oVacation.FResultCount=0 then %>
	<tr height="25">
		<td colspan="21" align="center" bgcolor="#FFFFFF">등록(검색)된 내용이 없습니다.</td>
	</tr>
	<% else %>

	<% for lp=0 to oVacation.FResultCount - 1 %>
	<tr align="center" bgcolor="<% if (oVacation.FitemList(lp).Fdeleteyn="N") and (oVacation.FitemList(lp).IsAvailableVacation="Y") then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>" height="30">
		<td><a href="javascript:ViewDetail(<%=oVacation.FitemList(lp).Fidx%>)"><%= oVacation.FitemList(lp).Fidx %></a></td>
		<td nowrap><%= oVacation.FitemList(lp).GetDivCDStr %></td>
		<td nowrap><a href="javascript:ViewDetail(<%=oVacation.FitemList(lp).Fidx%>)"><%=oVacation.FitemList(lp).Fempno%></a></td>
		<td nowrap><%= oVacation.FitemList(lp).Fusername %></td>

		<td nowrap><%= Left(oVacation.FitemList(lp).Fjoinday, 10) %></td>
		<td nowrap>
			<% if Not IsNull(oVacation.FitemList(lp).Frealjoinday) then %>
				<% if (oVacation.FitemList(lp).Fjoinday <> oVacation.FitemList(lp).Frealjoinday) then %>
					<font color="red"><%= oVacation.FitemList(lp).Frealjoinday %></font>
				<% else %>
					<%= oVacation.FitemList(lp).Frealjoinday %>
				<% end if %>
			<% end if %>
		</td>
		<td nowrap>
			<% if Not IsNull(oVacation.FitemList(lp).Fretireday) then %>
				<%= oVacation.FitemList(lp).Fretireday %>
			<% end if %>
		</td>

		<td><%= oVacation.FitemList(lp).FdepartmentNameFull %></td>
		<% if C_ADMIN_AUTH or C_PSMngPart then %><td><%= oVacation.FitemList(lp).Fposit_name %></td><% end if %>
		<td><%= oVacation.FitemList(lp).Fjob_name %></td>
		<td><%= Left(oVacation.FitemList(lp).Fstartday,10) %> ~ <%= Left(oVacation.FitemList(lp).Fendday,10) %></td>
		<td>
			<a href="javascript:jsPartList('<%=oVacation.FitemList(lp).Fempno%>');"><%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).Ftotalvacationday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %></a>
		</td>
		<td>
			<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).Fusedvacationday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
		</td>
		<td>
			<b>
			<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).Frequestedday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
			</b>
		</td>

		<td>
			<% if (oVacation.FitemList(lp).Fdivcd = "1") or (oVacation.FitemList(lp).Fdivcd = "7") then %>
				<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).FpromotionDay) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
			<% end if %>
		</td>
		<td>
			<% if (oVacation.FitemList(lp).Fdivcd = "1") or (oVacation.FitemList(lp).Fdivcd = "7") then %>
				<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).FjungsanDay) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
			<% end if %>
		</td>
		<td>
			<% if (oVacation.FitemList(lp).Fdivcd = "1") or (oVacation.FitemList(lp).Fdivcd = "7") then %>
				<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).FretireJungsanDay) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
			<% end if %>
		</td>

		<td>
			<%= GetDayOrHourWithPositSN(oVacation.FitemList(lp).Fposit_sn, oVacation.FitemList(lp).GetRemainVacationDay) %> <%= GetDayOrHourNameWithPositSN(oVacation.FitemList(lp).Fposit_sn) %>
		</td>
		<td><%= oVacation.FitemList(lp).IsAvailableVacation %></td>
		<td><%= oVacation.FitemList(lp).Fdeleteyn %></td>
		<td><%= oVacation.FitemList(lp).Fregisterid %></td>
	</tr>
	<% next %>

	<% end if %>
<!-- 메인 목록 끝 -->

<!-- 페이지 시작 -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21" align="center">
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
<form name="frmupdate" method="post" action="domodifyvacation.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="insDivcode" value="">
</form>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
