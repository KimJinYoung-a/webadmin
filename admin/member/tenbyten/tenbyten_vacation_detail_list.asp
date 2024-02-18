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

Dim page, masteridx
dim i
dim part_sn
dim userid

page = Request("page")
masteridx = Request("masteridx")

if page="" then page=1

userid = session("ssBctId")

'// 로그인정보(등급)에 따라 기본 부서 설정(마스터 이상:2 및 시스템팀:7 제외)
'if Not (session("ssAdminLsn")<=2 or session("ssAdminPsn")=7) then
'	part_sn = session("ssAdminPsn")
'end if

dim oVacation
Set oVacation = new CTenByTenVacation

oVacation.FRectMasterIdx = masteridx
oVacation.FRectpart_sn = part_sn

oVacation.GetMasterOne

oVacation.FPageSize = 40

oVacation.GetDetailList

%>
<!-- 검색 시작 -->
<script language="javascript">
<!--

function AddItem()
{
	window.open("pop_vacation_detail_modify.asp?masteridx=<%= masteridx %>","popAddIem","width=500,height=600,scrollbars=yes");
}

function ViewList()
{
	location.href = "/admin/member/tenbyten/tenbyten_vacation_list.asp";
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
		frm.comment.value = document.frmDetail.comment.value;

		frm.submit();
	}
}

function SubmitDeleteMaster(masteridx)
{
	var frm = document.frmmodify;

	if (confirm("삭제하시겠습니까?") == true) {
		frm.mode.value = "deletemaster";
		frm.masteridx.value = masteridx;

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
		frm.comment.value = document.frmDetail.comment.value;

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

function SubmitModify(masteridx)
{
	var frmmaster = document.frm;
	var frm = document.frmmodify;

	if (jsCheckIsDigit(frmmaster.totalvacationday) != true) {
		alert("숫자만 입력가능합니다.");
		frmmaster.totalvacationday.focus();
		return;
	}

	if (jsCheckIsDigit(frmmaster.promotionDay) != true) {
		alert("숫자만 입력가능합니다.");
		frmmaster.promotionDay.focus();
		return;
	}

	if (jsCheckIsDigit(frmmaster.jungsanDay) != true) {
		alert("숫자만 입력가능합니다.");
		frmmaster.jungsanDay.focus();
		return;
	}

	if (jsCheckIsDigit(frmmaster.retireJungsanDay) != true) {
		alert("숫자만 입력가능합니다.");
		frmmaster.retireJungsanDay.focus();
		return;
	}

	if (confirm("변경하시겠습니까?") == true) {
		frm.mode.value = "modifymaster";
		frm.divcd.value = frmmaster.divcd.value;

		<% if (oVacation.FItemOne.Fposit_sn = 13) then %>
			// 시급계약직은 시간을 일자로 변경해준다.
			// 1일은 8시간, 한시간은 0.125(= 1/8)
			frmmaster.totalvacationday.value = frmmaster.totalvacationday.value * 1.0 * 0.125;
			frmmaster.promotionDay.value = frmmaster.promotionDay.value * 1.0 * 0.125;
			frmmaster.jungsanDay.value = frmmaster.jungsanDay.value * 1.0 * 0.125;
			frmmaster.retireJungsanDay.value = frmmaster.retireJungsanDay.value * 1.0 * 0.125;
		<% end if %>

		frm.totalvacationday.value = frmmaster.totalvacationday.value;
		frm.promotionDay.value = frmmaster.promotionDay.value;
		frm.jungsanDay.value = frmmaster.jungsanDay.value;
		frm.retireJungsanDay.value = frmmaster.retireJungsanDay.value;
		frm.startday.value = frmmaster.startday.value;
		frm.endday.value = frmmaster.endday.value;
		frm.comment.value = frmmaster.comment.value;

		frm.masteridx.value = masteridx;

		frm.submit();
	}
}

function jsCheckIsDigit(obj) {
	if ((obj.value == "") || (obj.value*0 != 0)) {
		return false;
	}

	return true;
}

// 페이지 이동
function goPage(pg)
{
	document.frm.page.value=pg;
	document.frm.submit();
}

function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

//-->
</script>
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">이름(아이디)</td>
		<td align="left">
			<%= oVacation.FItemOne.Fusername %>(<% if (oVacation.FItemOne.Fuserid = "") then response.write oVacation.FItemOne.Fempno else response.write oVacation.FItemOne.Fuserid end if %>)
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">부서</td>
		<td align="left">
			<%= oVacation.FItemOne.Fpart_name %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">구분</td>
		<td align="left">
			<select class="select" name=divcd>
				<option value="1" <% if (oVacation.FItemOne.Fdivcd = "1") then %>selected<% end if %> >연차</option>
				<!--
				<option value="2" <% if (oVacation.FItemOne.Fdivcd = "2") then %>selected<% end if %> >월차</option>
				-->
				<option value="3" <% if (oVacation.FItemOne.Fdivcd = "3") then %>selected<% end if %> >포상</option>
				<option value="4" <% if (oVacation.FItemOne.Fdivcd = "4") then %>selected<% end if %> >위로</option>
				<option value="6" <% if (oVacation.FItemOne.Fdivcd = "6") then %>selected<% end if %> >경조사</option>
				<option value="5" <% if (oVacation.FItemOne.Fdivcd = "5") then %>selected<% end if %> >장기</option>
				<option value="7" <% if (oVacation.FItemOne.Fdivcd = "7") then %>selected<% end if %> >휴일대체</option>
				<option value="8" <% if (oVacation.FItemOne.Fdivcd = "8") then %>selected<% end if %> >기타휴가</option>
				<option value="9" <% if (oVacation.FItemOne.Fdivcd = "9") then %>selected<% end if %> >보상휴가</option>
				<option value="9" <% if (oVacation.FItemOne.Fdivcd = "A") then %>selected<% end if %> >생일휴가</option>
			</select>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">총일수</td>
		<td align="left">
			<% if (oVacation.FItemOne.Fposit_sn <> 13) then %>
				<input type="text" class="text" name="totalvacationday" size="2" value="<%= oVacation.FItemOne.Ftotalvacationday %>">
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% else %>
				<input type="text" class="text" name="totalvacationday" size="2" value="<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FItemOne.Ftotalvacationday) %>">
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
				(시급계약직)
			<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">사용가능기간</td>
		<td align="left">
    		<input type="text" name="startday" class="text_ro" size="11" maxlength="10" value="<%= Left(oVacation.FItemOne.Fstartday,10) %>" onClick="jsPopCal('frm','startday');" style="cursor:hand;">
    		-
    		<input type="text" name="endday" class="text_ro" size="11" maxlength="10" value="<%= Left(oVacation.FItemOne.Fendday,10) %>" onClick="jsPopCal('frm','endday');" style="cursor:hand;">
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">사용일수</td>
		<td align="left">
			<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FItemOne.Fusedvacationday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">승인대기</td>
		<td align="left">
			<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FItemOne.Frequestedday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">촉진일수</td>
		<td align="left">
			<% if (oVacation.FItemOne.Fposit_sn <> 13) then %>
				<input type="text" class="text<% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>_ro<% end if %>" name="promotionDay" size="2" value="<%= oVacation.FItemOne.FpromotionDay %>" <% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>readonly<% end if %> >
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% else %>
				<input type="text" class="text<% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>_ro<% end if %>" name="promotionDay" size="2" value="<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FItemOne.FpromotionDay) %>" <% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>readonly<% end if %> >
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">정산일수</td>
		<td align="left">
			<% if (oVacation.FItemOne.Fposit_sn <> 13) then %>
				<input type="text" class="text<% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>_ro<% end if %>" name="jungsanDay" size="2" value="<%= oVacation.FItemOne.FjungsanDay %>" <% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>readonly<% end if %> >
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% else %>
				<input type="text" class="text<% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>_ro<% end if %>" name="jungsanDay" size="2" value="<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FItemOne.FjungsanDay) %>" <% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>readonly<% end if %> >
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% end if %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">청산일수</td>
		<td align="left">
			<% if (oVacation.FItemOne.Fposit_sn <> 13) then %>
				<input type="text" class="text<% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>_ro<% end if %>" name="retireJungsanDay" size="2" value="<%= oVacation.FItemOne.FretireJungsanDay %>" <% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>readonly<% end if %> >
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% else %>
				<input type="text" class="text<% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>_ro<% end if %>" name="retireJungsanDay" size="2" value="<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FItemOne.FretireJungsanDay) %>" <% if (oVacation.FItemOne.Fdivcd <> "1") and (oVacation.FItemOne.Fdivcd <> "7") then %>readonly<% end if %> >
				<%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">잔여일수</td>
		<td align="left">
			<b> 
			<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, (oVacation.FItemOne.GetRemainVacationDay)) %> <%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
			</b>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">관리자 코멘트</td>
		<td align="left">
			<input type="text" name="comment" value="<%=replace(oVacation.FItemOne.Fcomment ,"""","&quot;")%>" class="text" style="width:96%;" />
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">사용가능</td>
		<td align="left">
			<b><%= oVacation.FItemOne.IsAvailableVacation %></b>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">삭제여부</td>
		<td align="left">
			<%= oVacation.FItemOne.Fdeleteyn %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="40" width="100" bgcolor="<%= adminColor("gray") %>">관리자기능</td>
		<td align="center" colspan="3">
			<%
			'// 로그인정보(등급)에 따라 기본 부서 설정(파트선임 이상:3 및 시스템팀:7 경영관리팀:8 제외)
			if (session("ssAdminLsn")<=3 or session("ssAdminPsn")=7 or session("ssAdminPsn")=8) or C_PSMngPart or C_ADMIN_AUTH then
			%>
				<input type="button" class="button" value="변경하기" onClick="javascript:SubmitModify(<%= masteridx %>);">
			<% end if %>
		</td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
<%
'// 로그인정보(등급)에 따라 기본 부서 설정(파트선임 이상:3 및 시스템팀:7 경영관리팀:8 제외)
if (session("ssAdminLsn")<=3 or session("ssAdminPsn")=7 or session("ssAdminPsn")=8 or C_PSMngPart or C_ADMIN_AUTH) then
%>
			<input type="button" class="button" value="관리자휴가내역등록" onClick="javascript:AddItem('');">
<% end if %>
			<input type="button" class="button" value="휴가달력보기" onClick="javascript:ViewCalendar('');">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 상단 띠 시작 -->
<form name="frmDetail" method="get" action="">
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
		<td width="50">상태</td>
		<td width="150">기간</td>
		<td width="60">신청일수</td>
		<td width="60">삭제여부</td>
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
		<td><a href="javascript:ModiItem('<%=oVacation.FitemList(i).Fidx%>')"><%=oVacation.FitemList(i).Fidx%></a></td>
		<td><%= oVacation.FitemList(i).GetStateDivCDStr %></td>
		<td><%= Left(oVacation.FitemList(i).Fstartday,10) %> - <%= Left(oVacation.FitemList(i).Fendday,10) %></td>
		<td>
			<%= GetDayOrHourWithPositSN(oVacation.FItemOne.Fposit_sn, oVacation.FitemList(i).Ftotalday) %> <%= GetDayOrHourNameWithPositSN(oVacation.FItemOne.Fposit_sn) %>
		</td>
		<td><%= oVacation.FitemList(i).Fdeleteyn %></td>
		<td><%= oVacation.FitemList(i).Fregistername %></td>
		<td><%= oVacation.FitemList(i).Fapprovername %></td>
		<td>
<%
'// 로그인정보(등급)에 따라 기본 부서 설정(파트선임 이상:3 및 시스템팀:7 경영관리팀:8 제외)
if (session("ssAdminLsn")<=3 or session("ssAdminPsn")=7 or session("ssAdminPsn")=8 or C_PSMngPart or C_ADMIN_AUTH) then
%>
			<% if (oVacation.FitemList(i).Fdeleteyn="N") and (oVacation.FitemList(i).Fstatedivcd="R") then %>
			코멘트: <input type="text" name="comment" class="text" value="" style="width:70%" /><br />
			<input type=button value=" 승 인 " class="button" onclick="SubmitAllow(<%= masteridx %>, <%=oVacation.FitemList(i).Fidx%>)"> <input type=button value=" 거 절 " class="button" onclick="SubmitDeny(<%= masteridx %>, <%=oVacation.FitemList(i).Fidx%>)">
			<% elseif oVacation.FitemList(i).Fcomment<>"" then %>
			코멘트: <%=oVacation.FitemList(i).Fcomment%><br />
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
</form>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value=" 리스트 " onClick="ViewList();">
<%
'// 로그인정보(등급)에 따라 기본 부서 설정(파트선임 이상:3 및 시스템팀:7 경영관리팀:8 제외)
if (session("ssAdminLsn")<=3 or session("ssAdminPsn")=7 or session("ssAdminPsn")=8 or C_PSMngPart or C_ADMIN_AUTH) then
%>
			<input type="button" class="button" value=" 휴가삭제 " onClick="SubmitDeleteMaster(<%= masteridx %>);">
<% end if %>
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<form name="frmmodify" method="post" action="domodifyvacation.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="" />
	<input type="hidden" name="masteridx" value="" />
	<input type="hidden" name="detailidx" value="" />
	<input type="hidden" name="divcd" value="" />
	<input type="hidden" name="totalvacationday" value="" />
	<input type="hidden" name="promotionDay" value="" />
	<input type="hidden" name="jungsanDay" value="" />
	<input type="hidden" name="retireJungsanDay" value="" />
	<input type="hidden" name="startday" value="" />
	<input type="hidden" name="endday" value="" />
	<input type="hidden" name="comment" value="" />
</form>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
