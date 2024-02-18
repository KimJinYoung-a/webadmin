<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  부서정보
' History : 2011.1.19 정윤정 생성
'			2011.12.16 한용민 수정
'           2018.03.30 허진원 - 직급 선택 표시
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
Dim page, SearchKey, SearchString, isUsing, part_sn, research, orderby, ilevel_sn, iTotCnt,iPageSize, iTotalPage, isdispmember
Dim job_sn, posit_sn, continuous_service_year, employeeonly, nodepartonly, criticinfouser, workdaycheck, yyyy1, yyyy2, mm1, mm2, dd1, dd2
dim fromDate, toDate, department_id, inc_subdepartment, rank_sn, lv1customerYN, lv2partnerYN, lv3InternalYN
workdaycheck = requestcheckvar(request("workdaycheck"),1)
lv1customerYN 	= requestCheckvar(request("lv1customerYN"),1)
lv2partnerYN 	= requestCheckvar(request("lv2partnerYN"),1)
lv3InternalYN 	= requestCheckvar(request("lv3InternalYN"),1)
yyyy1 = requestcheckvar(request("yyyy1"),4)
yyyy2 = requestcheckvar(request("yyyy2"),4)
mm1	  = requestcheckvar(request("mm1"),2)
mm2	  = requestcheckvar(request("mm2"),2)
dd1	  = requestcheckvar(request("dd1"),2)
dd2	  = requestcheckvar(request("dd2"),2)

iPageSize	  = requestcheckvar(request("pagesize"),10)

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(1)
fromDate = CStr(DateSerial(yyyy1, mm1, dd1))

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()) + 1)
	dd2 = Cstr(1)

	toDate = CStr(DateSerial(yyyy2, mm2, 0))

	yyyy2 = CStr(Year(toDate))
	mm2 = CStr(Month(toDate))
	dd2 = CStr(Day(toDate))
end if
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if (iPageSize = "") then
	iPageSize = 20
end if

page = requestCheckvar(Request("page"),10)
isUsing = requestCheckvar(Request("isUsing"),1)
SearchKey = requestCheckvar(Request("SearchKey"),1)
SearchString = requestCheckvar(Request("SearchString"),32)
part_sn = requestCheckvar(Request("part_sn"),10)
job_sn = requestCheckvar(Request("job_sn"),10)
posit_sn = requestCheckvar(Request("posit_sn"),10)

research = requestCheckvar(Request("research"),2)

orderby = requestCheckvar(Request("orderby"),1)

department_id = requestCheckvar(Request("department_id"),10)
inc_subdepartment = requestCheckvar(Request("inc_subdepartment"),1)
nodepartonly = requestCheckvar(Request("nodepartonly"),1)

criticinfouser = requestCheckvar(Request("criticinfouser"),10)
rank_sn = requestCheckvar(Request("rank_sn"),2)
ilevel_sn = requestCheckvar(Request("ilevel_sn"),10)

if isUsing="" and research="" then isUsing="Y"
if page="" then page=1

'// 로그인정보(등급)에 따라 기본 부서 설정(마스터 이상:2 및 시스템팀:7 제외)
'SCM 메뉴권한 설정에서 제어한다.
'if Not(session("ssAdminLsn")<=2 or session("ssAdminPsn")=7) then
'	part_sn = session("ssAdminPsn")
'end if

IF application("Svr_Info")="Dev" THEN
	isdispmember = true
else
	' ISMS 심사로 인해 개인정보 접근권한 생성/수정/변경 특정사람만 보이게(한용민,허진원,이문재)	' 2020.10.12 한용민
	if C_privacyadminuser or C_PSMngPart then
		isdispmember = true
	else
		isdispmember = false
	end if
end if

'// 내용 접수
dim oMember, arrList,intLoop
Set oMember = new CTenByTenMember

oMember.FPagesize 	= iPageSize
oMember.FCurrPage 	= page
oMember.FSearchType 	= searchKey
oMember.FSearchText 	= searchString
oMember.Fstatediv 		= isUsing
oMember.Fpart_sn 		= part_sn
oMember.Fjob_sn 		= job_sn
oMember.Fposit_sn 		= posit_sn

oMember.Frank_sn		= rank_sn
oMember.Fdepartment_id 		= department_id
oMember.Finc_subdepartment 	= inc_subdepartment
oMember.FRectNoDepartOnly 	= nodepartonly

oMember.FRectCriticInfoUser 	= criticinfouser
oMember.Flevel_sn = ilevel_sn
oMember.Forderby 		= orderby

if (workdaycheck = "Y") then
	oMember.FStartDate		= fromDate
	oMember.FEndDate		= toDate
end if
oMember.Frectlv1customerYN = lv1customerYN
oMember.Frectlv2partnerYN = lv2partnerYN
oMember.Frectlv3InternalYN = lv3InternalYN
arrList = oMember.fnGetMemberList
iTotCnt = oMember.FTotCnt
set oMember = nothing

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

Dim oAddLevel
set oAddLevel = new CPartnerAddLevel
''oAddLevel.FRectUserid=suserid
oAddLevel.FRectOnlyAdd = "on"

''if (oAddLevel.FRectUserID<>"") then
''	oAddLevel.getUserAddLevelList
''end if

dim i, j, k
%>

<style>
	p {margin:0; padding:0; border:0; font-size:100%;}
	i, em, address {font-style:normal; font-weight:normal;}
 .xls, .down {background-image:url(/images/partner/admin_element.png); background-repeat:no-repeat;}
.btn2 {display:inline-block; font-size:11px !important; letter-spacing:-0.025em; line-height:110%; border-left:1px solid #f0f0f0; border-top:1px solid #f0f0f0; border-right:1px solid #cdcdcd; border-bottom:1px solid #cdcdcd; background-color:#f2f2f2; background-image:-webkit-linear-gradient(#fff, #e1e1e1); background-image:-moz-linear-gradient(#fff, #e1e1e1); background-image:-ms-linear-gradient(#fff, #e1e1e1); background-image:linear-gradient(#fff, #e1e1e1); text-align:center; cursor:pointer;}
.btn2 a {display:block; font-size:11px !important; text-decoration:none !important;}
.btn2 span {display:block;}
.btn2 span em {display:block; padding-top:7px; padding-bottom:4px; text-align:center;}

.fIcon {padding-left:33px;}
.eIcon {padding-right:25px;}

.btn2 .xls {background-position:-125px -135px;}
.btn2 .down {background-position:right -231px;}
.cBk1, .cBk1 a {color:#000 !important;}
	</style>
<!-- 검색 시작 -->
<script language="javascript">

function ViewVacationByID(userid)
{
	window.open("/admin/member/tenbyten/tenbyten_vacation_list.asp?menupos=1178&research=on&page=&part_sn=&deleteyn=N&SearchKey=t.userid&SearchString=" + userid ,"ViewVacationByID","width=1000,height=450,scrollbars=yes")
}


	// 신규 사용자 등록
	function AddItem(){
		var w = window.open("pop_member_reg.asp","popAddIem","width=1400,height=800,scrollbars=yes");
		w.focus();
	}

	// 사용자 수정/삭제
	function ModiItem(empno){
		var w = window.open("pop_member_modify.asp?sEPN="+empno,"ModiItem","width=1400,height=800,scrollbars=yes");
		w.focus();
	}

	// 페이지 이동
	function jsGoPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.target="_self";
		document.frm.action="tenbyten_member_list.asp";
		document.frm.submit();
	}

	//급여정보 등록
	function ViewPay(empno){
		var w = window.open("pop_payform.asp?sEN="+empno,"ModiItem","width=700,height=800,scrollbars=yes");
		w.focus();
	}

	//어드민 권한관리
	function jsMngAuth(empno){
		var w = window.open("popAdminAuth.asp?sEPN="+empno,"popAuth","width=1400,height=768,scrollbars=yes");
		w.focus();
	}

	//추가부서등록
	function jsAddDev(empno){
	 var d = window.open("adddep_reg.asp?sEPN="+empno,"popdep","width=1024,height=600,scrollbars=yes");
		d.focus();
	}

	//검색
	function jsSearch(){
		document.frm.target="_self";
		document.frm.action="tenbyten_member_list.asp";
		document.frm.submit();
		}

	//사원CSV다운
	function jsMemDown(){
		document.frm.target="hidifr";
		document.frm.action="/admin/member/tenbyten/tenbyten_member_list_csv.asp";
		document.frm.submit();
		document.frm.target="";
		document.frm.action="";
	}
	//사원엑셀다운
	function jsMemExcelDown(){
		document.frm.target="hidifr";
		document.frm.action="/admin/member/tenbyten/tenbyten_member_list_excel.asp";
		document.frm.submit();
		document.frm.target="";
		document.frm.action="";
	}

</script>
<iframe id="hidifr" src="" width="0" height="0" frameborder="0"></iframe>
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		부서NEW:
		<%= drawSelectBoxDepartmentALL("department_id", department_id) %>
		<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > 하위 부서직원 제외
	</td>

	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		부서:
		<%=printPartOption("part_sn", part_sn)%>&nbsp;
		직위:
		<%=printPositOptionIN90("posit_sn", posit_sn)%>&nbsp;
		직책:
		<%=printJobOption("job_sn", job_sn)%>&nbsp;
		권한:
		<%=printLevelOption("ilevel_sn", ilevel_sn)%> &nbsp;
		&nbsp;
		개인정보취급권한 :
		<% 'Call DrawSelectBoxCriticInfoUser("criticinfouser", criticinfouser) %>
		<input type="checkbox" name="lv1customerYN" value="Y" <% if lv1customerYN = "Y" then %>checked<% end if %> >LV1(고객정보)
		<input type="checkbox" name="lv2partnerYN" value="Y" <% if lv2partnerYN = "Y" then %>checked<% end if %> >LV2(파트너정보)
		<input type="checkbox" name="lv3InternalYN" value="Y" <% if lv3InternalYN = "Y" then %>checked<% end if %> >LV3(내부정보)
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		재직여부:
		<select name="isUsing" class="select">
			<option value="">전체</option>
			<option value="Y">재직</option>
			<option value="N">퇴사</option>
		</select>
		&nbsp;
		검색:
		<select name="SearchKey" class="select">
			<option value="">::구분::</option>
			<option value="1" >아이디</option>
			<option value="2">사용자명</option>
			<option value="3">사번</option>
		</select>
		<input type="text" class="text" name="SearchString" size="17" value="<%=SearchString%>">
		&nbsp;
		정렬:
		<select name="orderby" class="select">
			<option value="">이름</option>
			<option value="6">입사일(최근)</option>
			<option value="7">입사일(과거순)</option>
			<option value="8">실제입사일(최근)</option>
			<option value="9">실제입사일(과거순)</option>
			<% if C_ADMIN_AUTH or C_PSMngPart then %>
			<option value="2">직급</option>
			<% end if %>
			<option value="3">직책</option>
			<option value="4">내선</option>
			<option value="5">퇴사일</option>
			<option value="1">입사일</option>
		</select>

		&nbsp;
		<input type="checkbox" name="workdaycheck" <% if workdaycheck="Y" then  response.write "checked" %> value="Y">근무일자
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

		&nbsp;
		<script language="javascript">
			document.frm.isUsing.value="<%= isUsing %>";
			document.frm.SearchKey.value="<%= SearchKey %>";
			document.frm.orderby.value="<%= orderby %>";
		</script>
		&nbsp;
		<input type="checkbox" name="nodepartonly" value="Y" <% if (nodepartonly = "Y") then %>checked<% end if %> > 부서NEW 미지정만 (재직인 경우 부서 사용안함 포함)
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="30">
		표시갯수:
		<select class="select" name="pagesize">
			<option value="20" <% if (iPageSize = "20") then %>selected<% end if %> >20 개</option>
			<option value="50" <% if (iPageSize = "50") then %>selected<% end if %> >50 개</option>
			<option value="100" <% if (iPageSize = "100") then %>selected<% end if %> >100 개</option>
			<option value="500" <% if (iPageSize = "500") then %>selected<% end if %> >500 개</option>
		</select>
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="5" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="신규등록" onClick="javascript:AddItem('');">
	</td>
	<td align="right">
		<% '<input type="button" value="CSV다운로드" onclick="jsMemDown();" class="button"> %>
		<input type="button" value="엑셀다운로드" onclick="jsMemExcelDown();" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 상단 띠 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		검색결과 : <b><%=iTotCnt%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=iTotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>직책</td>
	<td width="90">사번</td>
	<td width="60">이름</td>
	<td width="120">영문이름</td>
	<td width="100">아이디</td>

	<td width="90">입사일(정규직)</td>
	<td width="80">실제입사일</td>
	<td width="80">퇴사일</td>

	<td width="80">연차</td>
	<td>부서</td>
	<td>추가관리 부서</td>
	<!--<td>대표매장<br>(관리매장수)</td>-->
	<!--% if C_ADMIN_AUTH or C_PSMngPart then %><td>직급</td--><!--% end if %-->
	<td>직위</td>
	<td width="30">LV1<br>고객<br>정보</td>
	<td width="40">LV2<br>파트너<br>정보</td>
	<td width="30">LV3<br>내부<br>정보</td>
	<!--td>휴가</td-->

	<% if C_ADMIN_AUTH or C_PSMngPart then %>
		<td colspan="2">권한</td>
	<% end if %>

	<td>계약전환여부</td>
</tr>
<% if isArray(arrList) then %>
<% for intLoop=0 to ubound(arrList,2) %>
<tr height=30 align="center" bgcolor="<% if  (arrList(15,intLoop)="Y") then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">

	<td><%=arrList(14,intLoop)%></td>
	<td><a href="javascript:ModiItem('<%=arrList(0,intLoop)%>')"><%=arrList(0,intLoop)%></td>
	<td><%=arrList(1,intLoop)%></td>
	<td><%=arrList(37,intLoop)%></td>
	<td><a href="javascript:ModiItem('<%=arrList(0,intLoop)%>')"><%=arrList(2,intLoop)%></a></td>

	<td><%= Left(arrList(3,intLoop), 10) %></td>
	<td><% if Not IsNull(arrList(24,intLoop)) then %><%= Left(arrList(24,intLoop), 10) %><% end if %></td>
	<td>
	<%
		'실제 퇴사일
		if Not IsNull(arrList(4,intLoop)) then
			if arrList(15,intLoop) <> "N" then
				Response.Write "<font color=""gray"">" & Left(arrList(4,intLoop), 10) & "</font>"
			else
				if (arrList(26,intLoop) = 99) then
					Response.Write "<font color=""red"">" & Left(arrList(4,intLoop), 10) & "</font>"
				else
					Response.Write "<font color=""blue"">" & Left(arrList(4,intLoop), 10) & "</font>"
				end if
			end if
		end if
	%>
	</td>

	<td>
		<%IF Not isNull(arrList(3,intLoop)) and Left(arrList(0,intLoop), 1) = "1" THEN %>
			<% if Not IsNull(arrList(24,intLoop)) then %>
				<% if GetYearDiff(arrList(24,intLoop)) >= 1 then %>
					<%= GetYearDiff(arrList(24,intLoop))  %> 년
				<% end if %>
				<%if GetMonthDiff(arrList(24,intLoop)) > 0 THEN %>
			<%= GetMonthDiff(arrList(24,intLoop)) %> 개월
			<%end if%>
			<% else %><%=arrList(3,intLoop)%>
				<% if GetYearDiff(arrList(3,intLoop)) >= 1 then %>
					<%= GetYearDiff(arrList(3,intLoop))   %> 년
				<% end if %>
				<%if GetMonthDiff(arrList(3,intLoop)) > 0 THEN %>
				<%= GetMonthDiff(arrList(3,intLoop)) %> 개월
				<%end if%>
			<% end if %>
		<%END IF%>
	</td>
	<td align="left">
		<%=arrList(27,intLoop)%>
	</td>
	<!--<td align="left">
		<a href="javascript:shopreg('<%= arrList(0,intLoop) %>');" onfocus="this.blur()">

		<% if arrList(22,intLoop) <> "" then %>
			<%=arrList(21,intLoop)%>/<%=arrList(22,intLoop)%> (<%=arrList(23,intLoop)%>개)
		<% else %>
			<font color="grey">지정없음</font>
		<% end if %>

		</a>
	</td>-->
	<td>
		<% if arrList(23,intLoop) > 0 then %>
		<input type="button" class="button" style ="color:red;" value=" 관리" onClick="jsAddDev('<%= arrList(0,intLoop) %>');">
		<%else%>
		<input type="button" class="button" style ="color:blue;" value="등록" onClick="jsAddDev('<%= arrList(0,intLoop) %>');">
		<%end if%>

	</td>
	<!--% if C_ADMIN_AUTH or C_PSMngPart then %><td nowrap--><!--%=arrList(32,intLoop)%></td--><!--% end if %-->
	<td nowrap><%=arrList(13,intLoop)%></td>
	<!--<td><%'= GetCriticInfoUserLevelName(arrList(30,intLoop)) %></td>-->
	<td><%= arrList(34,intLoop) %></td>
	<td><%= arrList(35,intLoop) %></td>
	<td><%= arrList(36,intLoop) %></td>
	<!--td><input type="button" class="button" value="휴가" onClick="javascript:ViewVacationByID('<%=arrList(2,intLoop)%>');"></td-->

	<% if C_ADMIN_AUTH or C_PSMngPart then %>
		<td align="left">
			<%=arrList(12,intLoop)%>/<%=arrList(28,intLoop)%>
			<%
			if (arrList(2,intLoop) <> "") then
				oAddLevel.FRectUserid = arrList(2,intLoop)
				oAddLevel.getUserAddLevelList
				if (oAddLevel.FResultCount > 0) then
					for i = 0 to oAddLevel.FResultCount - 1
						response.write "<br>" & oAddLevel.FItemList(i).Fpart_name & "/" & oAddLevel.FItemList(i).Flevel_name & vbCrLf
					next
				end if
			end if
			%>
		</td>
		<td>
			<% if isdispmember then %>
				<input type="button" class="button" value="변경" onClick="javascript:jsMngAuth('<%=arrList(0,intLoop)%>');">
			<% end if %>
		</td>
	<%END IF%>

	<td><%if arrList(33,intLoop) >0 then%><font color="red">Y</font><%else%>N<%end if%></td>
</tr>
<% next %>
<% else %>
<tr>
	<td colspan="30" align="center" bgcolor="#FFFFFF">등록(검색)된 사용자가 없습니다.</td>
</tr>
<% end if %>
<!-- 메인 목록 끝 -->

<!-- 페이지 시작 -->
<%
Dim iStartPage,iEndPage,iX,iPerCnt
iPerCnt = 10

iStartPage = (Int((page-1)/iPerCnt)*iPerCnt) + 1

If (page mod iPerCnt) = 0 Then
	iEndPage = page
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30" align="center">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			<tr valign="bottom" height="25">
				<td valign="bottom" align="center">
					<% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
				<% else %>[pre]<% end if %>
				<%
					for ix = iStartPage  to iEndPage
						if (ix > iTotalPage) then Exit for
						if Cint(ix) = Cint(page) then
				%>
					<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
				<%		else %>
					<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
				<%
						end if
					next
				%>
				<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
				<% else %>[next]<% end if %>
				</td>
			</tr>
		</table>
	</td>
</tr>
</table>
<!-- 페이지 끝 -->

<!-- #include virtual="/lib/db/dbclose.asp" -->
