<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  계약직관리
' History : 2011.09.21 정윤정 생성
'			2011.12.16 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	if (isNull(session("ssAdminPOSITsn")) or session("ssAdminPOSITsn")="" or isNull(session("ssAdminPOsn")) or session("ssAdminPOsn")="" or isNull(session("ssAdminLsn")) or session("ssAdminLsn")="" or isNull(session("ssAdminPsn")) or session("ssAdminPsn")="") then
        response.write  "권한이 없습니다. - 로그아웃 후 다시 로그인해주세요. "
        dbget.close() : response.end
    end if

	if (session("ssAdminLsn")=0) then
		dim bufbuf : bufbuf = 1/0  ''raize Error
		dbget.close() : response.end
	end if

	if (Not (C_ADMIN_AUTH or C_MngPart or C_ManagerUpJob or C_PSMngPart)) then
        response.write  "권한이 없습니다. - 시스템팀 문의 "
        dbget.close() : response.end
    end if

	'// 2015-06-22, skyer9
	''if Not C_ManagerPartTimeMember then
    ''    response.write  "권한이 없습니다. - 시스템팀 문의 "
    ''    dbget.close() : response.end
	''end if

	'// CS팀장님 요청사항, 2015-04-08
	if (session("ssAdminPsn") = "10") and (session("ssBctId") <> "boyishP") and (session("ssBctId") <> "oesesang52") and (session("ssBctId") <> "rabbit1693") then
		response.write  "권한이 없습니다. - 시스템팀 문의 " ''eastone
		dbget.close() : response.end
	end if

	Dim page, SearchKey, SearchString, isUsing, part_sn, research, orderby
	Dim job_sn, posit_sn, chkdate ,shopid
	Dim iTotCnt,iPageSize, iTotalPage

	dim workdaycheck, yyyy1, yyyy2, mm1, mm2, dd1, dd2
	dim fromDate, toDate
	dim maxinoonly
	dim department_id, inc_subdepartment

	maxinoonly = request("maxinoonly")

	workdaycheck = request("workdaycheck")

	yyyy1 = request("yyyy1")
	yyyy2 = request("yyyy2")
	mm1	  = request("mm1")
	mm2	  = request("mm2")
	dd1	  = request("dd1")
	dd2	  = request("dd2")

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

	iPageSize	  = request("pagesize")
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
	chkdate = requestCheckvar(Request("chkdate"),1)
	research = requestCheckvar(Request("research"),2)

	orderby = requestCheckvar(Request("orderby"),1)
	if orderby = "" then orderby = 1

	department_id = requestCheckvar(Request("department_id"),10)
	inc_subdepartment = requestCheckvar(Request("inc_subdepartment"),1)

	if Not(session("ssAdminLsn")<=2 or session("ssAdminPsn")=7 or session("ssAdminPsn")=8  or session("ssAdminPsn")= 20 ) then	'마스터권한 이상이거나 경영관리팀, 시스템팀일경우만 부서리스트 보여주고 그 외에는 해당부서만 보여준다
		if (part_sn="") then
		    part_sn = session("ssAdminPsn")
		else
		    part_sn = checkValidPart(session("ssBctId"),part_sn)   '' if inValid return -999
	    end if

		if (department_id = "") then
			department_id = GetUserDepartmentID("",session("ssBctID"))
		end if
	end if

	if isUsing="" and research="" then isUsing="Y"
	if page="" then page=1

	'직영/가맹점
	if (C_IS_SHOP) then

		'/오프라인 관리자가 아닌
		if not(C_OFF_AUTH) then

			'/어드민권한 점장 미만
			'if getlevel_sn("",session("ssBctId")) > 6 then
				shopid = C_STREETSHOPID
			'end if
		end if
	end if
'response.write C_IS_SHOP
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
	oMember.Fposit_sn 	= posit_sn
	oMember.Forderby 		= orderby
	oMember.FchkDate		= chkdate
	oMember.fshopid		= shopid

	oMember.Fdepartment_id 		= department_id
	oMember.Finc_subdepartment 	= inc_subdepartment

	if (workdaycheck = "Y") then
		oMember.FStartDate		= fromDate
		oMember.FEndDate		= toDate
	end if

	oMember.FMaxInoOnly		= maxinoonly

	arrList = oMember.fnGetContractMemberList
	iTotCnt = oMember.FTotCnt
	set oMember = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<!-- 검색 시작 -->
<script language="javascript">
<!--
	// 신규 사용자 등록
	function jsAddMember()
	{
		var w = window.open("pop_member_reg.asp?menupos=<%=menupos%>","popMem","width=1400,height=800,scrollbars=yes,resizeable=yes");
		w.focus();
	}

	// 사용자 수정/삭제
	function jsModMember(empno)
	{
		var w = window.open("pop_member_modify.asp?menupos=<%=menupos%>&sEPN="+empno,"popMem","width=1400,height=800,scrollbars=yes,resizeable=yes");
		w.focus();
	}

	// 페이지 이동
	function jsGoPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}

	//계약정보 등록
	function jsViewPay(empno,ino){
		var wpay = window.open("pop_payform.asp?sEN="+empno+"&ino="+ino,"popPay","width=800,height=800,scrollbars=yes,resizeable=yes");
		wpay.focus();
	}

 	//근무시간 등록
 	function jsWorkTime(empno,ino){
 		var wwt =window.open("pop_worktime.asp?sEN="+empno+"&ino="+ino,"popWT","width=1020,height=600,scrollbars=yes,resizeable=yes");
		wwt.focus();
	}

	function jsCodeManage(){
		var winCode;
		winCode = window.open('/admin/member/tenbyten/popManageCode.asp','popCode','width=450,height=600,scrollbars=yes,resizable=yes');
		winCode.focus();
	}
 //-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			부서NEW:
			<% IF session("ssAdminLsn")<=2 or session("ssAdminPsn")=7 or session("ssAdminPsn")=8  or session("ssAdminPsn")= 20 THEN %>
			<%= drawSelectBoxDepartment("department_id", department_id) %>
			<% else %>
			<%= drawSelectBoxMyDepartment(session("ssBctId"), "department_id", department_id) %>
			<% end if %>
			<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > 하위 부서직원 제외
			&nbsp;&nbsp;
			계약구분:
			<%=printPositOptionPartTime("posit_sn",posit_sn)%>
			<!--직책:
			<%=printJobOption("job_sn", job_sn)%>&nbsp;-->&nbsp;
			계약상태:
			<select name="chkDate">
			<option value="">::전체::</option>
			<option value="1">진행중</option>
			<option value="2">종료예정</option>
			<option value="3">종료</option>
			</select>
		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
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
				<option value="2">이름</option>
				<option value="3">사번</option>
			</select>
			<input type="text" class="text" name="SearchString" size="17" value="<%=SearchString%>">
			&nbsp;
			정렬:
			<select name="orderby" class="select">
				<option value="1">사번</option>
				<option value="2">이름</option>
				<option value="3">직급</option>
				<!--<option value="4">직책</option>-->
				<option value="5">내선</option>
				<option value="6">퇴사일</option>
				<option value="7">입사일(최근순)</option>
			</select>
			&nbsp;
    		<input type="checkbox" name="workdaycheck" <% if workdaycheck="Y" then  response.write "checked" %> value="Y">근무일자
    		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

			&nbsp;
			<script language="javascript">
				document.frm.isUsing.value="<%= isUsing %>";
				document.frm.SearchKey.value="<%= SearchKey %>";
				document.frm.chkDate.value="<%= chkDate %>";
				document.frm.orderby.value="<%= orderby %>";
			</script>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="checkbox" name="maxinoonly" <% if maxinoonly="Y" then  response.write "checked" %> value="Y">마지막 회차만
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
			<input type="button" class="button" value="신규등록" onClick="javascript:jsAddMember();">
		</td>
		<td align="right">
			<input type="button" class="button" value="코드관리" onClick="jsCodeManage();">
			<!--<img src="/images/icon_excel.gif" border="0">-->
		</td>
	</tr>
</table>
<!-- 액션 끝 -->
<p>

<!-- 상단 띠 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			검색결과 : <b><%=iTotCnt%></b>
			&nbsp;
			페이지 : <b><%= page %> / <%=iTotalPage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="70">직책</td>
		<td width="100">사번</td>
		<td>이름</td>
		<td width="70">아이디</td>
		<td width="220">부서</td>
		<td>대표매장(관리매장수)</td>
		<td width="70">계약구분</td>
		<td width="70">입사일</td>
		<td width="70">퇴사일</td>
		<td>계약회차</td>
		<td width="70">계약시작일</td>
		<td width="70">계약종료일</td>
		<td width="70">시급</td>
		<td width="100">총급여</td>
		<td>계약설정</td>
		<td>근무시간</td>
    </tr>
	<% if isArray(arrList) then %>
	<% for intLoop=0 to ubound(arrList,2) %>
	<tr height=30 align="center" bgcolor="<% if  (arrList(15,intLoop)="Y") then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">

		<td><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><%=arrList(14,intLoop)%></a></td>
		<td><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><%=arrList(0,intLoop)%></a></td>
		<td><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><%=arrList(1,intLoop)%></a></td>
		<td><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><%=arrList(2,intLoop)%></a></td>
		<td>
			<a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><%=arrList(28,intLoop)%></a>
		</td>
		<td align="left">

			<a href="javascript:shopreg('<%= arrList(0,intLoop) %>');" onfocus="this.blur()">

			<% if arrList(24,intLoop) <> "" then %>
				<%=arrList(23,intLoop)%>/<%=arrList(24,intLoop)%> (<%=arrList(25,intLoop)%>개)
			<% else %>
				<font color="grey">지정없음</font>
			<% end if %>

			</a>
		</td>
		<td><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><%IF arrList(6,intLoop)  =   13 THEN%><font color="#D2691E"><%END IF%><%=arrList(13,intLoop)%></a></td>
		<td><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><%= Left(arrList(3,intLoop), 10) %></a></td>
		<td><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')">
			<%IF not isNull(arrList(4,intLoop)) and arrList(15,intLoop) ="N" THEN %>
				<% if (arrList(27,intLoop) = 99) then %><font color="red"><% end if %>
				<%= Left(arrList(4,intLoop), 10) %>
			<%END IF%>
		</a>

	</td>
		<td><%=arrList(22,intLoop)%></td>
		<td><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><font color="blue"><%=arrList(17,intLoop)%></font></a></td>
		<td><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><font color="blue"><%=arrList(18,intLoop)%></font></a></td>
		<td align="right"><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><%=formatnumber(arrList(19,intLoop),0)%></a></td>
		<td align="right"><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><%=formatnumber(arrList(20,intLoop),0)%></a></td>
		<td> <a href="javascript:jsViewPay('<%=arrList(0,intLoop)%>','<%=arrList(22,intLoop)%>');"><%IF not isNull(arrList(16,intLoop))  THEN%><font color="blue">[수정]</font><%ELSE%><font color="red">[설정]</font><%END IF%></a> </td>
		<td>
			<%'IF arrList(6,intLoop)  =   13 THEN		' 시급직만 시간 표시 %>
				<a href="javascript:jsWorkTime('<%=arrList(0,intLoop)%>','<%=arrList(22,intLoop)%>')"><font color="#D2691E">
					<% if (arrList(26,intLoop) <> 0) then %>
						<%= Fix((arrList(26,intLoop)/ 60)) %>:<%= Format00(2, (arrList(26,intLoop) mod 60)) %>
					<% else %>
						[입력]
					<% end if %>
				</a>
			<%'END IF%>
		</td>
	</tr>
	<% next %>
	<% else %>
	<tr>
		<td colspan="20" align="center" bgcolor="#FFFFFF">등록(검색)된 사용자가 없습니다.</td>
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
		<td colspan="20" align="center">
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
