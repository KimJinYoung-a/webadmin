<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  계약직월별급여
' History : 2011.09.07 정윤정 생성
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
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenPayCls.asp" -->
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
    response.write  "권한이 없습니다. - 시스템팀 문의 " ''eastone
    dbget.close() : response.end
end if

'// 2015-06-22, skyer9
''if Not C_ManagerPartTimeMember then
''	response.write  "권한이 없습니다. - 시스템팀 문의 "
''	dbget.close() : response.end
''end if

'if (session("ssAdminPsn") = "10") and (session("ssBctId") <> "bseo") and (session("ssBctId") <> "boyishP") then
'	'// CS팀장님 요청사항, 2015-04-08 2017-10-23 정윤정 삭제
'	response.write  "권한이 없습니다. - 시스템팀 문의 " ''eastone
'	dbget.close() : response.end
'end if

Dim page, SearchKey, SearchString, isUsing, part_sn, research, orderby, selState
Dim job_sn, posit_sn, intY, intM, sYear, sMonth ,shopid
Dim iTotCnt,iPageSize, iTotalPage, totsum,totReCalSum
dim totBasePay, totOverTimePay, totNightTimePay, totHolidayPay, totFoodPay, totPositionPay, totBestPay, totLongWorkPay, totAddPay, totYearPay, totBonusPay, totWorkTime
dim department_id, inc_subdepartment
dim tmpM
	iPageSize	  = request("pagesize")
	if (iPageSize = "") then
		iPageSize = 50
	end if

	page = requestCheckvar(Request("page"),10)
	isUsing = requestCheckvar(Request("isUsing"),1)
	SearchKey = requestCheckvar(Request("SearchKey"),1)
	SearchString = requestCheckvar(Request("SearchString"),32)
	part_sn = requestCheckvar(Request("part_sn"),10)
	job_sn = requestCheckvar(Request("job_sn"),10)
	posit_sn = requestCheckvar(Request("posit_sn"),10)
	sYear = requestCheckvar(Request("sel_DY"),4)
	sMonth = requestCheckvar(Request("sel_DM"),2)
	research = requestCheckvar(Request("research"),2)

	orderby = requestCheckvar(Request("orderby"),1)
	selState = requestCheckvar(Request("selState"),4)

	department_id = requestCheckvar(Request("department_id"),10)
	inc_subdepartment = requestCheckvar(Request("inc_subdepartment"),1)

if isUsing="" and research="" then isUsing="Y"
IF orderby = "" then orderby  = 1
if page="" then page=1
IF sYear = "" and research="" THEN sYear = year(date())
IF sMonth = "" and research="" THEN sMonth = month(date())

'// 로그인정보(등급)에 따라 기본 부서 설정(마스터 이상:2 및 시스템팀:7 제외)
'SCM 메뉴권한 설정에서 제어한다.
if Not (session("ssAdminLsn")<=2  or session("ssAdminPsn")=7  or session("ssAdminPsn")= 8  or session("ssAdminPsn")= 20 )  then
    if (part_sn="") then
	    ''part_sn = session("ssAdminPsn")
	else
	    ''part_sn = checkValidPart(session("ssBctId"),part_sn)   '' if inValid return -999
    end if

	if (department_id = "") then
		department_id = GetUserDepartmentID("",session("ssBctID"))
	end if
end if

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
'-------------------------------------------------------------------------------
dim dSPayDate,dEPayDate,dPreEndDate,sPreYear,sPreMonth,dNextDate,dEndDay
dim searchdate
	searchdate ="N"
IF sYear <> ""   then
	if sMonth <> "" then
		searchdate ="Y"
		dPreEndDate = dateadd("d", -1, dateserial(sYear,sMonth,1)) '이전달 마지막 날짜
		sPreYear = year(dPreEndDate) '이전달
		sPreMonth = month(dPreEndDate) '이전달
		dNextDate = dateadd("m",1, dateserial(sYear,sMonth,1)) '다음달
		dEndDay = day(dateadd("d",-1,dNextDate)) '이번달 마지막날짜

		IF  sYear&"-"&format00(2,sMonth)  = "2014-01" THEN
			dSPayDate = dateserial(sYear,sMonth,1) '급여시작일: 해당월 1일부터
			dEPayDate =dateserial(sYear,sMonth,25)
		ELSEIF sYear&"-"&format00(2,sMonth) > "2014-01" and sYear&"-"&format00(2,sMonth) < "2017-01" THEN
			dSPayDate = dateserial(sPreYear,sPreMonth,26) '급여시작일: 이전월 26일부터
			dEPayDate = dateserial(sYear,sMonth,25) '급여종료일: 해당월 25일까지
		ELSE
			dSPayDate = dateserial(sYear,sMonth,1) '급여시작일: 해당월 1일부터
			dEPayDate = dateserial(sYear,sMonth,dEndDay)  '급여종료일: 해당월 말일까지
		END IF
	 else

	 	 IF  sYear   = "2014" THEN
			dSPayDate = dateserial(sYear,1,1) '급여시작일: 해당월 1일부터
			dEPayDate =dateserial(sYear,12,25)
		ELSEIF sYear > "2014" and sYear < "2017" THEN
			dSPayDate = dateserial(sYear,1,1) '급여시작일: 이전월 26일부터
			dEPayDate = dateserial(sYear,12,25) '급여종료일: 해당월 25일까지
		ELSE
			dSPayDate = dateserial(sYear,1,1) '급여시작일: 해당월 1일부터
			dEPayDate = dateserial(sYear,12,31)  '급여종료일: 해당월 말일까지
		END IF
	end if
ELSE
	dSPayDate = ""
	dEPayDate =""
END IF
'-------------------------------------------------------------------------------
'// 리스트
dim clsPay, arrList,intLoop
Set clsPay = new CPay
	clsPay.FPageSize 	= iPageSize
	clsPay.FCurrPage 	= page
	clsPay.FSYYYYMM		= dSPayDate
	clsPay.FEYYYYMM		= dEPayDate
	clsPay.FSearchType 	= searchKey
	clsPay.FSearchText 	= searchString
	clsPay.Fstatediv 	= isUsing
	clsPay.Fpart_sn 	= part_sn
	clsPay.Fjob_sn 		= job_sn
	clsPay.Fposit_sn 	= posit_sn
	clsPay.Forderby 	= orderby
	clsPay.Fstate 		= selState
	clsPay.FIsMonth		=  1
	clsPay.fshopid		= shopid
	clsPay.FSearchDate  = searchdate
	clsPay.Fdepartment_id 		= department_id
	clsPay.Finc_subdepartment 	= inc_subdepartment

	arrList = clsPay.fnGetMonthlypayList
	iTotCnt = clsPay.FTotCnt
set clsPay = nothing

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

totsum = 0
totBasePay = 0
totOverTimePay = 0
totNightTimePay = 0
totHolidayPay = 0
totFoodPay = 0
totPositionPay = 0
totBestPay = 0
totLongWorkPay = 0
totAddPay = 0
totYearPay = 0
totBonusPay = 0
totWorkTime = 0
totReCalSum = 0
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

	// 페이지 이동
	function jsGoPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.target="_self";
		document.frm.submit();
	}


 	//급여정보 생성
 	function jsSetMonthly(){
 	var wmon = window.open("","popmon","width=700,height=600,scrollbars=yes");
 		document.frm.action = "pop_defaultmonthlypay_list.asp";
 		document.frm.target = "popmon";
 		document.frm.submit();
		wmon.focus();
 	}

 	//급여수정
 	function jsModPay(empno,ino,sYear, sMonth){
 		var wpay	= window.open("tenbyten_pay_reg.asp?menupos=<%=menupos%>&sEN="+empno+"&ino="+ino+"&selY="+sYear+"&selM="+sMonth,"poppay","width=1300,height=600,scrollbars=yes,resizable=yes");
 		wpay.focus();
 	}

 	//프린트
 	function jsPrint(empno,ino){
 	 var winPrint = window.open("print_worktime.asp?sEN="+empno+"&ino="+ino+"&selY=<%=sYear%>&selM=<%=sMonth%>","prtWT","width=1020,height=600,scrollbars=yes,resizable=yes");
 	 winPrint.focus();
 	}

 	// 사용자 수정/삭제
	function jsModMember(empno)
	{
		var w = window.open("pop_member_modify.asp?menupos=<%=menupos%>&sEPN="+empno,"popMem","width=700,height=600,scrollbars=yes,resizeable=yes");
		w.focus();
	}

	//계약정보 등록
	function jsViewPay(empno,ino){
		var wpay = window.open("pop_payform.asp?sEN="+empno+"&ino="+ino,"popPay","width=700,height=600,scrollbars=yes,resizeable=yes");
		wpay.focus();
	}

	//검색
	function jsSearch(){
		document.frm.target="_self";
		document.frm.action="tenbyten_pay_list.asp";
		document.frm.submit();
		}
//엑셀다운
	function jsMemDown(){
		document.frm.target="hidifr";
		document.frm.action="tenbyten_pay_list_csv.asp";
		document.frm.submit();
		}
</script>
<iframe id="hidifr" src="" width="0" height="0" frameborder="0"></iframe>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<form name="frm" method="get" action="tenbyten_pay_list.asp">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="research" value="on">
		<input type="hidden" name="page" value="">
		<tr align="center" bgcolor="#FFFFFF" >
			<td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
			<td align="left">
				부서NEW:
				<% if session("ssAdminLsn")<=2 or session("ssAdminPsn")=7 or  session("ssAdminPsn")=8  or session("ssAdminPsn")= 20 then %>

					<%= drawSelectBoxDepartment("department_id", department_id) %>
				<% else %>

					<%= drawSelectBoxMyDepartment(session("ssBctId"), "department_id", department_id) %>
				<% end if %>
				<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > 하위 부서직원 제외
				&nbsp;&nbsp;

				근무년월:
				<select name="sel_DY">
				<option value="">-선택-</option>
				<%For intY = Year(date()) to 2010 step -1 %>
				<option value="<%=intY%>" <%IF Cstr(sYear) = Cstr(intY) THEN%>selected<%END IF%>><%=intY%></option>
				<%Next%>
				</select>
				<select name="sel_DM">
				<option value="">-선택-</option>
				<%For intM = 1 to 12 %>
				<option value="<%=intM%>" <%IF Cstr(sMonth) = Cstr(intM) THEN%>selected<%END IF%>><%=format00(2,intM)%></option>
				<%Next%>
				</select>&nbsp;&nbsp;&nbsp;
				계약구분:
				<%=printPositOptionPartTime("posit_sn", posit_sn)%>&nbsp;
			</td>
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
				<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
			</td>
		</tr>
		<tr  bgcolor="#FFFFFF" >
			<td>
				검색:
				<select name="SearchKey" class="select">
					<option value="">::구분::</option>
					<option value="1" >사번</option>
					<option value="2">이름</option>
				</select>
				<input type="text" class="text" name="SearchString" size="16" value="<%=SearchString%>">
				&nbsp;&nbsp;&nbsp;
				<!--계약상태:
				<select name="selInDate">
				<option value="1">진행중</option>
				<option value="2">종료예정</option>
				<option value="3">종료</option>
				</select>
				&nbsp;&nbsp;&nbsp;-->
				상태:
				<select name="selState">
				<option value="">::전체::</option>
				<option value="-1">입력대기</option>
				<option value="0">작성중</option>
				<option value="1">작성완료</option>
				<option value="5">확인완료</option>
				<option value="7">입금완료</option>
				</select>&nbsp;&nbsp;&nbsp;
				정렬:
				<select name="orderby" class="select">
					<option value="1">사번</option>
					<option value="2">이름</option>
				</select>

				표시갯수:
				<select class="select" name="pagesize">
					<option value="20" <% if (iPageSize = "20") then %>selected<% end if %> >20 개</option>
					<option value="50" <% if (iPageSize = "50") then %>selected<% end if %> >50 개</option>
					<option value="100" <% if (iPageSize = "100") then %>selected<% end if %> >100 개</option>
					<option value="500" <% if (iPageSize = "500") then %>selected<% end if %> >500 개</option>
				</select>

				&nbsp;
				<script language="javascript">
					//document.frm.isUsing.value="<%= isUsing %>";
					document.frm.SearchKey.value="<%= SearchKey %>";
					document.frm.orderby.value="<%= orderby %>";
					document.frm.selState.value="<%= selState %>";
				</script>

			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<!-- 검색 끝 -->
 <tr>
	<td align="right"><!--<input type="button" class="button" value="월별급여 생성" onClick="jsSetMonthly();">-->
	<span class="btn2 cBk1" style="vertical-align:top;"><a href="javascript:jsMemDown();"><span class="eIcon down"><em class="fIcon xls">계약직월별급여</em></span></a></span>
	</td>
</tr>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="25">
				검색결과 : <b><%=iTotCnt%></b>
				&nbsp;
				페이지 : <b><%= page %> / <%=iTotalPage%></b>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>날짜</td>
			<td>사번</td>
			<td>이름</td>
			<td>부서</td>
			<td>대표매장(관리매장수)</td>
			<td>계약구분</td>
			<td>시간급</td>
			<td>기본급</td>
			<td>시간외수당</td>
			<td>야간근무수당</td>
			<td>휴일근무수당</td>
			<td>식대지원</td>
			<td>직책수당</td>
			<td>우수사원수당</td>
			<td>장기근속수당</td>
			<td>추가수당</td>
			<td>연차수당</td>
			<td>상여금</td>
			<td>총근무시간</td>
			<td>합계</td>
			<td>상태</td>
			<%if C_MngPart or C_ADMIN_AUTH or C_PSMngPart then%>
			<td>전월차액금</td>
			<%end if%>
			<td>프린트</td>
	    </tr>

		<% if isArray(arrList) then %>
		<%
		for intLoop=0 to ubound(arrList,2)

			totsum = totsum + arrList(16,intLoop)

	if arrList(0,intLoop)= "90201501120013" or arrList(0,intLoop)="90201610010124" or arrList(0,intLoop)="90201611130141" or arrList(0,intLoop)="90201611140136" or arrList(0,intLoop)="90201611200158" or arrList(0,intLoop)="90201611260172" or arrList(0,intLoop)="90201612060169" or arrList(0,intLoop)="90201612100180" or arrList(0,intLoop)="90201612120174" or arrList(0,intLoop)="90201612210190" then
			totBasePay 			= totBasePay + arrList(7,intLoop)
			totOverTimePay 		= totOverTimePay + arrList(8,intLoop)
			totNightTimePay 	= totNightTimePay + arrList(9,intLoop)
			totHolidayPay 		= totHolidayPay + arrList(10,intLoop)
			totFoodPay 			= totFoodPay + arrList(11,intLoop)
	else
			totBasePay 			= totBasePay + arrList(7,intLoop)+arrList(39,intLoop)
			totOverTimePay 		= totOverTimePay + arrList(8,intLoop)+arrList(40,intLoop)
			totNightTimePay 	= totNightTimePay + arrList(9,intLoop)+arrList(41,intLoop)
			totHolidayPay 		= totHolidayPay + arrList(10,intLoop)+arrList(42,intLoop)
			totFoodPay 			= totFoodPay + arrList(11,intLoop)+arrList(43,intLoop)
	end if
			totPositionPay 		= totPositionPay + arrList(12,intLoop)
			totBestPay 			= totBestPay + arrList(13,intLoop)
			totLongWorkPay 		= totLongWorkPay + arrList(14,intLoop)
			totAddPay 			= totAddPay + arrList(31,intLoop)
			totYearPay 			= totYearPay + arrList(36,intLoop)
			totBonusPay 		= totBonusPay + arrList(37,intLoop)

			if arrList(26,intLoop)=13 then
				totWorkTime 		= totWorkTime + arrList(35,intLoop)
			end if

			totReCalSum = totReCalSum + arrList(44,intLoop)

		%>
		<tr height=30 align="center" bgcolor="#FFFFFF">
			<td nowrap><%=arrList(45,intLoop)%></TD>
			<td><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><%=arrList(0,intLoop)%></a></td>
			<td><a href="javascript:jsModMember('<%=arrList(0,intLoop)%>')"><%=arrList(22,intLoop)%></a></td>
			<td>
				<%=arrList(38,intLoop)%>
			</td>
			<td align="left">
				<a href="javascript:shopreg('<%= arrList(0,intLoop) %>');" onfocus="this.blur()">

				<% if arrList(33,intLoop) <> "" then %>
					<%=arrList(32,intLoop)%>/<%=arrList(33,intLoop)%> (<%=arrList(34,intLoop)%>개)
				<% else %>
					<font color="grey">지정없음</font>
				<% end if %>

				</a>
			</td>
			<td><a href="javascript:jsViewPay('<%=arrList(0,intLoop)%>','<%=arrList(30,intLoop)%>')"><%IF arrList(26,intLoop)  =   13 THEN%><font color="#D2691E"><%END IF%><%=arrList(28,intLoop)%></a></td>
			<td><%=formatnumber(arrList(20,intLoop),0)%></td>
			<%if arrList(0,intLoop)= "90201501120013" or arrList(0,intLoop)="90201610010124" or arrList(0,intLoop)="90201611130141" or arrList(0,intLoop)="90201611140136" or arrList(0,intLoop)="90201611200158" or arrList(0,intLoop)="90201611260172" or arrList(0,intLoop)="90201612060169" or arrList(0,intLoop)="90201612100180" or arrList(0,intLoop)="90201612120174" or arrList(0,intLoop)="90201612210190" then%>
			<td><%=formatnumber(arrList(7,intLoop),0)%></td>
			<td><%=formatnumber(arrList(8,intLoop),0)%></td>
			<td><%=formatnumber(arrList(9,intLoop),0)%></td>
			<td><%=formatnumber(arrList(10,intLoop),0)%></td>
			<td><%=formatnumber(arrList(11,intLoop),0)%></td>
			<%else%>
			<td><%=formatnumber(arrList(7,intLoop)+arrList(39,intLoop),0)%></td>
			<td><%=formatnumber(arrList(8,intLoop)+arrList(40,intLoop),0)%></td>
			<td><%=formatnumber(arrList(9,intLoop)+arrList(41,intLoop),0)%></td>
			<td><%=formatnumber(arrList(10,intLoop)+arrList(42,intLoop),0)%></td>
			<td><%=formatnumber(arrList(11,intLoop)+arrList(43,intLoop),0)%></td>
			<%end if%>
			<td><%=formatnumber(arrList(12,intLoop),0)%></td>
			<td><%=formatnumber(arrList(13,intLoop),0)%></td>
			<td><%=formatnumber(arrList(14,intLoop),0)%></td>
			<td><%=formatnumber(arrList(31,intLoop),0)%></td>
			<td><%=formatnumber(arrList(36,intLoop),0)%></td>
			<td><%=formatnumber(arrList(37,intLoop),0)%></td>
			<td>
			    <% IF arrList(26,intLoop)=13 THEN%>
					<%=fnSetTimeFormat(arrList(35,intLoop))%>
			    <% end if %>
			</td>
			<td><%=formatnumber(arrList(16,intLoop),0)%></td>
			<td>
			<% tmpM = right(arrList(45,intLoop),2)
			if left(tmpM,1) = 0 then tmpM = right(tmpM,1)
			%>
			<a href="javascript:jsModPay('<%=arrList(0,intLoop)%>','<%=arrList(30,intLoop)%>','<%=left(arrList(45,intLoop),4)%>','<%=tmpM%>');"><%=fnGetStateDesc(arrList(17,intLoop))%></a></td>
			<%if C_MngPart or C_ADMIN_AUTH or C_PSMngPart then%>
			<td><%=formatnumber(arrList(44,intLoop),0)%></td>
			<%end if%>
			<td><a href="javascript:jsPrint('<%=arrList(0,intLoop)%>','<%=arrList(30,intLoop)%>')" onFocus="this.blur()"><img src="/images/icon_print02.gif" border="0"></a></td>
		</tr>
		<% next %>
		<tr align="center" bgcolor="#FFFFFF">
			<td colspan=7>합계</td>
			<td><%=formatnumber(totBasePay,0)%></td>
			<td><%=formatnumber(totOverTimePay,0)%></td>
			<td><%=formatnumber(totNightTimePay,0)%></td>
			<td><%=formatnumber(totHolidayPay,0)%></td>
			<td><%=formatnumber(totFoodPay,0)%></td>
			<td><%=formatnumber(totPositionPay,0)%></td>
			<td><%=formatnumber(totBestPay,0)%></td>
			<td><%=formatnumber(totLongWorkPay,0)%></td>
			<td><%=formatnumber(totAddPay,0)%></td>
			<td><%=formatnumber(totYearPay,0)%></td>
			<td><%=formatnumber(totBonusPay,0)%></td>
			<td><%=fnSetTimeFormat(totWorkTime)%></td>
			<td><%=formatnumber(totsum,0)%></td>
			<td></td>
			<td><%=formatnumber(totReCalSum,0)%></td>
			<td></td>
		</tr>
		<% else %>
		<tr>
			<td colspan="25" align="center" bgcolor="#FFFFFF">등록(검색)된 사용자가 없습니다.</td>
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
			<td colspan="25" align="center">
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
	</td>
</tr>
</table>
<!-- 페이지 끝 -->
</body>
</html>
 <!-- #include virtual="/lib/db/dbclose.asp" -->
