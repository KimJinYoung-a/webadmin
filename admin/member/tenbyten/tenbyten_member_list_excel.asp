<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 사원관리 엑셀다운로드
' History : 2011.1.19 정윤정 생성
'           2018.03.30 허진원 - 직급 선택 표시
'			2023.08.23 한용민 수정(csv다운로드->엑셀다운로드 로 새로 만듬)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
Dim page, SearchKey, SearchString, isUsing, part_sn, research, orderby, ilevel_sn, iTotCnt,iPageSize, iTotalPage
Dim job_sn, posit_sn, continuous_service_year, employeeonly, nodepartonly, criticinfouser, workdaycheck
dim fromDate, toDate, department_id, inc_subdepartment, rank_sn, lv1customerYN, lv2partnerYN, lv3InternalYN
dim  yyyy1, yyyy2, mm1, mm2, dd1, dd2, oMember, arrList,i, j, k,p, sTitle, bufStr, menupos
	workdaycheck = requestcheckvar(request("workdaycheck"),1)
	lv1customerYN 	= requestCheckvar(request("lv1customerYN"),1)
	lv2partnerYN 	= requestCheckvar(request("lv2partnerYN"),1)
	lv3InternalYN 	= requestCheckvar(request("lv3InternalYN"),1)
	yyyy1 = requestcheckvar(getNumeric(request("yyyy1")),4)
	yyyy2 = requestcheckvar(getNumeric(request("yyyy2")),4)
	mm1	  = requestcheckvar(getNumeric(request("mm1")),2)
	mm2	  = requestcheckvar(getNumeric(request("mm2")),2)
	dd1	  = requestcheckvar(getNumeric(request("dd1")),2)
	dd2	  = requestcheckvar(getNumeric(request("dd2")),2)
	iPageSize	  = requestcheckvar(getNumeric(request("pagesize")),10)
	page = requestCheckvar(getNumeric(Request("page")),10)
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
    menupos = requestCheckvar(getNumeric(Request("menupos")),10)

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

if isUsing="" and research="" then isUsing="Y"
if page="" then page=1

Set oMember = new CTenByTenMember
	oMember.FPagesize 	= 10000
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
	arrList = oMember.fnGetMemberList_csv
	iTotCnt = oMember.FTotalCount
set oMember = nothing

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENEmployeeManagementList" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
downFilemenupos=menupos
downPersonalInformation_rowcnt=ubound(arrLIst,2)+1
%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<html>
<head>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="17">
		검색결과 : <b><%= ubound(arrLIst,2)+1 %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>직책</td>
    <td>사번</td>
    <td>이름</td>
    <td>영문이름</td>
    <td>아이디</td>
    <td>입사일(정규직)</td>
    <td>실제입사일</td>
    <td>퇴사일</td>
    <td>연차</td>
    <td>부서</td>
    <td>직위</td>
    <td>LV1(고객정보)</td>
    <td>LV2(파트너정보)</td>
    <td>LV3(내부정보)</td>
    <td>계약전환여부</td>
    <td>GSSHOP아이디</td>
    <td>핸드폰번호</td>
</tr>
<% if isarray(arrLIst) then %>
<% for i=0 to ubound(arrLIst,2) %>
<tr bgcolor="#FFFFFF" align="center">
    <td><%= arrLIst(14,i) %></td>
    <td class="txt"><%= arrLIst(0,i) %></td>
    <td><%= arrLIst(1,i) %></td>
    <td><%= arrLIst(44,i) %></td>
    <td class="txt"><%= arrLIst(2,i) %></td>
    <td><%= Left(arrList(3,i), 10) %></td>
    <td>
        <% if Not IsNull(arrList(24,i)) then %>
            <%= Left(arrList(24,i), 10) %>
        <% end if %>
    </td>
    <td>
        <% if Not IsNull(arrList(4,i)) then %>
            <% if arrList(15,i) <> "N" then %>
                <%= Left(arrList(4,i), 10) %>
            <% else %>
                <% if (arrList(26,i) = 99) then %>
                    <%= Left(arrList(4,i), 10) %>
                <% else %>
                    <%= Left(arrList(4,i), 10) %>
                <% end if %>
            <% end if %>
        <% end if %>
	</td>
    <td>
        <% IF Not isNull(arrList(3,i)) and Left(arrList(0,i), 1) = "1" THEN %>
            <% if Not IsNull(arrList(24,i)) then %>
                <% if GetYearDiff(arrList(24,i)) >= 1 then %>
                    <%= GetYearDiff(arrList(24,i))  &"년" %>
                <% end if %>
                <% if GetMonthDiff(arrList(24,i)) > 0 THEN %>
                    <%= GetMonthDiff(arrList(24,i)) &"개월" %>
                <% end if %>
            <% else %>
                <%= arrList(3,i) %>
                <% if GetYearDiff(arrList(3,i)) >= 1 then %>
                    <%= GetYearDiff(arrList(3,i))&"년" %>
                <% end if %>
                <% if GetMonthDiff(arrList(3,i)) > 0 THEN %>
                    <%= GetMonthDiff(arrList(3,i))&"개월" %>
                <% end if %>
            <% end if %>
        <% end if %>
	</td>
    <td><%= arrLIst(27,i) %></td>
    <td><%= arrLIst(13,i) %></td>
    <td><%= arrLIst(41,i) %></td>
    <td><%= arrLIst(42,i) %></td>
    <td><%= arrLIst(43,i) %></td>
    <td>
		<% if arrList(33,i) >0 then %>
			Y
		<% else %>
			N
		<% end if %>
	</td>
    <td><%= db2html(arrList(40,i)) %></td>
    <td><%= arrLIst(17,i) %></td>
</tr>
<%
if i mod 500 = 0 then
    Response.Flush		' 버퍼리플래쉬
end if

next
%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="17" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
