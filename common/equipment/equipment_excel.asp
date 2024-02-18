<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 장비자산관리
' History : 2008년 06월 27일 한용민 생성
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/common/equipment/equipment_cls.asp"-->
<%
dim page, equip_gubun, part_sn, idx, using_userid, using_username, usingIp , equip_code ,equip_name ,manufacture_company, manufacture_sn
dim ipgocheck, yyyy1, yyyy2, mm1, mm2, dd1, dd2, fromDate, toDate, useip ,property_gubun ,research, BIZSECTION_CD, BIZSECTION_NM
dim totalcurrsum,totaljasan, Alltotaljasan, part_code ,state , parameter, only1000, i, accountassetcode, sorttype
dim onlyusing, accountGubun, department_id, outcheck, yyyy3, yyyy4, mm3, mm4, dd3, dd4, fromDate2, toDate2, paymentrequestidx, buyCompanyName
	accountassetcode = requestcheckvar(request("accountassetcode"),32)
	paymentrequestidx = requestcheckvar(request("paymentrequestidx"),10)
	page = requestcheckvar(request("page"),10)
	equip_gubun = requestcheckvar(Request("equip_gubun"),2)
	part_sn = requestcheckvar(Request("part_sn"),10)
	using_userid = requestcheckvar(Request("using_userid"),32)
	using_username = requestcheckvar(Request("using_username"),32)
	equip_code = requestcheckvar(request("equip_code"),20)
	ipgocheck = requestcheckvar(request("ipgocheck"),2)
	yyyy1 = requestcheckvar(request("yyyy1"),4)
	yyyy2 = requestcheckvar(request("yyyy2"),4)
	mm1	  = requestcheckvar(request("mm1"),2)
	mm2	  = requestcheckvar(request("mm2"),2)
	dd1	  = requestcheckvar(request("dd1"),2)
	dd2	  = requestcheckvar(request("dd2"),2)
	part_code = requestcheckvar(Request("part_code"),10)
	equip_name = requestcheckvar(Request("equip_name"),64)
	manufacture_company = requestcheckvar(Request("manufacture_company"),64)
	buyCompanyName = requestcheckvar(Request("buyCompanyName"),64)
	manufacture_sn = requestcheckvar(Request("manufacture_sn"),64)
	property_gubun = requestcheckvar(Request("property_gubun"),10)
	state = requestcheckvar(Request("state"),10)
	research = requestcheckvar(Request("research"),2)
	onlyusing = requestcheckvar(Request("onlyusing"),2)
	accountGubun = requestcheckvar(Request("accountGubun"),5)
	department_id = requestcheckvar(Request("department_id"),5)
	BIZSECTION_CD = requestcheckvar(Request("BIZSECTION_CD"),15)
	BIZSECTION_NM = requestcheckvar(Request("BIZSECTION_NM"),55)
	only1000 = requestcheckvar(Request("only1000"),55)
	outcheck = requestcheckvar(request("outcheck"),2)
	yyyy3 = requestcheckvar(request("yyyy3"),4)
	yyyy4 = requestcheckvar(request("yyyy4"),4)
	mm3	  = requestcheckvar(request("mm3"),2)
	mm4	  = requestcheckvar(request("mm4"),2)
	dd3	  = requestcheckvar(request("dd3"),2)
	dd4	  = requestcheckvar(request("dd4"),2)
	sorttype = requestcheckvar(request("sorttype"),1)

if sorttype="" then sorttype="1"
if page="" then page=1
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if (yyyy3="") then yyyy3 = Cstr(Year(now()))
if (mm3="") then mm3 = Cstr(Month(now()))
if (dd3="") then dd3 = Cstr(day(now()))
if (yyyy4="") then yyyy4 = Cstr(Year(now()))
if (mm4="") then mm4 = Cstr(Month(now()))
if (dd4="") then dd4 = Cstr(day(now()))

fromDate2 = CStr(DateSerial(yyyy3, mm3, dd3))
toDate2 = CStr(DateSerial(yyyy4, mm4, dd4+1))

if (research = "") then
	onlyusing = "Y"
end if

'view목록과 sync
'if research <> "on" and property_gubun = "" then
'	'//시스템팀 일경우 자산구분 초기값 시스템자산
'	if getpart_sn("",session("ssBctId")) = "7" then
'		property_gubun = "10"
'		
'	'//오프라인 본사,매장 일경우 자산구분 초기값 오프라인자산
'	elseif getpart_sn("",session("ssBctId")) = "13" or getpart_sn("",session("ssBctId")) = "18" then
'		property_gubun = "11"
'		
'	else
'		property_gubun = "10"
'	end if
'end if
'if property_gubun = "" then property_gubun = "10"

dim oequip
set oequip = new CEquipment
	oequip.FPageSize = 50
	oequip.FCurrPage = page
	oequip.FRectequip_gubun = equip_gubun
	oequip.FRectpart_sn = part_sn
	oequip.FRectusing_userid = using_userid
	oequip.FRectusing_username = using_username
	oequip.Frectequip_code = equip_code
	oequip.frectequip_name = equip_name
	oequip.frectmanufacture_company = manufacture_company
	oequip.fRectBuyCompanyName = buyCompanyName
	oequip.frectmanufacture_sn = manufacture_sn
	oequip.frectproperty_gubun = property_gubun
	oequip.frectstate = state
	oequip.FRectIsusing = onlyusing
	oequip.FRectAccountGubun = accountGubun
	oequip.FRectDepartmentID = department_id
	oequip.FRectBIZSECTION_CD = BIZSECTION_CD
	oequip.FRectOnly1000 = only1000
	oequip.frectaccountassetcode = accountassetcode
	oequip.frectpaymentrequestidx = paymentrequestidx
	oequip.frectsorttype = sorttype

	if ipgocheck = "on" then
		oequip.frectbuy_startdate = fromDate
		oequip.frectbuy_enddate = toDate
	end if

	if outcheck = "on" then
		oequip.frectout_startdate = fromDate2
		oequip.frectout_enddate = toDate2
	end if

	oequip.getEquipmentList

totalcurrsum = 0
totaljasan	 = 0
Alltotaljasan = 0
%>
<!-- 엑셀파일로 저장 헤더 부분 -->
<%
Response.ContentType = "application/vnd.ms-excel"
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="+"equipment.xls"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body>

<!-- 리스트 시작 -->
<table align="center" border=1 cellspacing="1" bordercolor="black">
<tr bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oequip.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oequip.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>장비코드</td>
	<td>회계자산<br>관리코드</td>
	<td>자산구분</td>
	<td>손익부서</td>
	<td>위치</td>
	<td>사용자(사용처)</td>
	<td>장비구분</td>
	<td>제품명</td>
	<td>구매일자</td>
	<td>구매처</td>
	<td>구매원가</td>    	
	<td>구매가</td>
	<td>자산가치</td>
	<td>상태</td>
	<td>사용여부</td>
</tr>
<% if oequip.FResultCount > 0 then %>
<% for i=0 to oequip.FResultCount - 1 %>
<%
	totalcurrsum = totalcurrsum + oequip.FItemList(i).Fbuy_sum
	totaljasan	 = totaljasan + oequip.FItemList(i).GetCurrentvalue
%>
<tr align="center" bgcolor="#FFFFFF" onMouseOver= this.style.background='f1f1f1'; onMouseOut=this.style.background='#ffffff';>
	<td><%= oequip.FItemList(i).Fequip_code %></td>
	<td><%= oequip.FItemList(i).faccountassetcode %></td>
	<td><%= oequip.FItemList(i).GetAccountGubunName %></td>
	<td><%= oequip.FItemList(i).FBIZSECTION_NM %></td>
	<td><%= oequip.FItemList(i).Flocate_gubun_name %></td>
	<td>
		<%= oequip.FItemList(i).fusingusername %>
		<% if oequip.FItemList(i).fstatediv <> "Y" then %>
			<font color="red">[퇴사]</font><Br>
		<% end if %>
		<% if oequip.FItemList(i).Fusing_userid <> "" then %>
			<%= oequip.FItemList(i).Fusing_userid %>
		<% end if %>
	</td>
	<td><%= oequip.FItemList(i).Fequip_gubun_name %></td>
	<td><%= oequip.FItemList(i).Fequip_name %></td>
	<td><%= oequip.FItemList(i).Fbuy_date %></td>
	<td><%= oequip.FItemList(i).fbuy_company_name %></td>
	<td><%= formatNumber(oequip.FItemList(i).fwonga_cost,0) %></td>
	<td><%= formatNumber(oequip.FItemList(i).Fbuy_sum,0) %></td>
	<td>
		<% if oequip.FItemList(i).getCurrentValue<>"" then %>
			<font color="#EE3333"><%=formatNumber(oequip.FItemList(i).GetCurrentvalue,0)%></font>
		<% else %>
			<%=formatNumber(oequip.FItemList(i).GetCurrentvalue,0)%>
		<% end if %>
	</td>
	<td><%= oequip.FItemList(i).fstate_name %></td>
	<td><%= oequip.FItemList(i).fisusing %></td>
</tr>   
<% next %>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan="12">총계</td>
	<td align="right"><%= formatNumber(totalcurrsum,0) %></td>
	<td align="right"><font color="#EE3333"><%= formatNumber(totaljasan,0) %></font></td>
	<td></td>
</tr>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>

<% end if %>
</table>

</body>
</html>

<%
set oequip = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->