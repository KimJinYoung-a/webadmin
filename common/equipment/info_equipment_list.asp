<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 정보자산관리
' History : 2015-06-04, skyer9
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/common/equipment/equipment_cls.asp"-->
<%
dim page, research
dim onlyinfoequip, onlyusing
dim equip_gubun, equip_code, info_gubun

''dim , part_sn, idx, using_userid, using_username, usingIp ,  ,equip_name ,manufacture_company, manufacture_sn
''dim ipgocheck, yyyy1, yyyy2, mm1, mm2, dd1, dd2, fromDate, toDate, useip ,property_gubun
''dim totalcurrsum,totaljasan, Alltotaljasan, getAllCurrentValue, part_code ,state , parameter
''dim accountGubun, department_id
''dim BIZSECTION_CD, BIZSECTION_NM
''dim only1000
''dim outcheck, yyyy3, yyyy4, mm3, mm4, dd3, dd4, fromDate2, toDate2

page = requestcheckvar(request("page"),10)
research = requestcheckvar(Request("research"),2)
onlyinfoequip = requestcheckvar(Request("onlyinfoequip"),2)
onlyusing = requestcheckvar(Request("onlyusing"),2)
equip_gubun = requestcheckvar(Request("equip_gubun"),2)
equip_code = requestcheckvar(request("equip_code"),20)
info_gubun = requestcheckvar(request("info_gubun"),2)

if (info_gubun="-1") then info_gubun=""
    
	''
	''	part_sn = requestcheckvar(Request("part_sn"),10)
	''	using_userid = requestcheckvar(Request("using_userid"),32)
	''	using_username = requestcheckvar(Request("using_username"),32)
	''	
	''	ipgocheck = requestcheckvar(request("ipgocheck"),2)
	''	yyyy1 = requestcheckvar(request("yyyy1"),4)
	''	yyyy2 = requestcheckvar(request("yyyy2"),4)
	''	mm1	  = requestcheckvar(request("mm1"),2)
	''	mm2	  = requestcheckvar(request("mm2"),2)
	''	dd1	  = requestcheckvar(request("dd1"),2)
	''	dd2	  = requestcheckvar(request("dd2"),2)
	''	part_code = requestcheckvar(Request("part_code"),10)
	''	equip_name = requestcheckvar(Request("equip_name"),64)
	''	manufacture_company = requestcheckvar(Request("manufacture_company"),64)
	''	manufacture_sn = requestcheckvar(Request("manufacture_sn"),64)
	''	property_gubun = requestcheckvar(Request("property_gubun"),10)
	''	state = requestcheckvar(Request("state"),10)
	''
	''	onlyusing = requestcheckvar(Request("onlyusing"),2)
	''	accountGubun = requestcheckvar(Request("accountGubun"),5)
	''	department_id = requestcheckvar(Request("department_id"),5)
	''	BIZSECTION_CD = requestcheckvar(Request("BIZSECTION_CD"),15)
	''	BIZSECTION_NM = requestcheckvar(Request("BIZSECTION_NM"),55)
	''	only1000 = requestcheckvar(Request("only1000"),55)

	''outcheck = requestcheckvar(request("outcheck"),2)
	''	yyyy3 = requestcheckvar(request("yyyy3"),4)
	''	yyyy4 = requestcheckvar(request("yyyy4"),4)
	''	mm3	  = requestcheckvar(request("mm3"),2)
	''	mm4	  = requestcheckvar(request("mm4"),2)
	''	dd3	  = requestcheckvar(request("dd3"),2)
	''	dd4	  = requestcheckvar(request("dd4"),2)

if page="" then page=1
''if (yyyy1="") then yyyy1 = Cstr(Year(now()))
''if (mm1="") then mm1 = Cstr(Month(now()))
''if (dd1="") then dd1 = Cstr(day(now()))
''if (yyyy2="") then yyyy2 = Cstr(Year(now()))
''if (mm2="") then mm2 = Cstr(Month(now()))
''if (dd2="") then dd2 = Cstr(day(now()))
''
''fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
''toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))
''
''if (yyyy3="") then yyyy3 = Cstr(Year(now()))
''if (mm3="") then mm3 = Cstr(Month(now()))
''if (dd3="") then dd3 = Cstr(day(now()))
''if (yyyy4="") then yyyy4 = Cstr(Year(now()))
''if (mm4="") then mm4 = Cstr(Month(now()))
''if (dd4="") then dd4 = Cstr(day(now()))
''
''fromDate2 = CStr(DateSerial(yyyy3, mm3, dd3))
''toDate2 = CStr(DateSerial(yyyy4, mm4, dd4+1))
''
if (research = "") then
	onlyusing = "Y"
end if



dim oequip,i
set oequip = new CEquipment
	oequip.FPageSize = 50
	oequip.FCurrPage = page
	oequip.FRectequip_gubun = equip_gubun
	oequip.Frectequip_code = equip_code
	oequip.FRectinfo_gubun = info_gubun
	''	oequip.FRectpart_sn = part_sn
	''	oequip.FRectusing_userid = using_userid
	''	oequip.FRectusing_username = using_username
	''	
	''	oequip.frectequip_name = equip_name
	''	oequip.frectmanufacture_company = manufacture_company
	''	oequip.frectmanufacture_sn = manufacture_sn
	''	oequip.frectproperty_gubun = property_gubun
	''	oequip.frectstate = state
	oequip.FRectIsusing = onlyusing
	''	oequip.FRectAccountGubun = accountGubun
	''	oequip.FRectDepartmentID = department_id
	''	oequip.FRectBIZSECTION_CD = BIZSECTION_CD
	oequip.FRectOnlyInfoEquip = onlyinfoequip
    
	''if ipgocheck = "on" then
	''		oequip.frectbuy_startdate = fromDate
	''		oequip.frectbuy_enddate = toDate
	''	end if
	''
	''	if outcheck = "on" then
	''		oequip.frectout_startdate = fromDate2
	''		oequip.frectout_enddate = toDate2
	''	end if

	oequip.getInfoEquipmentList

	''totalcurrsum = 0
	''totaljasan	 = 0
	''Alltotaljasan = 0
	''
	''parameter = "page="&page&"&equip_gubun="&equip_gubun&"&part_sn="&part_sn&"&using_userid="&using_userid&"&using_username="&using_username&"&equip_code="&equip_code&_
	''parameter = parameter & "&ipgocheck="&ipgocheck&"&yyyy1="&yyyy1&"&yyyy2="&yyyy2&"&mm1="&mm1&"&mm2="&mm2&"&dd1="&dd1&"&dd2="&dd2&_
	''parameter = parameter & "&part_code="&part_code&"&equip_name="&equip_name&"&manufacture_company="&manufacture_company&"&manufacture_sn="&manufacture_sn&_
	''parameter = parameter & "&property_gubun="&property_gubun&"&state="&state&"&research="&research
%>

<script language='javascript'>

//신규등록
function pop_Equipmentreg(idx){
	var pop_Equipmentreg = window.open('/common/equipment/pop_equipmentreg.asp?idx=' + idx,'pop_Equipmentreg','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_Equipmentreg.focus();
}

function NextPage(page){
	frm.page.value= page;
	frm.submit();
}

</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		자산구분 : <% DrawEquipMentGubun "10","equip_gubun",equip_gubun ," onchange='NextPage("""");'" %>
		&nbsp;
		정보자산 구분 <% drawInfoEquipmentGubun "info_gubun" ,info_gubun, " onchange='NextPage("""");'" %>
		
		&nbsp;
		자산코드 : <input type="text" name="equip_code" value="<%=equip_code%>">
		
		
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="NextPage('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		aaa
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="onlyinfoequip" value="Y" <% if (onlyinfoequip = "Y") then %>checked<% end if %> > 정보자산만 표시
		&nbsp;
		<input type="checkbox" name="onlyusing" value="Y" <% if (onlyusing = "Y") then %>checked<% end if %> > 삭제내역 제외
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">

	</td>
	<td align="right">
		<input type="button" class="button" onclick="pop_Equipmentreg('');" value="신규등록">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= oequip.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oequip.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=20><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td width=40>IDX</td>
	<td width=130>정보자산구분</td>
	<td width=130>자산코드</td>
	<td width=80>호스트명</td>
	<td width=80>사용자</td>
	<td width=100>운영체제</td>
	<td width=70>IP</td>
    <% if (FALSE) then %><td width=80>자산구분</td><% end if %>
	<td width=300>자산명</td>
	<td width=100>위치</td>
	<td width=80>장비구분</td>
	<td width=80>상태</td>
	<td width=80>폐기일자</td>
	<td width=30>사용<br>여부</td>
    <td width=30>C</td>
    <td width=30>I</td>
    <td width=30>A</td>
    <td width=30>등급</td>
    <td width=30>점수</td>
	<td>비고</td>
</tr>
<% if oequip.FResultCount > 0 then %>
<% for i=0 to oequip.FResultCount - 1 %>
<form name="frmBuyPrc_<%= i %>" >
<input type="hidden" name="idx" value="<%= oequip.FItemList(i).Fidx %>">
	<tr align="center" bgcolor="#FFFFFF" height="25">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td><%= oequip.FItemList(i).Fidx %></td>
	<td><%= GetInfoGubunCodeName(oequip.FItemList(i).Finfo_gubun) %></td>
	<td>
		<a href="javascript:pop_Equipmentreg('<%= oequip.FItemList(i).Fidx %>');" onfocus="this.blur()">
		<%= oequip.FItemList(i).Fequip_code %></a>
	</td>
	<td><%= oequip.FItemList(i).Finfo_HOSTNM %></td>
	<td><%= oequip.FItemList(i).Fusing_username %></td>
	<td><%= oequip.FItemList(i).Finfo_OS %></td>
	<td><%= SplitValue(oequip.FItemList(i).Finfo_IP,".",0)&".***.***."&SplitValue(oequip.FItemList(i).Finfo_IP,".",3) %></td>
	<% if (FALSE) then %><td><%= oequip.FItemList(i).GetAccountGubunName %></td><% end if %>
	<td align="left">
		<%= oequip.FItemList(i).Fequip_name %>
	</td>
	<td>
		<%= oequip.FItemList(i).Flocate_gubun_name %>
	</td>
	<td>
		<%= oequip.FItemList(i).Fequip_gubun_name %>
	</td>
	<td>
		<%= oequip.FItemList(i).fstate_name %>
	</td>
	<td>
		<%= oequip.FItemList(i).fout_date %>
	</td>
	<td width=30>
		<%= oequip.FItemList(i).fisusing %>
	</td>
    <td><%= getCIALevelName(oequip.FItemList(i).Finfo_importance_C) %></td>
    <td><%= getCIALevelName(oequip.FItemList(i).Finfo_importance_I) %></td>
    <td><%= getCIALevelName(oequip.FItemList(i).Finfo_importance_A) %></td>
    <td><%= oequip.FItemList(i).getCIATotalLevelName %></td>
    <td><%= oequip.FItemList(i).getCIATotalValue %></td>
	<td>
	</td>
</tr>
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
    	<% if oequip.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oequip.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oequip.StartScrollPage to oequip.FScrollCount + oequip.StartScrollPage - 1 %>
			<% if i>oequip.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oequip.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="17" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>

<%
set oequip = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
