<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매장 환율 관리
' History : 2010.08.07 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">

<script language='javascript'>

function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<% if session("sslgnMethod")<>"S" then %>
	<!-- USB키 처리 시작 (2008.06.23;허진원) -->
	<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
	<script language="javascript" src="/js/check_USBToken.js"></script>
	<!-- USB키 처리 끝 -->
<% end if %>
</head>
<body bgcolor="#F4F4F4" onload="checkUSBKey()">

<%
dim idx, sitename, currencyUnit, currencyChar, exchangeRate, basedate, regdate, lastupdate
dim reguserid, lastuserid, page, i, menupos
	currencyUnit = request("currencyUnit")
	sitename = request("sitename")
	idx = request("idx")
	menupos = request("menupos")
	page = request("page")

if page="" then page=1
	
dim oexchangerate
set oexchangerate = new COffShopChargeUser
	oexchangerate.frectidx = idx
	oexchangerate.frectcurrencyUnit = currencyUnit
	oexchangerate.frectsitename = "10X10OFFLINE"
	
	if (currencyUnit <> "" and sitename <> "") or idx <> "" then
		oexchangerate.fexchangerate_oneitem
		
		if oexchangerate.ftotalcount > 0 then
			idx = oexchangerate.FOneItem.fidx
			sitename = oexchangerate.FOneItem.fsitename
			currencyUnit = oexchangerate.FOneItem.fcurrencyUnit
			currencyChar = oexchangerate.FOneItem.fcurrencyChar
			exchangeRate = oexchangerate.FOneItem.fexchangeRate
			basedate = oexchangerate.FOneItem.fbasedate
			regdate = oexchangerate.FOneItem.fregdate
			lastupdate = oexchangerate.FOneItem.flastupdate
			reguserid = oexchangerate.FOneItem.freguserid
			lastuserid = oexchangerate.FOneItem.flastuserid
		end if
	end if

dim oexchangerateList
set oexchangerateList = new COffShopChargeUser
	oexchangerateList.FPageSize=50
	oexchangerateList.FCurrPage= page
	oexchangerateList.fexchangerate_list

if exchangeRate = "" then exchangeRate = 0	
%>

<script language='javascript'>

function SavecurrencyUnit(frm){
    if (frm.sitename.value==''){
        alert('사이트구분을 선택하세요.');
        frm.sitename.focus();
        return;
    }
    
    if (frm.currencyUnit.value==''){
        alert('화폐단위를 입력하세요.');
        frm.currencyUnit.focus();
        return;
    }

    if (frm.currencyChar.value==''){
        alert('화폐기호를 입력하세요.');
        frm.currencyChar.focus();
        return;
    }
    
    if (frm.exchangeRate.value==''){
        alert('환율을 입력하세요');
        frm.exchangeRate.focus();
        return;
    }
    
    if (frm.basedate.value==''){
        alert('기준일을 입력하세요.');
        frm.basedate.focus();
        return;
    }
        
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }    
}

//신규등록
function newcurrencyUnit(){
	location.href='/common/offshop/exchangerate/exchangerate.asp?menupos=<%=menupos%>'
}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmcurrencyUnit" method="post" action="exchangerate_process.asp">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="exchangeRateedit">
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">IDX</td>
    <td align="left">
    	<%= IDX %>
		<input type="hidden" name="idx" value="<%=idx%>">
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">사이트구분</td>
    <td align="left">
    	<%
    	'//수정모드
    	if IDX <> "" then 
    	%>
    		<%= sitename %>
    		<input type="hidden" name="sitename" value="<%= sitename %>">
		<% else %>    
			<% drawoffshop_commoncode "sitename", sitename, "sitename", "MAIN", "", "" %>
		<% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">화폐단위</td>
    <td align="left">
    	<%
    	'//수정모드
    	if IDX <> "" then 
    	%>
    		<%= currencyUnit %>
    		<input type="hidden" name="currencyUnit" value="<%= currencyUnit %>">    	
			<%' DrawexchangeRate "currencyUnit", currencyUnit,"" %>    	
		<% else %>
    		<input type="text" name="currencyUnit" value="<%= currencyUnit %>">		EX) USD
		<% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td>화폐기호</td>
    <td align="left">
        <input type="text" name="currencyChar" value="<%= currencyChar %>" maxlength="10" size="10">
        EX) $
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td>환율</td>
    <td align="left">
        <input type="text" name="exchangeRate" value="<%= exchangeRate %>" maxlength="10" size="10">
        EX) 1200
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td>기준일</td>
    <td align="left">
		<input type="text" name="basedate" size=6 maxlength=10 value="<%= basedate %>">			
		<a href="javascript:calendarOpen3(frmcurrencyUnit.basedate,'기준일',frmcurrencyUnit.basedate.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>     	
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SavecurrencyUnit(frmcurrencyUnit);" class="button"></td>
</tr>
</form>
</table>

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
    </td>
    <td align="right">
    	<input type="button" onclick="newcurrencyUnit();" value="신규등록" class="button">
    </td>
</tr>	
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= oexchangerateList.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oexchangerateList.FTotalPage %></b>
	</td>
</tr>
<% if oexchangerateList.FResultCount > 0 then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>IDX</td>
	<td>사이트구분</td>
    <td>화폐단위</td>
    <td>화폐기호</td>
    <td>환율</td>
    <td>기준일</td>
    <td>비고</td>
</tr>
<% for i=0 to oexchangerateList.FResultCount-1 %>

<% if oexchangerateList.FItemList(i).fidx = idx then %>
	<tr bgcolor="orange" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='orange'; align="center">
<% else %>
	<tr bgcolor="#ffffff" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff'; align="center">
<% end if %>
	<td><%= oexchangerateList.FItemList(i).fidx %></td>
	<td><%= oexchangerateList.FItemList(i).fsitename %></td>
    <td><%= oexchangerateList.FItemList(i).fcurrencyUnit %></td>
    <td><%= oexchangerateList.FItemList(i).fcurrencychar %></td>
    <td align="right"><%= oexchangerateList.FItemList(i).fexchangeRate %></td>
    <td><%= oexchangerateList.FItemList(i).fbasedate %></td>
    <td width=60><input type="button" onclick="location.href='?idx=<%= oexchangerateList.FItemList(i).fidx %>&page=<%= page %>'" value="수정" class="button"></td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	    <% if oexchangerateList.HasPreScroll then %>
			<a href="?page=<%= oexchangerateList.StarScrollPage-1 %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
	
		<% for i=0 + oexchangerateList.StarScrollPage to oexchangerateList.FScrollCount + oexchangerateList.StarScrollPage - 1 %>
			<% if i>oexchangerateList.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>">[<%= i %>]</a>
			<% end if %>
		<% next %>
	
		<% if oexchangerateList.HasNextScroll then %>
			<a href="?page=<%= i %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>

<% else %>
<tr bgcolor="#FFFFFF">
	<td align="center">내용이 없습니다.</td>
</tr>	
<% end if %>
</table>

</body>
</html>

<%
set oexchangerate = Nothing
set oexchangerateList = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
