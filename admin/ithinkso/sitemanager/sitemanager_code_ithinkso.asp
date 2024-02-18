<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<% session.codePage = 65001 %>
<%
'###########################################################
' Description : 아이띵소 사이트 관리
' Hieditor : 2013.05.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

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

<!-- #include virtual="/lib/classes/ithinkso/sitemanager/sitemanager_cls_ithinkso_utf8.asp"-->

<%
dim page, menupos, ocodeone, ocodeList, i
dim code,codetype, codename, imagetype, imagewidth, imageheight, isusing, imagecount
dim regdate, lastdate, regadminid, lastupdateadminid
	menupos = request("menupos")
	code = request("code")
	page = request("page")

if page="" then page=1

set ocodeone = new csitemanager_list
	ocodeone.FRectcode = code
	
	if code <> "" then
		ocodeone.fsitemanager_code_one()
		
		if ocodeone.ftotalcount > 0 then
			code = ocodeone.FOneItem.fcode
			codetype = ocodeone.FOneItem.fcodetype
			codename = ocodeone.FOneItem.fcodename
			imagetype = ocodeone.FOneItem.fimagetype
			imagewidth = ocodeone.FOneItem.fimagewidth
			imageheight = ocodeone.FOneItem.fimageheight
			isusing = ocodeone.FOneItem.fisusing
			imagecount = ocodeone.FOneItem.fimagecount
			regdate = ocodeone.FOneItem.fregdate
			lastdate = ocodeone.FOneItem.flastdate
			regadminid = ocodeone.FOneItem.fregadminid
			lastupdateadminid = ocodeone.FOneItem.flastupdateadminid
		end if
	end if
set ocodeone = Nothing

set ocodeList = new csitemanager_list
	ocodeList.FPageSize=50
	ocodeList.FCurrPage= page
	ocodeList.fsitemanager_code_list()

%>

<script language='javascript'>

function Savecode(frm){
	if (!IsDouble(frm.code.value)){
		alert('코드는 숫자만 가능합니다.');
		frm.code.focus();
		return;
	}
    if (frm.codename.value.length<1){
        alert('코드명을 입력하세요.');
        frm.codename.focus();
        return;
    }
	if (!IsDouble(frm.imagecount.value)){
		alert('이미지수는 숫자만 가능합니다.');
		frm.imagecount.focus();
		return;
	}
	if (!IsDouble(frm.imagewidth.value)){
		alert('이미지 가로 사이즈는 숫자만 가능합니다.');
		frm.imagewidth.focus();
		return;
	}
    if (frm.imagetype.value.length<1){
        alert('링크타입을 선택하세요.');
        frm.imagetype.focus();
        return;
    }
    if (frm.isusing.value==''){
        alert('사용여부를 입력하세요.');
        frm.isusing.focus();
        return;
    }

    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}

function codeedit(page,code){
	location.href="?code=" + code + "&page="+page;
}

</script>

<form name="frmcode" method="post" action="/admin/ithinkso/sitemanager/sitemanager_code_process_ithinkso.asp" style="margin:0px;">
<input type="hidden" name="mode" value="codereg">

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">코드</td>
    <td align="left">
    	<% if code <> "" then %>
        	<%= code %>
        	<input type="hidden" name="orgcode" value="<%= code %>" >
        	<input type="hidden" name="code" value="<%= code %>" >
        <% else %>
        	<input type="text" name="code" value="<%= code %>" maxlength="7" size="5">
        	(숫자)
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">코드명</td>
    <td align="left">
        <input type="text" name="codename" value="<%= codename %>" maxlength="32" size="64">
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">이미지수</td>
    <td align="left">
        <input type="text" name="imagecount" value="<%= imagecount %>" maxlength="2" size="2">
        (숫자만 입력 하세요)
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">이미지 WIDTH</td>
    <td align="left">
        <input type="text" name="imagewidth" value="<%= imagewidth %>" maxlength="16" size="8">
        (이미지 Width Size 숫자)
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">이미지 HEIGHT</td>
    <td align="left">
        <input type="text" name="imageheight" value="<%= imageheight %>" maxlength="16" size="8">
        (이미지 Height Size 숫자 : 0 인경우 height 지정 안함)
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">링크타입</td>
    <td align="left">
        <select name="imagetype">
	        <option value="" <% if imagetype = "" then response.write " selected" %>>CHOICE</option>
	        <option value="map" <% if imagetype = "map" then response.write " selected" %>>map</option>
	        <option value="link" <% if imagetype = "link" then response.write " selected" %>>link</option>                
	        <option value="flash" <% if imagetype = "flash" then response.write " selected" %>>flash</option>
	        <option value="multi" <% if imagetype = "multi" then response.write " selected" %>>multi</option>
	        <option value="xml" <% if imagetype = "xml" then response.write " selected" %>>xml</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td width="150">사용여부</td>
    <td align="left">
        <% drawSelectBoxisusingYN "isusing", isusing, "" %>
    </td>
</tr>
</table>

</form>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" value="저장" onClick="Savecode(frmcode);" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<br><br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ocodeList.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= ocodeList.FTotalPage %></b>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="?code="><img src="/images/icon_new_registration.gif" border="0"></a>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>code</td>
    <td>구분명</td>
    <td>링크타입</td>
    <td>Image수</td>
    <td>사용여부</td>
    <td>비고</td>
</tr>
<% if ocodeList.FResultCount > 0 then %>
<% for i=0 to ocodeList.FResultCount-1 %>
	<% if CStr(ocodeList.FItemList(i).Fcode)=cstr(code) then %>
		<tr bgcolor="orange" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='orange'; align="center">
	<% else %>
		<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff'; align="center">
	<% end if %>
	
    <td><%= ocodeList.FItemList(i).Fcode %></td>
    <td align="left"><%= ocodeList.FItemList(i).FcodeName %></td>
    <td><%= ocodeList.FItemList(i).fimagetype %></td>
    <td><%= ocodeList.FItemList(i).fimagecount %></td>
    <td><%= ocodeList.FItemList(i).Fisusing %></td>
    <td>
    	<input type="button" value="수정" onclick="codeedit('<%=page%>','<%= ocodeList.FItemList(i).Fcode %>');" class="button">
    </td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	    <% if ocodeList.HasPreScroll then %>
			<a href="?page=<%= ocodeList.StartScrollPage-1 %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
	
		<% for i=0 + ocodeList.StartScrollPage to ocodeList.FScrollCount + ocodeList.StartScrollPage - 1 %>
			<% if i>ocodeList.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>">[<%= i %>]</a>
			<% end if %>
		<% next %>
	
		<% if ocodeList.HasNextScroll then %>
			<a href="?page=<%= i %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan="15">내용이 없습니다.</td>
</tr>	
<% end if %>
</table>

</body>
</html>

<%
session.codePage = 949

set ocodeList = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
