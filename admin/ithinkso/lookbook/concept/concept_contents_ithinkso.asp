<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'###########################################################
' Description : 아이띵소 룩북 관리
' Hieditor : 2013.05.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/ithinkso/lookbook/lookbook_cls_ithinkso.asp"-->

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
dim ix, olookbook
dim idx, lookbookgubun, title, contents, imagemain, imagemain_over, linkpath
dim isusing, regdate, lastdate, regadminid, lastupdateadminid
	idx = request("idx")

set olookbook = new clookbook_list
	if idx <> "" then
		olookbook.FRectIdx = idx
		olookbook.flookbook_concept_one
		
		if olookbook.ftotalcount > 0 then
			idx = olookbook.FOneItem.fidx
			lookbookgubun = olookbook.FOneItem.flookbookgubun
			title = olookbook.FOneItem.ftitle
			contents = olookbook.FOneItem.fcontents
			imagemain = olookbook.FOneItem.fimagemain
			imagemain_over = olookbook.FOneItem.fimagemain_over
			linkpath = olookbook.FOneItem.flinkpath
			isusing = olookbook.FOneItem.fisusing
			regdate = olookbook.FOneItem.fregdate
			lastdate = olookbook.FOneItem.flastdate
			regadminid = olookbook.FOneItem.fregadminid
			lastupdateadminid = olookbook.FOneItem.flastupdateadminid
		end if
	end if
set olookbook = Nothing

if isusing="" then isusing="Y"
%>

<script language='javascript'>

function SaveMainContents(frm){
    if (frm.title.value.length<1){
        alert('제목을 입력 하세요.');
        frm.title.focus();
        return;
    }
    if (frm.isusing.value==''){
        alert('사용여부를 선택 하세요.');
        frm.isusing.focus();
        return;
    }
    
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}

function ChangeGubun(comp, idx){
    location.href = "?code=" + comp + "&idx=" + idx;
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="center">
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmcontents" method="post" action="<%=staticimgurl%>/linkweb/ithinkso/concept_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="ckUserId" value="<%=session("ssBctId")%>">
<input type="hidden" name="mode" value="conceptreg">
<tr bgcolor="#FFFFFF">
    <td width="150" align="center">Idx :</td>
    <td>
        <% if idx<>"" then %>
        	<%= idx %>
        	<input type="hidden" name="idx" value="<%= idx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" align="center">제목 :</td>
    <td>
        <input type="text" name="title" value="<%= title %>" size=80>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="150" align="center">메인이미지 :</td>
	<td>
		<input type="file" name="imagemain" value="" size="32" maxlength="32" class="file">
		
		<% if imagemain<>"" then %>
			<br><img src="<%=imagemain %>" border="0"> 
			<br><%= imagemain %>
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="150" align="center">메인오버이미지 :</td>
	<td>
		<input type="file" name="imagemain_over" value="" size="32" maxlength="32" class="file">
		
		<% if imagemain_over<>"" then %>
			<br><img src="<%= imagemain_over %>" border="0"> 
			<br><%= imagemain_over %>
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="150" align="center">메인링크파일 :</td>
	<td>
		<input type="file" name="linkpath" value="" size="32" maxlength="32" class="file">
		
		<% if linkpath<>"" then %>
			<br><a href="<%= linkpath %>" onfocus="this.blur()"><%= linkpath %></a>
		<% end if %>
	</td>
</tr>
<% if regdate <> "" then %>
<tr bgcolor="#FFFFFF">
    <td width="150" align="center">등록일 :</td>
    <td>
        <%= regdate %> (<%= regadminid %>)
    </td>
</tr>
<% end if %>
<% if lastdate <> "" then %>
<tr bgcolor="#FFFFFF">
    <td width="150" align="center">수정일 :</td>
    <td>
        <%= lastdate %> (<%= lastupdateadminid %>)
    </td>
</tr>
<% end if %>

<tr bgcolor="#FFFFFF">
    <td width="150" align="center">사용여부 :</td>
    <td>
        <% drawSelectBoxisusingYN "isusing", isusing, "" %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center" colspan=2>
    	<input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);" class="button">
    </td>
</tr>	
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->