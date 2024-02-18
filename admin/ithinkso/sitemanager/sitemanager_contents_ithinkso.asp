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
dim reload , ix, oContents, ocode
dim imagepath3, imagepath2, code, codename, imagetype, imagewidth, imageheight, isusing, idx
dim imagepath, linkpath, regdate, imagecount, image_order, lastdate, regadminid, lastupdateadminid
	idx = request("idx")
	code = request("code")
	reload = request("reload")

if reload="on" then
    session.codePage = 949
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End    
end if

set ocode = new csitemanager_list
	ocode.frectcode = code
	ocode.frectisusing = "Y"
	
	if (code<>"") then
	    ocode.fsitemanager_code_one()
	end if
	
set oContents = new csitemanager_list
	if idx <> "" then
		oContents.FRectIdx = idx
		oContents.fsitemanager_one
		
		if oContents.ftotalcount > 0 then
    		imagepath3 = oContents.FOneItem.fimagepath3
    		imagepath2 = oContents.FOneItem.fimagepath2
			code = oContents.FOneItem.fcode
			codename = oContents.FOneItem.fcodename
			imagetype = oContents.FOneItem.fimagetype
			imagewidth = oContents.FOneItem.fimagewidth
			imageheight = oContents.FOneItem.fimageheight
			isusing = oContents.FOneItem.fisusing
			idx = oContents.FOneItem.fidx
			imagepath = oContents.FOneItem.fimagepath
			linkpath = oContents.FOneItem.flinkpath
			regdate = oContents.FOneItem.fregdate
			imagecount = oContents.FOneItem.fimagecount
			image_order = oContents.FOneItem.fimage_order
 			lastdate = oContents.FOneItem.flastdate
			regadminid = oContents.FOneItem.fregadminid
			lastupdateadminid = oContents.FOneItem.flastupdateadminid
		end if
	end if
set oContents = Nothing

if isusing="" then isusing="Y"
%>

<script language='javascript'>

function SaveMainContents(frm){
    if (frm.code.value.length<1){
        alert('구분을 먼저 선택 하세요.');
        frm.code.focus();
        return;
    }
    if (frm.image_order.value.length<1){
        alert('이미지 우선순위를 입력 하세요.');
        frm.image_order.focus();
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

<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/ithinkso/image_proc.asp" onsubmit="return false;" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="ckUserId" value="<%=session("ssBctId")%>">
<input type="hidden" name="mode" value="contentsreg">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
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
    <td width="150" align="center">구분명 :</td>
    <td>
        <% if idx<>"" then %>
			<%= codename %> (<%= code %>)
			<input type="hidden" name="code" value="<%= code %>">
        <% else %>
			<% call DrawsitemanagerCode("code", code, " onChange='ChangeGubun(this.value, """&idx&""" );'") %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" align="center">이미지정렬우선순위 :</td>
    <td>
        <% if idx<>"" then %>
			<select name="image_order">
				<option>선택</option>
				<% for ix = 1 to 100 %>
					<option value="<%=ix%>" <% if cint(image_order) = cint(ix) then response.write " selected"%>><%= ix %></option>				
				<% next %>						
			</select>
        <% else %>
            <% if code<>"" then %>
				<select name="image_order">
					<option>선택</option>
					<% for ix = 1 to 100 %>
						<option value="<%=ix%>"><%= ix %></option>				
					<% next %>						
				</select>
				실서버 적용시 숫자가 작을경우 우선노출
            <% else %>
            <font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
        <% end if %>
    </td>
</tr>	
<tr bgcolor="#FFFFFF">
    <td width="150" align="center">링크구분 :</td>
    <td>
        <% if idx<>"" then %>
        <%= imagetype %>
        <% else %>
            <% if code<>"" then %>
            <%= ocode.FOneItem.fimagetype %>
            <% else %>
            <font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="150" align="center">메인이미지 :</td>
	<td>
		<input type="file" name="file1" value="" size="32" maxlength="32" class="file">
		
		<% if imagepath<>"" then %>
			<br><img src="<%=uploadUrl%>/ithinkso/sitemanager/<%= imagepath %>" border="0"> 
			<br><%=uploadUrl%>/ithinkso/sitemanager/<%= imagepath %>
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="150" align="center">서브이미지 :</td>
	<td>
		<input type="file" name="file2" value="" size="32" maxlength="32" class="file"> ※필요한경우에만등록
		
		<% if imagepath2<>"" then %>
			<br><img src="<%=uploadUrl%>/ithinkso/sitemanager/<%= imagepath2 %>" border="0"> 
			<br><%=uploadUrl%>/ithinkso/sitemanager/<%= imagepath2 %>
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="150" align="center">서브이미지2 :</td>
	<td>
		<input type="file" name="file3" value="" size="32" maxlength="32" class="file"> ※필요한경우에만등록
		
		<% if imagepath3<>"" then %>
			<br><img src="<%=uploadUrl%>/ithinkso/sitemanager/<%= imagepath3 %>" border="0"> 
			<br><%=uploadUrl%>/ithinkso/sitemanager/<%= imagepath3 %>
		<% end if %>
	</td>
</tr>	
<tr bgcolor="#FFFFFF">
	<td width="150" align="center">사용할 이미지수 :</td>
	<td>
	    <% if code<>"" then %>
			<%= imagecount %>
	    <% else %>
			<font color="red">구분을 먼저 선택하세요</font>
	    <% end if %>
	</td>
</tr>	
<tr bgcolor="#FFFFFF">
	<td width="150"  align="center">이미지Width :</td>
	<td>
		<% if code<>"" then %>
			<%= imagewidth %>
		<% else %>
			<font color="red">구분을 먼저 선택하세요</font>
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="150" align="center">이미지Height :</td>
	<td>
		<% if code<>"" then %>
			<%= imageheight %>
		<% else %>
			<font color="red">구분을 먼저 선택하세요</font>
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" align="center">링크값 :</td>
    <td>
        <% if idx<>"" then %>
            <% if imagetype="map" then %>
            	<textarea name="linkpath" cols="60" rows="6"><%= linkpath %></textarea>
            <% else %>
            	<input type="text" name="linkpath" value="<%= linkpath %>" maxlength="128" size="128">
            <% end if %>
        <% else %>
            <% if code<>"" then %>
                <% if ocode.FOneItem.fimagetype="map" then %>
                    <textarea name="linkpath" cols="60" rows="6"><map name='Map1'></map></textarea>
                    <br>(이미지맵 변수값 변경 금지)
                <% else %>
                    <input type="text" name="linkpath" value="" maxlength="128" size="128">
                    <br>(상대경로로 표시해 주세요  ex: /culturestation/culturestation_event.asp?evt_code=7)
                <% end if %>
            <% else %>
            <font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
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
    <td  align="center" colspan=2>
    	<input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);" class="button">
    </td>
</tr>	
</table>

</form>

<%
session.codePage = 949

set ocode = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->