<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/admin/Helpmenucls.asp"-->
<%
dim menupos
menupos = requestCheckVar(request("menupos"),10)


dim imenupos, menuposStr
set imenupos = new CMenuList
imenupos.FRectID = menupos

if menupos<>"" then
	imenupos.getOneMenu
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">  
<style type="text/css">
<!--
	body {  font-size: 9pt}
	td {  font-size: 9pt}
	a {  text-decoration: none}
-->
</style>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script> 
<!-- daumeditor head ------------------------->
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
 <meta http-equiv="X-UA-Compatible" content="IE=10" /> 
 <link rel="stylesheet" href="/lib/util/daumeditor/css/editor.css" type="text/css" charset="utf-8"/>
 <script src="/lib/util/daumeditor/js/editor_loader.js" type="text/javascript" charset="utf-8"></script>
 <script src="/lib/util/daumeditor/js/editor_creator.js" type="text/javascript" charset="utf-8"></script>
 <script type="text/javascript">
    var config = {
        initializedId: "",
        wrapper: "tx_trex_container",
        form: 'frmSubmit',
        txIconPath: "/lib/util/daumeditor/images/icon/editor/",
        txDecoPath: "/lib/util/daumeditor/images/deco/contents/",
        events: {
            preventUnload: false
        },
        sidebar: {
            attachbox: {
                show: true
            },
            attacher: {
                 image: {
                    popPageUrl: "/lib/util/daumeditor/pages/trex/image.asp"
                } 
            }
        }
    }; 
 </script>
<!-- //daumeditor head ------------------------->
<script language='javascript'>
function SaveMenuEtc(frm){
	if (frm.menuname.value.length<1){
		alert('메뉴명을 입력하세요.');
		frm.menuname.focus();
		return;
	}

	if (frm.viewidx.value.length<1){
		alert('전시순서를 입력하세요.');
		frm.viewidx.focus();
		return;
	}

 	var content = Editor.getContent();
//	  	var str = chkContent(content); 
//		  if (str) {
//		   alert("script태그 또는 form태그는 사용할 수 없는 문자열 입니다.");
//		   return ;  
//		  } 
      document.getElementById("menuposhelp").value = content; 
      
	if (confirm('저장 하시겠습니까?')){
		frm.submit();
	}
}
</script>
</head>
<body topmargin="0">
<% if imenupos.FResultCount>0 then %>
	<!-- 표 상단바 시작-->
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
		</tr>
		<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        	<%= imenupos.FOneItem.FMenuStr %>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
		</tr>
	</table>
	<!-- 표 상단바 끝-->
	
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#bababa">

		<% if imenupos.FOneItem.Fmenuposnotice<>"" then %>
		<tr bgcolor="<%= adminColor("tabletop") %>">
			<td>
				<img src="/images/icon_arrow_down.gif" align="absbottom">
				<font color="red"><strong>간단설명</strong></font><br>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>
				<%= nl2br(imenupos.FOneItem.Fmenuposnotice) %>
			</td>
		</tr>
		<% end if %>
	
		<% if imenupos.FOneItem.Fmenuposhelp<>"" then %>
		<tr bgcolor="<%= adminColor("tabletop") %>">
			<td>
				<img src="/images/icon_arrow_down.gif" align="absbottom">
				<font color="red"><strong>상세설명</strong></font><br>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>
				<%= nl2br(imenupos.FOneItem.Fmenuposhelp) %>
			</td>
		</tr>
		<% end if %>
		
		<% if imenupos.FOneItem.Fmenuposnotice="" and imenupos.FOneItem.Fmenuposhelp="" then %>
		<tr align="center" bgcolor="#FFFFFF">
			<td>
				내용이 없습니다.
			</td>
		</tr>
		<% end if %>
		
	</table>
	
	<!-- 표 하단바 시작-->
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	    <tr valign="bottom" height="25">
	        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="bottom" align="center">&nbsp;</td>
	        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	    </tr>
	    <tr valign="top" height="10">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_08.gif"></td>
	        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	    </tr>
	</table>
	<!-- 표 하단바 끝-->


<p>

	<% if C_ADMIN_AUTH then %>
	<!-- 표 상단바 시작-->
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
		</tr>
		<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        관리자 메뉴
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
		</tr>
	</table> 
	
	<table width="100%" border="0" align="center" cellspacing="1" class="a" bgcolor="#bababa">
	   	<tr height="10" valign="bottom"> 
		<td style="padding-left:20px;"  bgcolor="<%= adminColor("topbar") %>">
	<!-- 표 상단바 끝-->
	<form name="frmSubmit" method="post" action="do_menuhelpedit.asp">
		<input type="hidden" name="id" value="<%= imenupos.FOneItem.FMenuid %>"> 
	<table width="900" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>"> 
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">메뉴명</td>
			<td bgcolor="#FFFFFF"><input type="text" name="menuname" size="50" value="<%= imenupos.FOneItem.FMenuName %>"></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">전시순서</td>
			<td bgcolor="#FFFFFF"><input type="text" name="viewidx" size="6" value="<%= imenupos.FOneItem.FViewIndex %>"></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">링크</td>
			<td bgcolor="#FFFFFF"><input type="text" name="linkurl" size="50" value="<%= imenupos.FOneItem.FLinkURL %>"></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">Color</td>
			<td bgcolor="#FFFFFF"><input type="text" name="menucolor" size="7" value="<%= imenupos.FOneItem.Fmenucolor %>"></td>
		</tr>
	
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">메뉴사용여부</td>
			<td bgcolor="#FFFFFF">
				<% if imenupos.FOneItem.Fisusing="Y" then %>
				<input type="radio" name="isusing" value="Y" checked > 사용함
				<input type="radio" name="isusing" value="N"> 사용안함
				<% else %>
				<input type="radio" name="isusing" value="Y" > 사용함
				<input type="radio" name="isusing" value="N" checked > <font color="red">사용안함</font>
				<% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">간단설명</td>
			<td bgcolor="#FFFFFF">
				<textarea name="menuposnotice" cols="90" rows="8"><%= imenupos.FOneItem.Fmenuposnotice %></textarea>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">상세설명</td>
			<td bgcolor="#FFFFFF"> 
					<textarea name="menuposhelp" id="menuposhelp" style="width: 100%; height: 490px;"><%=imenupos.FOneItem.Fmenuposhelp%></textarea>  
          	<!-- daumeditor  --> 
          	<script type="text/javascript">  
          	    EditorCreator.convert(document.getElementById("menuposhelp"), '/lib/util/daumeditor/teneditor/editorForm.html', function () {
          	        EditorJSLoader.ready(function (Editor) {
          	            new Editor(config);
          	            Editor.modify({
          	                content: document.getElementById("menuposhelp") 
          	            });
          	        });
          	    });
          	
          	</script> 
          	<!-- daumeditor   -->
			</td>
		</tr> 
	</table> 
	</form>
</td>
</tr>
</table>
	<!-- 표 하단바 시작-->
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	    <tr valign="bottom" height="25">
	        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="bottom" align="center">
	        	<input type="button" value=" 저 장 " onclick="SaveMenuEtc(frmSubmit);">
	        </td>
	        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	    </tr>
	    <tr valign="top" height="10">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_08.gif"></td>
	        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	    </tr>
	</table>
	<!-- 표 하단바 끝-->
	
	<% end if %>
<% end if %>
</body>
</html>


<%
set imenupos = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	session.codePage = 949
%>
