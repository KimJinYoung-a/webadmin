<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'###########################################################
' Description : 아이띵소 게시판 관리
' Hieditor : 2013.05.09 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/ithinkso/notice/boardCls.asp"-->

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
dim userid, brd_subject, brd_content, brd_hit, brd_regdate, brd_fixed, brd_isusing, brd_type, brd_sn
dim menupos, mBoard, lastuserid, brd_lastupdate, adminuserid
	brd_sn = request("brd_sn")
	menupos 	= request("menupos")

adminuserid = session("ssBctId")

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	
set mBoard = new Board
	mBoard.Frectbrd_sn = brd_sn
	
	if brd_sn <> "" then
		mBoard.fnBoardmodify
	end if

	if mBoard.ftotalcount > 0 then
		brd_sn =  mBoard.FOneItem.fbrd_sn
		userid = mBoard.FOneItem.fuserid
		lastuserid = mBoard.FOneItem.flastuserid
		brd_lastupdate = mBoard.FOneItem.fbrd_lastupdate
		brd_subject = mBoard.FOneItem.Fbrd_subject
		brd_content = mBoard.FOneItem.Fbrd_content
		brd_hit = mBoard.FOneItem.Fbrd_hit
		brd_regdate = mBoard.FOneItem.Fbrd_regdate
		brd_fixed = mBoard.FOneItem.Fbrd_fixed
		brd_isusing = mBoard.FOneItem.Fbrd_isusing
		brd_type = mBoard.FOneItem.Fbrd_type
		
	end if

if brd_type = "" then brd_type = 1
if brd_isusing = "" then brd_isusing = "Y"
if brd_fixed = "" then brd_fixed = 2
%>

<script language="javascript">

function frm_check(){
	var frm = document.frm;

	if(frm.brd_type.value == "")
	{
		alert("공지 구분을 선택하세요");
		return false;
	}
	
	if(frm.brd_subject.value == ""){
		alert("제목을 입력하세요");
		frm.brd_subject.focus();
		return false;
	}

	// 이노디터로 저장한 값을 textarea에 할당 시작
	var strHTMLCode = fnGetEditorHTMLCode(true, 0);
	if(strHTMLCode == ''){
		alert("내용을 입력하세요");	
		return false;
	}else{
		frm["brd_content"].value = strHTMLCode;	
	}
	// 이노디터로 저장한 값을 textarea에 할당 끝
	
	frm.action = "board_proc.asp";
	frm.submit();
}

function frm_list(){
	location.href='board_list.asp?menupos=<%=menupos%>';
}

</script>

<!-- 이노디터 인크루드 JS -->
<script language="javascript" type="text/javascript">
	var g_arrSetEditorArea = new Array();
	g_arrSetEditorArea[0] = "EDITOR_AREA_CONTAINER";
</script>
<script language="javascript" type="text/javascript" src="/lib/util/innoditor_u/js/customize.js"></script>
<script language="javascript" type="text/javascript" src="/lib/util/innoditor_u/js/customize_ui.js"></script>
<script language="javascript" type="text/javascript" src="/lib/util/innoditor_u/js/loadlayer.js"></script>
<script language="javascript" type="text/javascript">
	//이노디터에서 업로드 할 URL설정
	//Fd로 저장될 폴더를 파라메타로 넘기고 webimage에서 폴더를 만들어줘야한다.///webimage/innoditor/파라메타값
	var g_strUploadImageURL = "/lib/util/innoditor/pop_upload_img.asp?Fd=ithinkso_notice";

	// 크기, 높이 재정의
	g_nEditorWidth = 800;
	g_nEditorHeight = 800;
</script>
<!-- 이노디터 인크루드 JS 끝 -->

<link rel="stylesheet" href="/js/js_minical.css" type="text/css"  />

<table border="0" width="100%" cellpadding="0" cellspacing="0" class="a">
<form name="frm" method="post">
<input type="hidden" name="mode" value="brdreg">
<input type="hidden" name="brd_fixed" value="<%=brd_fixed%>">
<input type="hidden" name="brd_isusing" value="<%=brd_isusing%>">
<textarea name="brd_content" rows="0" cols="0" style="display:none"><%=brd_content%></textarea> <!-- 실제 이노디터 에디터의 값이 저장되는 부분(에디터에 저장한 것이 textarea에 stlye:none으로 저장 -->
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">번호</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				No.<%=brd_sn%><input type = "hidden" name="brd_sn" value="<%= brd_sn %>">
			</td>
		</tr>
		
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">공지구분</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td>
						<table class="a">
						<tr>
							<td colspan=2><%= fnBrdType("w", "Y","brd_type", brd_type, "") %></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">제 목</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="text" class="text" name="brd_subject" value="<%= brd_subject %>" size="95" maxlength="128">
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">내용</td>
			<td bgcolor="#FFFFFF" style="padding: 5 4 2 5">
				<div id="EDITOR_AREA_CONTAINER"></div>
			</td>
		</tr>	
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">고정여부</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<label><input type="radio" onclick="document.getElementById('brd_fixed').value = 1;" name="tmpbrd_fixed" value="1" <% If brd_fixed = "1" Then response.write "checked" End If %>>Y</label>&nbsp;&nbsp;&nbsp;
				<label><input type="radio" onclick="document.getElementById('brd_fixed').value = 2;" name="tmpbrd_fixed" value="2" <% If brd_fixed = "2" Then response.write "checked" End If %> >N</label><br>
				<font color = "RED"> ※Y를 선택하시면 게시글의 최상단에 위치하게 됩니다.</font>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<label id="brd_isusing"><input type="radio" onclick="document.getElementById('brd_isusing').value = 'Y';" name="tmpbrd_isusing" <% If brd_isusing = "Y" Then response.write "checked" End If %> value="Y">Y</label>&nbsp;&nbsp;&nbsp;
				<label id="brd_isusing"><input type="radio" onclick="document.getElementById('brd_isusing').value = 'N';" name="tmpbrd_isusing" <% If brd_isusing = "N" Then response.write "checked" End If %> value="N">N</label><br>
			</td>
		</tr>
		
		<% if userid <> "" then %>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">등록</td>
				<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%=userid%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;※ 등록일: <%= brd_regdate %>
				</td>
			</tr>
		<% end if %>
		<% if lastuserid <> "" then %>			
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">최근수정</td>
				<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%=lastuserid%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;※ 수정일: <%= brd_lastupdate %>
				</td>
			</tr>			
		<% end if %>			
	
		<tr bgcolor="#FFFFFF">
			<td colspan=2 align="right">
				<input type="button" onclick="frm_list();" value="목록으로" class="button">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" onclick="frm_check();" value="저장" class="button">
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>

<% if brd_sn <> "" then %>
	<!-- 글 수정시 textarea에 값 전달 시작 -->
	<script>
		var strHTMLCode = document.frm["brd_content"].value;
		fnSetEditorHTMLCode(strHTMLCode, false, 0);
	</script>
	<!-- 글 수정시 textarea에 값 전달 끝 -->
<% end if %>

<% set mBoard = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
