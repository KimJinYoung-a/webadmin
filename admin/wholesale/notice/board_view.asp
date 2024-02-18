<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'###########################################################
' Description : 텐바이텐 대량구매 사이트 게시판 관리
' Hieditor : 2013.05.09 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/wholesale/notice/boardCls.asp"-->

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

function frm_list(){
	location.href='board_list.asp?menupos=<%=menupos%>';
}

</script>


<link rel="stylesheet" href="/js/js_minical.css" type="text/css"  />

<table border="0" width="100%" cellpadding="0" cellspacing="0" class="a">
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
							<td colspan=2><%= fnBrdType("v", "Y","brd_type", brd_type, "") %></td>
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
				<%= nl2br(brd_subject) %>
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">내용</td>
			<td bgcolor="#FFFFFF" style="padding: 5 4 2 5">
				<%= brd_content %>
			</td>
		</tr>	
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">고정여부</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<% if brd_fixed = "1" then %>
					Y
				<% else %>
					N
				<% end if %>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<%= brd_isusing %>
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
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<% set mBoard = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
