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
dim detail_search, SearchString, search_type, menupos, lBoard, page, i, brd_isusing, adminuserid
	detail_search 	= request("detail_search")
	brd_isusing 	= request("brd_isusing")
	SearchString 	= request("SearchString")
	search_type 	= request("brd_type")
	menupos 	= request("menupos")
	page = request("page")

adminuserid = session("ssBctId")

If page = "" Then page = 1
if search_type = "" then search_type = 1
	
set lBoard = new board
	lBoard.Frectdetail_search = detail_search
	lBoard.Frectsearchstr = SearchString
	lBoard.FPageSize = 50
	lBoard.FCurrPage = page
	lBoard.frectbrd_isusing = brd_isusing
	lBoard.Frectsearch_type = search_type
	lBoard.fnBoardlist
%>

<script language="javascript">

function searching(){
	var sform = document.fsearch;
	sform.submit();
}

function goView(brd_sn){
	if( brd_sn == ""){
		alert("게시판 No.가 없습니다.");
		return;
	}
	
	location.href = "board_view.asp?brd_sn="+brd_sn;
}

function goedit(brd_sn){
	location.href = "board_edit.asp?brd_sn="+brd_sn;
}

function gosubmit(page){
    var frm = document.fsearch;

	if(frm.detail_search.value != "" && frm.SearchString.value == ""){
		frm.SearchString.focus();
		alert("검색어를 입력하세요");
		return false;
	}
	if(frm.detail_search.value == "" && frm.SearchString.value != ""){
		frm.detail_search.focus();
		alert("상세 검색 조건을 선택하세요");
		return false;
	}
	    
    frm.page.value=page;
	frm.submit();
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
			<tr>
				<td width="400" style="padding:5; border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999"  background="/images/menubar_1px.gif">
					<font color="#333333"><b>[ON]해외상품관리&gt;&gt;아이띵소해외게시판관리</b></font>
				</td>
				
				<td align="right" style="border-bottom:1px solid #999999" bgcolor="#F4F4F4">
					<!-- 마스터이상 메뉴권한 설정 -->
					
					<a href="Javascript:PopMenuEdit('1491');"><img src="/images/icon_chgauth.gif" border="0" valign="bottom"></a>
					
					<!-- Help 설정 -->
					
					<a href="Javascript:PopMenuHelp('1491');"><img src="/images/icon_help.gif" border="0" valign="bottom"></a>
					
				</td>
				
			</tr>
		</table>
	</td>
</tr>
</table>

<form name="fsearch" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page">
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 공지구분 : <%= fnBrdType( "w", "Y", "brd_type",search_type, " onchange='gosubmit("""");'") %>
		&nbsp;&nbsp;
		* 상세 검색 :
		<select name="detail_search" class="a">
			<option value="">--선택--</option>
			<option value="subject" <% If detail_search = "subject" Then response.write "selected" End If%> >제목</option>
			<option value="content" <% If detail_search = "content" Then response.write "selected" End If%> >내용</option>
			<option value="writer"  <% If detail_search = "writer" Then response.write "selected" End If%> >등록ID</option>
		</select>
		<input type="text" name="SearchString" size="30" value="<%=SearchString%>"><input type="text" style="display: none;" />
		&nbsp;&nbsp;
		* 사용여부 : <% drawSelectBoxisusingYN "brd_isusing",brd_isusing, " onchange=""gosubmit('');""" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<Br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
    </td>
    <td align="right">
    	<input type="button" class="button" value="신규등록" onclick="goedit('');">
    </td>
</tr>	
</table>
<!-- 표 중간바 끝-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= lBoard.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= lBoard.FTotalPage %></b>
		</td>
	</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>번호</td>
	<td>공지구분</td>
	<td>등록ID</td>
	<td>제목</td>
	<td>등록일</td>
	<td>사용여부</td>
	<td>조회수</td>
	<td>비고</td>
</tr>
<%
if lBoard.fresultcount >0 then
	
For i = 0 to lBoard.fresultcount -1
%>
<tr height="25" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" >
	<td align="center" width="50" onClick="goView('<%=lboard.FItemList(i).Fbrd_sn%>')" style="cursor:pointer"><%=lboard.FItemList(i).Fbrd_sn%></td>
	<td align="center" width="70" onClick="goView('<%=lboard.FItemList(i).Fbrd_sn%>')" style="cursor:pointer"><%= fnBrdType("v", "","", lboard.FItemList(i).fbrd_type, "") %></td>
	<td align="center" width="70" onClick="goView('<%=lboard.FItemList(i).Fbrd_sn%>')" style="cursor:pointer"><%=lboard.FItemList(i).fuserid%></td>	
	<td onClick="goView('<%=lboard.FItemList(i).Fbrd_sn%>')" style="cursor:pointer">
		<%
			If lboard.FItemList(i).Fbrd_fixed = "1" Then
				response.write "<b>"&lboard.FItemList(i).Fbrd_subject&"</b>"
			Else
				response.write lboard.FItemList(i).Fbrd_subject
			End If
		%>
	</td>
	<td align="center" width="70" onClick="goView('<%=lboard.FItemList(i).Fbrd_sn%>')" style="cursor:pointer"><%=left(lboard.FItemList(i).Fbrd_regdate,10)%></td>
	<td align="center" width="70" onClick="goView('<%=lboard.FItemList(i).Fbrd_sn%>')" style="cursor:pointer"><%=lboard.FItemList(i).fbrd_isusing%></td>
	<td align="center" width="70" onClick="goView('<%=lboard.FItemList(i).Fbrd_sn%>')" style="cursor:pointer"><%=lboard.FItemList(i).Fbrd_hit%></td>
	<td align="center" width="70">
		<%' if lboard.FItemList(i).fuserid = adminuserid or C_ADMIN_AUTH then %>
			<input type="button" class="button" value="수정" onclick="goedit('<%=lboard.FItemList(i).Fbrd_sn%>');">
		<%' end if %>
	</td>
</tr>
<%
	Next
%>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If lboard.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= lboard.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + lboard.StartScrollPage to lboard.StartScrollPage + lboard.FScrollCount - 1 %>
			<% If (i > lboard.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(lboard.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If lboard.HasNextScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
<%ELSE%>	
<tr bgcolor='FFFFFF' align='center'>
	<td colspan=15>등록된 내역이 없습니다</td>
</tr>		
<% end if %>
</table>

<% set lBoard = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->