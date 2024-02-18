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
dim research,isusing, page, menupos, i
	isusing = request("isusing")
	research= request("research")
	page    = request("page")
	menupos = request("menupos")

if (research="") and (isusing="") then 
    isusing = "Y"
end if

if page="" then page=1

dim olookbook
set olookbook = new clookbook_list
	olookbook.FPageSize = 50
	olookbook.FCurrPage = page
	olookbook.FRectIsusing = isusing
	olookbook.frectlookbookgubun = "1"
	olookbook.flookbook_concept_list()
%>

<script language="javascript">

//신규등록 & 수정
function AddNewMainContents(idx){
    var AddNewMainContents = window.open('/admin/ithinkso/lookbook/concept/concept_contents_ithinkso.asp?idx='+ idx,'AddNewMainContents','width=1024,height=768,scrollbars=yes,resizable=yes');
    AddNewMainContents.focus();
}

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

document.domain ='10x10.co.kr';

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
			<tr>
				<td width="400" style="padding:5; border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999"  background="/images/menubar_1px.gif">
					<font color="#333333"><b>[ON]해외상품관리&gt;&gt;아이띵소해외룩북컨셉관리</b></font>
				</td>
				
				<td align="right" style="border-bottom:1px solid #999999" bgcolor="#F4F4F4">
					<!-- 마스터이상 메뉴권한 설정
					
					<a href="Javascript:PopMenuEdit('1491');"><img src="/images/icon_chgauth.gif" border="0" valign="bottom"></a> -->
					
					<!-- Help 설정
					
					<a href="Javascript:PopMenuHelp('1491');"><img src="/images/icon_help.gif" border="0" valign="bottom"></a> -->
				</td>
				
			</tr>
		</table>
	</td>
</tr>
</table>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="fidx">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	    * 사용여부 : <% drawSelectBoxisusingYN "isusing", isusing, " onchange='frmsubmit("""")'" %>
	</td>
	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">

	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" value="신규등록" class="button" onClick="javascript:AddNewMainContents('');">						
	</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= olookbook.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= olookbook.FTotalPage %></b>	
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">IDX</td>
    <td align="center">제목</td>
    <td align="center">Image</td>
    <td align="center">사용여부</td>
    <td align="center">등록일</td>
    <td align="center">비고</td>
</tr>
<% if olookbook.FResultCount > 0 then %> 
<tr align="center" bgcolor="#FFFFFF">
<% for i=0 to olookbook.FResultCount - 1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">			 		
	<% if olookbook.FItemList(i).FIsusing="N" then %>
		<tr bgcolor="#DDDDDD" align="center">
	<% else %>
		<tr bgcolor="#FFFFFF" align="center">
	<% end if %>
	
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>		
    <td><%= olookbook.FItemList(i).Fidx %><input type="hidden" name="idx" value="<%= olookbook.FItemList(i).Fidx %>"></td>
    <td><%= olookbook.FItemList(i).ftitle %></td>
    <td>
    	<% if olookbook.FItemList(i).fimagemain <> "" then %>
    		<img width=50 height=50 src="<%= olookbook.FItemList(i).fimagemain %>" border="0">
    	<% end if %>
    </td>
    <td><%= olookbook.FItemList(i).FIsusing %></td>
    <td><%= olookbook.FItemList(i).fregdate %></td>
    <td><input type="button" value="수정" onclick="AddNewMainContents('<%= olookbook.FItemList(i).Fidx %>');" class="button"></td>
</tr>
</form>	
<% next %>
</tr>   

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if olookbook.HasPreScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= olookbook.StartScrollPage-1 %>');">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + olookbook.StartScrollPage to olookbook.StartScrollPage + olookbook.FScrollCount - 1 %>
			<% if (i > olookbook.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(olookbook.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:frmsubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if olookbook.HasNextScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= i %>');">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>

<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</table>

<%
set olookbook = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->