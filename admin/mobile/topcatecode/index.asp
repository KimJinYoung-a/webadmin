<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/submenu/inc_subhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/topcatecodeCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : gnb 카테코드 저장
' History :2015-09-14 이종화
'###############################################

	Dim isusing , gcode
	dim page 
	Dim i
	dim subcodeList
	Dim sDt , modiTime

	page = request("page")
	gcode = request("gcode")
	isusing = RequestCheckVar(request("isusing"),1)

	if page="" then page=1
	If isusing = "" Then isusing ="Y"

	set subcodeList = new GNBsubcode
	subcodeList.FPageSize		= 20
	subcodeList.FCurrPage		= page
	subcodeList.Fisusing			= isusing
	subcodeList.FRectgnbcode	= gcode
	subcodeList.GetSubCodeList()

%>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
<!--
function popCateCodeManage(){
	var popwin = window.open('/admin/mobile/topcatecode/popcateinsert.asp','popcatecode','width=690,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popSubCodeManage(){
	var popwin = window.open('/admin/mobile/topcatecode/cate_insert.asp?menupos=<%=menupos%>','popcatecode','width=690,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function jsmodify(v){
	var popwin = window.open('/admin/mobile/topcatecode/cate_insert.asp?menupos=<%=menupos%>&idx='+v,'popcatecode','width=690,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function jssearch(){
	document.frm.submit();
}
-->
</script>
<!-- 검색 시작 -->
<table width="800" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
		<td align="left">
			<div >
			* 사용여부 :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			* GNB 메뉴 : 
			<% Call drawSelectBoxGNB("gcode" , gcode) %>
			</div>
		</td>
		<td width="150" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onclick="jssearch();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" class="button" value="GNB관리" onClick="popCateCodeManage();">
		</td>
	</tr>
</form>	
</table>
<!-- 검색 끝 -->

<table width="800" align="left" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
    <td align="right">
		<!-- 신규등록 -->
    	<a href="" onclick="popSubCodeManage();return false;"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  리스트 -->
<table width="800" align="left" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		총 등록수 : <b><%=subcodeList.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=subcodeList.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
    <td width="18%">GNB이름</td>
    <td width="18%">전시카테고리이름</td>
    <td width="10%">사용여부</td>
</tr>
<% 
	for i=0 to subcodeList.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(subcodeList.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=subcodeList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=subcodeList.FItemList(i).Fidx%></td>
    <td onclick="jsmodify('<%=subcodeList.FItemList(i).Fidx%>');" style="cursor:pointer;">[<%=subcodeList.FItemList(i).Fgnbcode%>]<%=subcodeList.FItemList(i).Fgnbname%></td>
    <td onclick="jsmodify('<%=subcodeList.FItemList(i).Fidx%>');" style="cursor:pointer;">[<%=subcodeList.FItemList(i).Fdispcode%>]<%=subcodeList.FItemList(i).Fdispname%></td>
    <td><%=chkiif(subcodeList.FItemList(i).Fisusing="N","사용안함","사용함")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="800" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if subcodeList.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= subcodeList.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + subcodeList.StartScrollPage to subcodeList.StartScrollPage + subcodeList.FScrollCount - 1 %>
				<% if (i > subcodeList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(subcodeList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if subcodeList.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set subcodeList = Nothing
%>
<!-- 검색 끝 -->
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->