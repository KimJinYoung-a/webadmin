<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/admin/MenuCls.asp"-->
<%

dim i,j, strFTree, strColor, tmpMenuName
dim url, admSelPosit, admSelLevel
dim searchString

if session("ssAdminLsn")="1" then
	'관리자라면 레벨정보를 받을 수 있음
	admSelPosit = Request("admSelPosit")
	admSelLevel = Request("admSelLevel")
end if

searchString = replace(html2db(Request("searchString")), Chr(34), "")

if admSelPosit="" then admSelPosit=session("ssAdminPsn")
if admSelLevel="" then admSelLevel=session("ssAdminLsn")


'// ============================================================================
'// 메뉴
dim oMenuList

set oMenuList = new CMenuList

oMenuList.FRectPart_sn = admSelPosit
oMenuList.FRectLevel_sn = admSelLevel

'// 특정 권한 있으면 다른 부서의 메뉴 조회 가능
oMenuList.FRectHasAdminAuth = "N"

if (session("ssAdminLsn")="1") then
    if (Request("admSelPosit")="") then
		oMenuList.FRectHasAdminAuth = "Y"
    end if
elseif (iiisAdmin) then
	oMenuList.FRectHasAdminAuth = "Y"
else
    oMenuList.FRectUserID = ""
end if

oMenuList.FRectUserID = session("ssBctID")
oMenuList.FRectSearchString = searchString

oMenuList.GetLeftMenuListNew


'// ============================================================================
'// 즐겨찾기
dim oFavMenuList

set oFavMenuList = new CMenuList

oFavMenuList.FRectPart_sn = admSelPosit
oFavMenuList.FRectLevel_sn = admSelLevel

'// 특정 권한 있으면 다른 부서의 메뉴 조회 가능
oFavMenuList.FRectHasAdminAuth = "N"
if (session("ssAdminLsn")="1") then
    if (Request("admSelPosit")="") then
		oFavMenuList.FRectHasAdminAuth = "Y"
    end if
elseif (iiisAdmin) then
	oFavMenuList.FRectHasAdminAuth = "Y"
else
    oFavMenuList.FRectUserID = ""
end if

oFavMenuList.FRectUserID = session("ssBctID")
oFavMenuList.FRectIsFavorite = "Y"

oFavMenuList.GetLeftMenuListNew

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link rel="StyleSheet" href="/css/dtree.css" type="text/css" />
<style type="text/css">
<!--
	body {  font-size: 9pt}
	td {  font-size: 9pt}
	a {  text-decoration: none;color:#000000;}
-->
</style>
</head>
<body topmargin="0" leftmargin=0>
<script language='javascript'>
<!--
	function PopMenuHelp(menupos){
		var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'admin_PopMenuHelp_d','width=800, height=600, scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	function fnVPsubmit() {
		var frm = document.frmVP;
		if(!frm.admSelPosit.value) {
			alert("부서를 선택해주세요.");
			return false;
		}
		if(!frm.admSelLevel.value) {
			alert("등급을 선택해주세요.");
			return false;
		}
	}

	function fnPopEditFavorite() {
		var popwin = window.open("/admin/menu/popEditFavorite.asp","fnPopEditFavorite","width=700, height=400, scrollbars=yes,resizable=yes");
		popwin.focus();
	}
//-->
</script>
<script language="javascript" src="/js/dtree.js"></script>

<table width="100%" border="0" cellSpacing="0" cellPadding="0">
<%
	'관리자라면 등급을 지정할 수 있도록 표시(2010.08.10; 허진원)
	if session("ssAdminLsn")="1" then
%>
<form name="frmVP" method="GET" onSubmit="return fnVPsubmit();">
<tr>
	<td valign="top" style="padding:5px;" bgcolor="#F8F8F8">
		<b>메뉴선택 보기</b><br>
		<%=printPartOption("admSelPosit", admSelPosit)%><br>
		<%=printLevelOption("admSelLevel", admSelLevel)%>
		<input type="submit" value="변경" class="button">
	</td>
</tr>
</form>
<%	end if %>
<tr>
	<td valign="top">
		<img src="/images/icon_help.gif" width="50" height="20" onclick="PopMenuHelp('');" style="cursor:pointer">
	</td>
</tr>
<tr>
	<form name="frmSearch" method="get">
	<td align="left" height="35">
		&nbsp;
		&nbsp;
		<input type="text" class="text" name="searchString" size="12" value="<%= searchString %>">
		<input type="submit" class="button" value="검색">
	</td>
	</form>
</tr>
<tr>
	<td valign="top">
		<script type="text/javascript">
		var menuFavorite = new dTree("menuFavorite");

		menuFavorite.config.useCookies = false;

		menuFavorite.add(0,-1,"즐겨찾기 <a href='javascript:fnPopEditFavorite()' onfocus='this.blur();'><font color='blue'>[수정]</font></a>");

		<%
		for i=0 to oFavMenuList.FResultCount - 1
			url = oFavMenuList.FItemList(i).Fmenu_linkurl
			if IsNull(url) then
				url = ""
			end if

			if (url <> "") then
				if instr(url,"?")>0 then
					url=url & "&menupos=" & oFavMenuList.FItemList(i).Fmenu_id
				else
					url=url & "?menupos=" & oFavMenuList.FItemList(i).Fmenu_id
				end if
			end if

			strColor = oFavMenuList.FItemList(i).Fmenu_color
			tmpMenuName = oFavMenuList.FItemList(i).Fmenu_name
			if IsNull(strColor) then
				strColor = ""
			end if

			if (strColor <> "") then
				tmpMenuName = "<font color='" + CStr(strColor) + "'>" + CStr(tmpMenuName) + "</font>"
			end if


			%>menuFavorite.add(<%= oFavMenuList.FItemList(i).Fcid %>, <%= oFavMenuList.FItemList(i).Fpid %>, "<%= tmpMenuName %>", "<%= url %>"); <%
		next
		%>

		document.write(menuFavorite);
		</script>
		<br>
	</td>
</tr>
<tr>
	<td valign="top">
		<script type="text/javascript">
		var menuAll = new dTree("menuAll");

		menuAll.config.useCookies = false;

		menuAll.add(0,-1,"Admin");

		<%
		for i=0 to oMenuList.FResultCount - 1
			url = oMenuList.FItemList(i).Fmenu_linkurl
			if IsNull(url) then
				url = ""
			end if

			if (url <> "") then
				if instr(url,"?")>0 then
					url=url & "&menupos=" & oMenuList.FItemList(i).Fmenu_id
				else
					url=url & "?menupos=" & oMenuList.FItemList(i).Fmenu_id
				end if
			end if

			strColor = oMenuList.FItemList(i).Fmenu_color
			tmpMenuName = oMenuList.FItemList(i).Fmenu_name
			if IsNull(strColor) then
				strColor = ""
			end if

			if (strColor <> "") then
				tmpMenuName = "<font color='" + CStr(strColor) + "'>" + CStr(tmpMenuName) + "</font>"
			end if

			%>menuAll.add(<%= oMenuList.FItemList(i).Fcid %>, <%= oMenuList.FItemList(i).Fpid %>, "<%= tmpMenuName %>", "<%= url %>"); <%
		next
		%>

		document.write(menuAll);

		</script>
		<br>
	</td>
</tr>
</table>

</body>
</html>
<%
set oMenuList = Nothing
set oFavMenuList = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
