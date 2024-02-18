<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/HelpMenuCls.asp"-->
 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link rel="StyleSheet" href="/css/dtree.css" type="text/css" />
<link rel="stylesheet" href="/js/jqueryui/css/jquery-ui.css">
<% '<link rel="stylesheet" href="//code.jquery.com/ui/1.11.1/themes/smoothness/jquery-ui.css"> %>
<script language="JavaScript" src="/cscenter/js/jquery-1.8.3.js"></script>
<script language="JavaScript" src="/cscenter/js/jquery-ui-1.9.2.min.js"></script>
<script language="javascript" src="/js/dtree.js"></script>
<style type="text/css">
<!--
	body {  font-size: 9pt}
	td {  font-size: 9pt}
	a {  text-decoration: none}
-->
</style>
</head>
<body topmargin="0" >
<%
dim menupos, udiv 
dim i,j, strFTree, strColor, tmpMenuName
menupos = requestCheckvar(request("menupos"),10)
udiv = requestCheckvar(request("udiv"),10)

if udiv="" then udiv="9999"

'dim allMenuItem,i,j
'
'set allMenuItem = new CMenu
'allMenuItem.FrectUsingOnly="Y"
'allMenuItem.getMenuItems udiv

dim url
dim parentmenuid


dim oMenuList

set oMenuList = new CMenuList  
oMenuList.FRectUserDiv = udiv
oMenuList.GetLeftMenuListNew
%>
<script language='javascript'>
function ResearchMenu(comp){
	document.reloadfrm.submit();
}
</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<form name="reloadfrm" method=get >
<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        	도움말
				<% if C_ADMIN_AUTH then %>
				<select name="udiv" onChange="ResearchMenu(this)">
				<option value="9999" <% if udiv="9999" then response.write "selected" %> > 업체(9999)
				<option value="999" <% if udiv="999" then response.write "selected" %> > 제휴사(999)
				<option value="9" <% if udiv="9" then response.write "selected" %> > 관리자(9)
				<option value="2" <% if udiv="2" then response.write "selected" %> > SCM(9,7,5,4,2,1)
				<option value="501" <% if udiv="501" then response.write "selected" %> > 직영점(501)
				<option value="502" <% if udiv="502" then response.write "selected" %> > 가맹점(502)
				<option value="503" <% if udiv="503" then response.write "selected" %> > 기타매장(503)
				<option value="101" <% if udiv="101" then response.write "selected" %> > 오프샾(101)
				<option value="301" <% if udiv="301" then response.write "selected" %> > 컬리지(301)
				</select>
				<% end if %>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</form>

<!-- 표 상단바 끝-->



<tr>
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top">
	<DIV id="folderTree" STYLE="padding-top: 8px;"> 
	 <script type="text/javascript">
	
		var menuAllNameArray = new Array(<%= oMenuList.FResultCount %>);
 
		var menuAll = new dTree("menuAll");

		menuAll.config.useCookies = false;

		menuAll.add(0,-1,"Admin");

		<%
		for i=0 to oMenuList.FResultCount - 1
			url =  "/common/help/help.asp"
			if IsNull(url) then
				url = ""
			end if

			if (url <> "") then
				if (oMenuList.FItemList(i).Fmenu_useSslYN = "Y") and (application("Svr_Info") <> "Dev") then
					url = "https://webadmin.10x10.co.kr" + url
				end if

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
			if (url <> "") then
				%>menuAllNameArray[<%= i %>] = "<%= oMenuList.FItemList(i).Fmenu_name %>"; <%
			else
				%>menuAllNameArray[<%= i %>] = "XXX"; <%
			end if
		next
		%>
 
		document.write(menuAll);

		var menuAllNameArrayUniq = [];
		$.each(menuAllNameArray, function(i, el){
			if($.inArray(el, menuAllNameArrayUniq) === -1) menuAllNameArrayUniq.push(el);
		});
		</script>

	</DIV>
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>


<!-- 표 하단바 시작-->
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
 
</body>
</html>
<%
set oMenuList = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->