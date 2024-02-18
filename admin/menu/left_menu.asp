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
	'�����ڶ�� ���������� ���� �� ����
	admSelPosit = Request("admSelPosit")
	admSelLevel = Request("admSelLevel")
end if

searchString = replace(html2db(Request("searchString")), Chr(34), "")

if admSelPosit="" then admSelPosit=session("ssAdminPsn")
if admSelLevel="" then admSelLevel=session("ssAdminLsn")


'// ============================================================================
'// �޴�
dim oMenuList

set oMenuList = new CMenuList

oMenuList.FRectPart_sn = admSelPosit
oMenuList.FRectLevel_sn = admSelLevel

'// Ư�� ���� ������ �ٸ� �μ��� �޴� ��ȸ ����
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
'// ���ã��
dim oFavMenuList

set oFavMenuList = new CMenuList

oFavMenuList.FRectPart_sn = admSelPosit
oFavMenuList.FRectLevel_sn = admSelLevel

'// Ư�� ���� ������ �ٸ� �μ��� �޴� ��ȸ ����
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
<link rel="stylesheet" href="/js/jqueryui/css/jquery-ui.css">
<% '<link rel="stylesheet" href="//code.jquery.com/ui/1.11.1/themes/smoothness/jquery-ui.css"> %>
<script language="JavaScript" src="/cscenter/js/jquery-1.8.3.js"></script>
<script language="JavaScript" src="/cscenter/js/jquery-ui-1.9.2.min.js"></script>
<script language='javascript'>

/*
$( document ).ready(function() {
	$( "#searchInputBox" ).autocomplete({
		source: menuAllNameArrayUniq
	});

	$( "#searchInputBox" ).autocomplete({
		select: function (a, b) {
			var txt = b.item.value.replace(/[[].+]/i, "");

			$(this).val(txt);
			$("#frmSearch").submit();
		}
	});
});
*/

</script>
<style type="text/css">
<!--
	body {  font-size: 9pt}
	td {  font-size: 9pt}
	a {  text-decoration: none;color:#000000;}

	.ui-autocomplete { height: 250px; width: 100%; overflow-y: auto; overflow-x: hidden; }
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
			alert("�μ��� �������ּ���.");
			return false;
		}
		if(!frm.admSelLevel.value) {
			alert("����� �������ּ���.");
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
	'�����ڶ�� ����� ������ �� �ֵ��� ǥ��(2010.08.10; ������)
	if session("ssAdminLsn")="1" then
%>
<form name="frmVP" method="GET" onSubmit="return fnVPsubmit();">
<tr>
	<td valign="top" style="padding:5px;" bgcolor="#F8F8F8">
		<b>�޴����� ����</b><br>
		<%=printPartOption("admSelPosit", admSelPosit)%><br>
		<%=printLevelOption("admSelLevel", admSelLevel)%>
		<input type="submit" value="����" class="button">
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
	<form id="frmSearch" name="frmSearch" method="get">
	<td align="left" height="35">
		&nbsp;
		&nbsp;
		<input type="text" class="text" name="searchString" size="12" value="<%= searchString %>" id="searchInputBox">
		<input type="submit" class="button" value="�˻�">
	</td>
	</form>
</tr>
<tr>
	<td valign="top">
		<script type="text/javascript">
		var menuFavorite = new dTree("menuFavorite");

		menuFavorite.config.useCookies = false;

		menuFavorite.add(0,-1,"���ã�� <a href='javascript:fnPopEditFavorite()' onfocus='this.blur();'><font color='blue'>[����]</font></a>");

		<%
		for i=0 to oFavMenuList.FResultCount - 1
			url = oFavMenuList.FItemList(i).Fmenu_linkurl
			if IsNull(url) then
				url = ""
			end if

			if (url <> "") then
				if (oFavMenuList.FItemList(i).Fmenu_useSslYN = "Y") then
					''url = "https://webadmin.10x10.co.kr" + url
					url = getSCMSSLURL + url
				end if

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
		var menuAllNameArray = new Array(<%= oMenuList.FResultCount %>);

		var menuAll = new dTree("menuAll");

		menuAll.config.useCookies = false;

		menuAll.add(0,-1,"<a href='/admin/scmmain.asp' target='contents'>Admin</a>");

		<%
		for i=0 to oMenuList.FResultCount - 1
			url = oMenuList.FItemList(i).Fmenu_linkurl
			if IsNull(url) then
				url = ""
			end if

			if (url <> "") then
				if (oMenuList.FItemList(i).Fmenu_useSslYN = "Y") then
					''url = "https://webadmin.10x10.co.kr" + url
					url = getSCMSSLURL + url
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

			%>
			<% If instr(url,"http://")>0 Then %>
				menuAll.add(<%= oMenuList.FItemList(i).Fcid %>, <%= oMenuList.FItemList(i).Fpid %>, "<%= tmpMenuName %>", "<%= url %>", "", "_blank"); 
			<% Else %>
				menuAll.add(<%= oMenuList.FItemList(i).Fcid %>, <%= oMenuList.FItemList(i).Fpid %>, "<%= tmpMenuName %>", "<%= url %>"); 
			<% End If %>
			<%
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
