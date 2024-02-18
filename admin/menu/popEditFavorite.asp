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
<style type="text/css">
<!--
	body {  font-size: 9pt}
	td {  font-size: 9pt}
	a {  text-decoration: none;color:#000000;}
-->
</style>
</head>
<body topmargin="0" leftmargin=0>
<script language="javascript">
<!--

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

function fnAddFavorite(menuid, menuname) {
	var sel = document.frmAddFav.selectAddFav;

	for (var i = 0; i < sel.options.length; i++) {
		if (sel.options[i].value == menuid) {
			alert("�̹� ���õ� �޴��Դϴ�.");
			return;
		}
	}

	sel.options[sel.options.length] = new Option(menuname + "(" + menuid + ")", menuid);
}

function fnDelFavorite(menuid, menuname) {
	var sel = document.frmDelFav.selectDelFav;

	for (var i = 0; i < sel.options.length; i++) {
		if (sel.options[i].value == menuid) {
			alert("�̹� ���õ� �޴��Դϴ�.");
			return;
		}
	}

	sel.options[sel.options.length] = new Option(menuname + "(" + menuid + ")", menuid);
}

function fnRealAddFavorite() {
	if (confirm("�����Ͻðڽ��ϱ�?") != true) {
		return;
	}

	var frm = document.frmAddFav;
	var sel = document.frmAddFav.selectAddFav;

	frm.menu_id.value = "-1";
	for (var i = 0; i < sel.options.length; i++) {
		frm.menu_id.value = frm.menu_id.value + "," + sel.options[i].value;
	}

	if (frm.menu_id.value == "-1") {
		alert("���õ� �޴��� �����ϴ�.");
		return;
	}

	frm.submit();
}

function fnRealDelFavorite() {
	if (confirm("�����Ͻðڽ��ϱ�?") != true) {
		return;
	}

	var frm = document.frmDelFav;
	var sel = document.frmDelFav.selectDelFav;

	frm.menu_id.value = "-1";
	for (var i = 0; i < sel.options.length; i++) {
		frm.menu_id.value = frm.menu_id.value + "," + sel.options[i].value;
	}

	if (frm.menu_id.value == "-1") {
		alert("���õ� �޴��� �����ϴ�.");
		return;
	}

	frm.submit();
}

function fnCloseWin() {
	opener.focus();
	window.close();
}

//-->
</script>
<script language="javascript" src="/js/dtree.js"></script>

<table width="100%" border="0" cellSpacing="3" cellPadding="3">
<%
	'�����ڶ�� ����� ������ �� �ֵ��� ǥ��(2010.08.10; ������)
	if session("ssAdminLsn")="1" then
%>
<form name="frmVP" method="GET" onSubmit="return fnVPsubmit();">
<tr>
	<td valign="top" style="padding:5px;" bgcolor="#F8F8F8" colspan="2">
		<b>�޴����� ����[�����ں�]</b><br>
		<%=printPartOption("admSelPosit", admSelPosit)%><br>
		<%=printLevelOption("admSelLevel", admSelLevel)%>
		<input type="submit" value="����" class="button">
	</td>
</tr>
</form>
<%	end if %>
<tr>
	<td valign="top" width="50%" height="120">
		<script type="text/javascript">
		var menuFavoritePop = new dTree("menuFavoritePop");

		menuFavoritePop.config.target = "ifrAct";

		menuFavoritePop.add(0,-1,"���ã�⿡�� <font color='red'>����</font>�ϱ�");

		<%
		for i=0 to oFavMenuList.FResultCount - 1
			url = oFavMenuList.FItemList(i).Fmenu_linkurl
			if IsNull(url) then
				url = ""
			end if

			if (url <> "") then
				url = "popEditFavorite_process.asp?mode=tmpdelfavorite&menu_id=" + CStr(oFavMenuList.FItemList(i).Fmenu_id) + "&menu_name=" + CStr(oFavMenuList.FItemList(i).Fmenu_name)
			end if

			strColor = oFavMenuList.FItemList(i).Fmenu_color
			tmpMenuName = oFavMenuList.FItemList(i).Fmenu_name
			if IsNull(strColor) then
				strColor = ""
			end if

			if (strColor <> "") then
				tmpMenuName = "<font color='" + CStr(strColor) + "'>" + CStr(tmpMenuName) + "</font>"
			end if


			%>menuFavoritePop.add(<%= oFavMenuList.FItemList(i).Fcid %>, <%= oFavMenuList.FItemList(i).Fpid %>, "<%= tmpMenuName %>", "<%= url %>"); <%
		next
		%>

		document.write(menuFavoritePop);
		</script>
		<br>
	</td>
	<td valign="top">
		<form name="frmDelFav" onSubmit="return false;" action="popEditFavorite_process.asp">
		<input type="hidden" name="mode" value="realdelfavorite">
		<input type="hidden" name="menu_id" value="">
		<select class="select" id="selectDelFav" name="selectDelFav" size="3" multiple style="width: 250px;height: 60px;">
		</select>
		<br>
		<input type="button" class="button" value="����" onClick="fnRealDelFavorite()">
		&nbsp;
		<input type="button" class="button" value="���" onClick="fnCloseWin()">
		</form>
	</td>
</tr>
<tr>
	<td valign="top">
		<script type="text/javascript">
		var menuAllPop = new dTree("menuAllPop");

		menuAllPop.config.target = "ifrAct";

		menuAllPop.add(0,-1,"���ã�⿡ <font color='blue'>�߰�</font>�ϱ�");

		<%
		for i=0 to oMenuList.FResultCount - 1
			url = oMenuList.FItemList(i).Fmenu_linkurl
			if IsNull(url) then
				url = ""
			end if

			if (url <> "") then
				url = "popEditFavorite_process.asp?mode=tmpaddfavorite&menu_id=" + CStr(oMenuList.FItemList(i).Fmenu_id) + "&menu_name=" + CStr(oMenuList.FItemList(i).Fmenu_name)
			end if

			strColor = oMenuList.FItemList(i).Fmenu_color
			tmpMenuName = oMenuList.FItemList(i).Fmenu_name
			if IsNull(strColor) then
				strColor = ""
			end if

			if (strColor <> "") then
				tmpMenuName = "<font color='" + CStr(strColor) + "'>" + CStr(tmpMenuName) + "</font>"
			end if

			%>menuAllPop.add(<%= oMenuList.FItemList(i).Fcid %>, <%= oMenuList.FItemList(i).Fpid %>, "<%= tmpMenuName %>", "<%= url %>"); <%
		next
		%>

		document.write(menuAllPop);

		</script>
		<br>
	</td>
	<td valign="top">
		<form name="frmAddFav" onSubmit="return false;" action="popEditFavorite_process.asp">
		<input type="hidden" name="mode" value="realaddfavorite">
		<input type="hidden" name="menu_id" value="">
		<select class="select" id="selectAddFav" name="selectAddFav" size="5" multiple style="width: 250px;height: 120px;">
		</select>
		<br><br>
		<input type="button" class="button" value="�߰�" onClick="fnRealAddFavorite()">
		&nbsp;
		<input type="button" class="button" value="���" onClick="fnCloseWin()">
		</form>
	</td>
</tr>
</table>

<iframe src="" width="100" height="100" frameborder="0" name="ifrAct"></iframe>

</body>
</html>
<%
set oMenuList = Nothing
set oFavMenuList = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
