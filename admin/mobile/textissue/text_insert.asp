<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterClass/main_TextIssueCls.asp"-->
<%
'###############################################
' Discription : �ؽ�Ʈ �̽�
' History : 2013.12.14 ����ȭ
'###############################################

	Dim idx, siteDiv

	idx = Request("idx")

	'// ���� ����
	dim oKeyword
	Set oKeyword = new CSearchKeyWord
	oKeyword.FRectIdx = idx

	if idx<>"" then
		oKeyword.GetSearchKeyWord
	end if
%>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language="javascript">
<!--
	// ������ ���� ����
	function goSubmit()
	{
		// Ű���� �Է¿��� �˻�
		if(!document.frm.keyword.value) {
			alert("���� Ű���带 �Է����ּ���.");
			document.frm.keyword.focus();
			return;
		}
		// ��ũ �Է¿��� �˻�
		if(!document.frm.linkinfo.value) {
			alert("Ű���� Ŭ���� �̵��� ��ũ�� �Է����ּ���.");
			document.frm.linkinfo.focus();
			return;
		}
		// ���� �Է¿��� �˻�
		if(!document.frm.sortNo.value) {
			alert("ǥ�� ������ �Է����ּ���.\n�� ������ �����̸� �������� ������ �����ϴ�.");
			document.frm.sortNo.focus();
			return;
		}

		<% if idx="" then %>
		if(confirm("�ۼ��Ͻ� ������ ����Ͻðڽ��ϱ�?")) {
			document.frm.mode.value="add";
			document.frm.action="dotextissue.asp";
			document.frm.submit();
		}
		<% else %>
		if(confirm("�����Ͻ� ������ �����Ͻðڽ��ϱ�?")) {
			document.frm.mode.value="modify";
			document.frm.action="dotextissue.asp";
			document.frm.submit();
		}
		<% end if %>
	}

	function putLinkText(key) {
		var frm = document.frm;
		switch(key) {
			case 'search':
				frm.linkinfo.value='/search/search_item.asp?rect=' + document.frm.keyword.value;
				break;
			case 'event':
				frm.linkinfo.value='/event/eventmain.asp?eventid=�̺�Ʈ��ȣ';
				break;
			case 'itemid':
				frm.linkinfo.value='/category/category_itemprd.asp?itemid=��ǰ�ڵ�';
				break;
			case 'category':
				frm.linkinfo.value='/category/category_list.asp?disp=ī�װ�';
				break;
			case 'brand':
				frm.linkinfo.value='/street/street_brand.asp?makerid=�귣����̵�';
				break;
		}
	}

	function fnChangeDiv(val) {
		if(val=="T") {
			document.getElementById("lyrExMCate").style.display="none";
		} else {
			document.getElementById("lyrExMCate").style.display="";
		}
	}
//-->
</script>
<!-- �� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="dotextissue.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="mode" value="">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<% if idx="" then %>
		<font color="red"><b>�ؽ�Ʈ�̽� ���</b></font>
		<% else %>
		<font color="red"><b>�ؽ�Ʈ�̽� ����</b></font>
		<% end if%>
	</td>
</tr>
<% if idx<>"" then %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">�Ϸù�ȣ</td>
	<td align="left"><input type="text" name="idx" value="<%=idx%>" readonly size="10" class="text_ro"></td>
</tr>
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">�ؽ�Ʈ�̽�</td>
	<td align="left"><input type="text" name="keyword" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).Ftextname%>" size="32" maxlength="32" class="text"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ũ</td>
	<td align="left">
		<table cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td colspan="2"><input type="text" name="linkinfo" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).Flinkinfo%>" size="80" maxlength="128" class="text"></td>
		<tr>
		<tr>
			<td valign="top"><font color="#707080">��)</font></td>
			<td valign="top">
				<font color="#707070">
				- <span style="cursor:pointer" onClick="putLinkText('search')">�˻���� ��ũ : /search/search_item.asp?rect=<font color="darkred">�˻���</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('event')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('itemid')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('category')">ī�װ� ��ũ : /category/category_list.asp?disp=<font color="darkred">ī�װ�</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('brand')">�귣����̵� ��ũ : /street/street_brand.asp?makerid=<font color="darkred">�귣����̵�</font></span>
				</font>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td align="left"><input id="prevDate" name="prevDate" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).Fenddate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
	<script type="text/javascript">
		var CAL_Start = new Calendar({
			inputField : "prevDate", trigger    : "prevDate_trigger",
			onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
	</script>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">ǥ�ü���</td>
	<td align="left"><input type="text" name="sortNo" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).FsortNo: else Response.Write "99" %>" size="3" class="text"></td></td>
</tr>
<tr>
	<td align="center" colspan="2" bgcolor="#FFFFFF">
		<input type="button" class="button" value="����" onClick="goSubmit()"> &nbsp;
		<input type="button" class="button" value="���" onClick="self.history.back()">
	</td>
</tr>
</form>
<!-- �� �� -->
</table>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
