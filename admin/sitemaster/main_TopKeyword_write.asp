<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : main_TopKeyword_write.asp
' Discription : ���� žŰ���� ���/����
' History : 2008.04.18 ������ ����
'           2022.07.01 �ѿ�� ����(isms���������, �ҽ�ǥ��ȭ)
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterClass/main_TopKeywrdCls.asp"-->
<%
	Dim idx, siteDiv
	idx = Request("idx")
	siteDiv = Request("siteDiv")

	'// ���� ����
	dim oKeyword
	Set oKeyword = new CSearchKeyWord
	oKeyword.FRectIdx = idx

	if idx<>"" then
		oKeyword.GetSearchKeyWord
		siteDiv = oKeyword.FitemList(0).FsiteDiv
	end if
%>
<script type='text/javascript'>
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
		// ����� ��ũ �˻� (ī�װ� �Һз� ����˻�)
		if(document.frm.siteDiv.value!="T"&&document.frm.linkinfo.value.indexOf("category_list")>0&&document.frm.linkinfo.value.indexOf("cds")>0) {
			alert("����� ī�װ����� �Һз��� ���� �� ����ϴ�.\n����� �������� �Һз� ��ũ�� Ȯ���� �ּ���.");
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
			document.frm.action="doMainTopKeyword.asp";
			document.frm.submit();
		}
		<% else %>
		if(confirm("�����Ͻ� ������ �����Ͻðڽ��ϱ�?")) {
			document.frm.mode.value="modify";
			document.frm.action="doMainTopKeyword.asp";
			document.frm.submit();
		}
		<% end if %>
	}

	function putLinkText(key) {
		var frm = document.frm;
		switch(key) {
			case 'search':
				frm.linkinfo.value='/search/search_result.asp?rect=' + document.frm.keyword.value;
				break;
			case 'cate':
				if(frm.siteDiv.value=="M"||frm.siteDiv.value=="E") {
					frm.linkinfo.value='/category/category_list.asp?cdl=���ڵ�&cdm=���ڵ�';
				} else {
					frm.linkinfo.value='/shopping/category_list.asp?cdl=���ڵ�&cdm=���ڵ�&cds=���ڵ�';
				}
				break;
			case 'cateM':
				frm.linkinfo.value='/category/category_itemList.asp?cdl=���ڵ�&cdm=���ڵ�&cds=���ڵ�';
				break;
			case 'event':
				frm.linkinfo.value='/event/eventmain.asp?eventid=�̺�Ʈ��ȣ';
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
<form name="frm" method="post" action="doMainTopKeyword.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="mode" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<% if idx="" then %>
		<font color="red"><b>žŰ���� ���</b></font>
		<% else %>
		<font color="red"><b>žŰ���� ����</b></font>
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
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">���뱸��</td>
	<td align="left">
		<select class="select" name="siteDiv" onchange="fnChangeDiv(this.value)">
			<option value="T" <%=chkIIF(siteDiv="T","selected","")%>>PC��</option>
			<option value="M" <%=chkIIF(siteDiv="M","selected","")%>>�����:�˻���</option>
			<option value="E" <%=chkIIF(siteDiv="E","selected","")%>>�����:�̺�Ʈ</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">Ű����</td>
	<td align="left">
		<input type="text" name="keyword" value="<% if idx<>"" then Response.Write ReplaceBracket(oKeyword.FitemList(0).FKeyword) %>" size="32" maxlength="32" class="text">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ũ</td>
	<td align="left">
		<table cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td colspan="2"><input type="text" name="linkinfo" value="<% if idx<>"" then Response.Write ReplaceBracket(oKeyword.FitemList(0).Flinkinfo) %>" size="80" maxlength="128" class="text"></td>
		<tr>
		<tr>
			<td valign="top"><font color="#707080">��)</font></td>
			<td valign="top">
				<font color="#707070">
				- <span style="cursor:pointer" onClick="putLinkText('search')">�˻���� ��ũ : /search/search_result.asp?rect=<font color="darkred">�˻���</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('cate')">ī�װ� ��ũ : /shopping/category_list.asp?cdl=<font color="darkred">���ڵ�</font>&cdm=<font color="darkred">���ڵ�</font>&cds=<font color="darkred">���ڵ�</font></span><br>
				<span id="lyrExMCate" style="<%=chkIIF(siteDiv="T","display:none;","")%>">- <span style="cursor:pointer;" onClick="putLinkText('cateM')">����� ī�װ� �Һз� : /shopping/category_itemList.asp?cdl=<font color="darkred">���ڵ�</font>&cdm=<font color="darkred">���ڵ�</font>&cds=<font color="darkred">���ڵ�</font></span><br></span>
				- <span style="cursor:pointer" onClick="putLinkText('event')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span>
				</font>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">ǥ�ü���</td>
	<td align="left"><input type="text" name="sortNo" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).FsortNo: else Response.Write "0" %>" size="3" class="text"></td></td>
</tr>
<tr>
	<td align="center" colspan="2" bgcolor="#FFFFFF">
		<input type="button" class="button" value="����" onClick="goSubmit()"> &nbsp;
		<input type="button" class="button" value="���" onClick="self.history.back()">
	</td>
</tr>
<!-- �� �� -->
</table>
</form>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
