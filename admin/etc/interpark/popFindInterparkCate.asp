<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/interpark/interparkcls.asp"-->
<%
Dim oInterpark, i, page, srcKwd
page		= request("page")
srcKwd		= request("srcKwd")

If page = ""	Then page = 1
'// ��� ����
Set oInterpark = new CInterpark
	oInterpark.FPageSize = 5000
	oInterpark.FCurrPage = page
	oInterpark.FRectSearchName = srcKwd
	oInterpark.getInterparkCateList
%>
<script language="javascript">
<!--
	// ������ �̵�
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}
	// ��ǰ�з� ����
	function fnSelDispCate(catekey, dispNm) {
		opener.document.frmAct.catekey.value = catekey;
		opener.document.getElementById("BrRow").style.display="";
		opener.document.getElementById("selBr").innerHTML= dispNm;
		self.close();
	}
//-->
</script>
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
</form>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>Interpark ��ǰ�з� �˻�</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<p>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �˻���� : <strong><%=oInterpark.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>���ù�ȣ</td>
	<td>Interpark ���ø�</td>
	<td>�׸��</td>
	<td>��Ȱ</td>
	<td>����</td>
	<td>���</td>
</tr>
<% If oInterpark.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to oInterpark.FresultCount - 1
%>
<tr align="center" height="25" onClick="fnSelDispCate('<%= oInterpark.FItemList(i).FCateKey %>', '<%= oInterpark.FItemList(i).FDispNm %>')" style="cursor:pointer" title="��ǰ�з� ����" bgcolor="#FFFFFF">
	<td><%= oInterpark.FItemList(i).FCateKey %></td>
	<td><%= oInterpark.FItemList(i).FDispNm %></td>
	<td><%= oInterpark.FItemList(i).FInfoGroupNm %></td>
	<td><%= oInterpark.FItemList(i).FIndustrial %></td>
	<td><%= oInterpark.FItemList(i).FElectric %></td>
	<td><%= oInterpark.FItemList(i).FChild %></td>
</tr>
<%
		Next
	End If
%>
</table>
</form>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% If oInterpark.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oInterpark.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>

		<% For i = 0 + oInterpark.StartScrollPage to oInterpark.FScrollCount + oInterpark.StartScrollPage - 1 %>
			<% If i>oInterpark.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<foNt color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% next %>

		<% If oInterpark.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% Else %>
			[next]
		<% End If %>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% Set oInterpark = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
