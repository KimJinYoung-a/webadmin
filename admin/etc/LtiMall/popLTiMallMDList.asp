<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/lotteiMallcls.asp"-->
<%
Dim oiMall, i, page
page = request("page")
If page = "" Then page = 1

'// ��� ����
Set oiMall = new CLotteiMall
	oiMall.FPageSize = 10
	oiMall.FCurrPage = page
	oiMall.getLotte_MDList
%>
<script language="javascript">
<!--
// ���MD����
function refreshMDList() {
	if(confirm("���MD�� �Ե����̸� �������� �����޾� �����Ͻðڽ��ϱ�?\n\n�� ��Ż��¿����� �ټ� �ð��� �ɸ� �� �ֽ��ϴ�.")) {
		document.getElementById("btnRefresh").disabled=true;
		xLink.location.href="actLotteiMallMDList.asp";
	}
}

// ������ �̵�
function goPage(pg) {
	self.location.href="?page="+pg;
}

// ���MD ��ǰ�� ���� �˾�
function popMDGroupList(mdcd) {
	var pMD = window.open("popLotteiMallMDCateGroup.asp?mdcd="+mdcd,"popMDGroup","width=500,height=500,scrollbars=yes,resizable=yes");
	pMD.focus();
}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>�Ե����̸� ���MD ���</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- �׼� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;"><input id="btnRefresh" type="button" class="button" value="���MD����" onclick="refreshMDList()"></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �˻���� : <strong><%=oiMall.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>MD�ڵ�</td>
	<td>MD�̸�</td>
	<td>��������</td>
	<td>���������</td>
	<td>���μ�����</td>
	<td>��ǰ��</td>
</tr>
<% If oiMall.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="6" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to oiMall.FresultCount - 1
%>
<tr align="center" height="25" bgcolor="<%=chkIIF(oiMall.FItemList(i).FisUsing="Y","#FFFFFF","#DDDDDD")%>">
	<td><%= oiMall.FItemList(i).FMDCode %></td>
	<td><%= oiMall.FItemList(i).FMDName %></td>
	<td><%= oiMall.FItemList(i).FSellFeeType %></td>
	<td><%= oiMall.FItemList(i).FNormalSellFee %></td>
	<td><%= oiMall.FItemList(i).FEventSellFee %></td>
	<td><input type="button" class="button" value="��ǰ�� ����" onClick="popMDGroupList('<%= oiMall.FItemList(i).FMDCode %>')"></td>
</tr>
<%
		Next
	End If
%>
</table>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% If oiMall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oiMall.StartScrollPage - 1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>

		<% For i=0 + oiMall.StartScrollPage to oiMall.FScrollCount + oiMall.StartScrollPage - 1 %>
			<% If i > oiMall.FTotalpage Then Exit For %>
			<% If CStr(page) = CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>

		<% If oiMall.HasNextScroll Then %>
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
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="100"></iframe>
</p>
<% Set oiMall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->