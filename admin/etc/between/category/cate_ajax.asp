<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ��Ʈ��
' History : 2014.10.02 ������ ����
'			2015.08.12 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/betweenItemcls.asp"-->
<%
Response.CharSet = "euc-kr"

Dim cDisp, vDepth, vCateCode, vParentCateCode, vCateName, vCateName_E, vUseYN, vSortNo, vResultCount, vdispyn
vDepth			= Request("depth")
vCateCode 		= Request("catecode_s")
vParentCateCode	= Request("parentcatecode")

SET cDisp = New cDispCate
	cDisp.FRectCateCode = vCateCode
	cDisp.GetDispCateDetail()
	
	vCateName		= cDisp.FCateName
	vUseYN			= cDisp.FUseYN
	vSortNo			= cDisp.FSortNo
	vdispyn	= cDisp.fdispyn
	vResultCount	= cDisp.FResultCount
SET cDisp = Nothing

If vUseYN = "" Then vUseYN = "Y" End If
If vdispyn = "" Then vdispyn = "N" End If
If vSortNo = "" Then vSortNo = "99" End If
%>
<script>
$(function() {
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
});
</script>
<input type="hidden" name="parentcatecode" value="<%=vParentCateCode%>">
<input type="hidden" name="catecode" value="<%=vCateCode%>">
<input type="hidden" name="depth" value="<%=vDepth%>">
<input type="hidden" name="completedel" id="completedel" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr>
	<td bgcolor="#F3F3FF" width="70" height="30"></td>
	<td bgcolor="#FFFFFF" align="center"><b>ī�װ� <%=CHKIIF(vCateCode="","����","����")%></b></td>
</tr>
<% If vCateCode <> "" Then %>
<tr>
	<td bgcolor="#F3F3FF" height="30">ī�װ��ڵ�</td>
	<td bgcolor="#FFFFFF"><%=vCateCode%></td>
</tr>
<% End If %>
<tr>
	<td bgcolor="#F3F3FF" height="30">ī�װ���</td>
	<td bgcolor="#FFFFFF"><input type="text" name="catename" style="width:250px;" value="<%=vCateName%>"> (�� ������ <u>Ư�����ڴ� ����</u>���ֽñ� �ٶ��ϴ�. Ư�� <u>��ǥ(,) Ȭ����ǥ(') �ֵ���ǥ(")</u>)</td>
</tr>
<tr>
	<td bgcolor="#F3F3FF" height="30">�������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="useyn" id="useyn_1" value="Y" <%=CHKIIF(vUseYN="Y","checked","")%> /><label for="useyn_1">���</label>
		<input type="radio" name="useyn" id="useyn_2" value="N" <%=CHKIIF(vUseYN="N","checked","")%> /><label for="useyn_2">������</label>
	</td>
</tr>
<tr>
	<td bgcolor="#F3F3FF" height="30">���⿩��</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="dispyn" id="dispyn_1" value="Y" <%=CHKIIF(vdispyn="Y","checked","")%> /><label for="dispyn_1">Y</label>
		<input type="radio" name="dispyn" id="dispyn_2" value="N" <%=CHKIIF(vdispyn="N","checked","")%> /><label for="dispyn_2">N</label>
	</td>
</tr>
<% If vCateCode <> "" Then %>
<tr>
	<td bgcolor="#F3F3FF" height="30">�������</td>
	<td bgcolor="#FFFFFF">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td><input type="button" class="button" value="��������" onClick="jsCateCompleteDel()"><td>
			<td valign="top">&nbsp;�� ���� : ���� ������ ������� <b>���� ����</b>(�����ȵ�). ī�װ��� ��ǰ <b>��� ����</b>(�����ȵ�)</td>
		</tr>
		</table>
	</td>
</tr>
<% End If %>
<tr>
	<td bgcolor="#F3F3FF" height="30">���Ĺ�ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" name="sortno" style="width:70px;" value="<%=vSortNo%>"> (�� ���ڰ� �������� ������ ��Ÿ���ϴ�.)</td>
</tr>
<tr>
	<td id="lyrSbmBtn" bgcolor="#FFFFFF" colspan="2">
		<table width="100%" class="a">
		<tr>
			<td></td>
			<td align="right"><input type="button" value="��  ��" onClick="jsSaveDispCate()"></td>
		</tr>
		</table>
		<script>
			$("#lyrSbmBtn input").button();
		</script>
	</td>
</tr>
</table>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->