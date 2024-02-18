<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/CategoryMaster/displaycate/classes/displaycateCls.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Response.CharSet = "euc-kr"
	
	Dim cDisp, vDepth, vCateCode, vParentCateCode, vCateName, vCateName_E, vUseYN, vSortNo, vResultCount, vJaehuname, vIsNew
	vDepth			= RequestCheckvar(Request("depth"),10)
	vCateCode 		= RequestCheckvar(Request("catecode_s"),10)
	vParentCateCode	= RequestCheckvar(Request("parentcatecode"),10)
	
	SET cDisp = New cDispCate
	cDisp.FRectCateCode = vCateCode
	cDisp.GetDispCateDetail()
	
	vCateName 	= cDisp.FCateName
	vCateName_E	= cDisp.FCateName_E
	vJaehuname = cDisp.FJaehuname
	vUseYN		= cDisp.FUseYN
	vSortNo		= cDisp.FSortNo
	vIsNew		= cDisp.FIsNew
	vResultCount = cDisp.FResultCount
	SET cDisp = Nothing
	
	If vUseYN = "" Then vUseYN = "Y" End If
	If vIsNew = "" Then vIsNew = "x" End If
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
<% If (session("ssBctID") <> "cogusdk") Then %>
<input type="hidden" name="jaehuname" id="completedel" value="<%=vJaehuname%>">
<% End If %>
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
<!--
<tr>
	<td bgcolor="#F3F3FF" height="30">ī�װ���(����)</td>
	<td bgcolor="#FFFFFF"><input type="text" name="catename_e" style="width:250px;" value="<%=vCateName_E%>"> (�� ������ <u>Ư�����ڴ� ����</u>���ֽñ� �ٶ��ϴ�. Ư�� <u>��ǥ(,) Ȭ����ǥ(') �ֵ���ǥ(")</u>)</td>
</tr>
//-->
<tr>
	<td bgcolor="#F3F3FF" height="30">�������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="useyn" id="useyn_1" value="Y" <%=CHKIIF(vUseYN="Y","checked","")%> /><label for="useyn_1" style="cursor:pointer;">���</label>
		<input type="radio" name="useyn" id="useyn_2" value="N" <%=CHKIIF(vUseYN="N","checked","")%> /><label for="useyn_2" style="cursor:pointer;">������</label>
		&nbsp;�� ���� : <%=vCateCode%> <b>���� depth</b> ī�װ� <b>���</b>, ������ ������ <b>����</b>�˴ϴ�.
	</td>
</tr>
<input type="hidden" name="isnew" value="x">
<!--
<tr>
	<td bgcolor="#F3F3FF" height="30"><img src="http://fiximage.10x10.co.kr/web2013/common/ico_new2.gif" /> ������ �������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isnew" id="isnew_1" value="o" <%=CHKIIF(vIsNew="o","checked","")%> /><label for="isnew_1" style="cursor:pointer;">���</label>
		<input type="radio" name="isnew" id="isnew_2" value="x" <%=CHKIIF(vIsNew="x","checked","")%> /><label for="isnew_2" style="cursor:pointer;">������</label>
	</td>
</tr>
//-->
<% If vCateCode <> "" Then %>
<tr>
	<td bgcolor="#F3F3FF" height="30">�������</td>
	<td bgcolor="#FFFFFF">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td valign="top"><input type="button" value="��������" onClick="jsCateCompleteDel()"><td>
			<td valign="top">&nbsp;�� ���� : ���� ������ ������� <b>���� ����</b>(�����ȵ�). ī�װ��� ��ǰ <b>��� ����</b>(�����ȵ�). �귣�� ����ī�װ� <b>����</b>(�����ȵ�).<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>���� depth�� ī�װ�</b>�� �������� ��ǰ�� �ٸ� ī�װ��� <b>�̵�</b>�� �ϰų� ���� depth ī�װ� <b>����</b> �� �����ϼ���.
			</td>
		</tr>
		</table>
	</td>
</tr>
<% End If %>
<tr>
	<td bgcolor="#F3F3FF" height="30">���Ĺ�ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" name="sortno" style="width:70px;" value="<%=vSortNo%>"> (�� ���ڰ� �������� ��ܿ� ��Ÿ���ϴ�.)</td>
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
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->