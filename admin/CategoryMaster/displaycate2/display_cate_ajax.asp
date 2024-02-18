<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/categorymaster/displaycate2/classes/displaycateCls.asp"-->

<%
	Response.CharSet = "euc-kr"

	Dim cDisp, vDepth, vCateCode, vParentCateCode, vCateName, vCateName_E, vUseYN, vSortNo, vResultCount, vJaehuname, vIsNew, vCateKeywords, vSafetyInfoType, vDownCateCount, vSearchKeywords
	vDepth			= Request("depth")
	vCateCode 		= Request("catecode_s")
	vParentCateCode	= Request("parentcatecode")

	SET cDisp = New cDispCate
	cDisp.FRectCateCode = vCateCode
	cDisp.GetDispCateDetail()

	vCateName 		= cDisp.FCateName
	vCateName_E		= cDisp.FCateName_E
	vJaehuname 		= cDisp.FJaehuname
	vUseYN			= cDisp.FUseYN
	vSortNo			= cDisp.FSortNo
	vIsNew			= cDisp.FIsNew
	vCateKeywords	= cDisp.FCateKeywords
	vSafetyInfoType = cDisp.FSafetyInfoType
	vDownCateCount = cDisp.FDownCateCount
	vResultCount 	= cDisp.FResultCount
	vSearchKeywords = cDisp.FsearchKeywords
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
	<td bgcolor="#F3F3FF" width="120" height="30"></td>
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
	<td bgcolor="#F3F3FF" height="30">ī�װ���(����)</td>
	<td bgcolor="#FFFFFF"><input type="text" name="catename_e" style="width:250px;" value="<%=vCateName_E%>"> (�� ������ <u>Ư�����ڴ� ����</u>���ֽñ� �ٶ��ϴ�. Ư�� <u>��ǥ(,) Ȭ����ǥ(') �ֵ���ǥ(")</u>)</td>
</tr>
<% If (session("ssBctID") = "cogusdk") Then %>
<tr>
	<td bgcolor="#F3F3FF" height="30">ī�װ���(����)</td>
	<td bgcolor="#FFFFFF"><input type="text" name="jaehuname" style="width:250px;" value="<%=vJaehuname%>"> (�� ������ <u>Ư�����ڴ� ����</u>���ֽñ� �ٶ��ϴ�. Ư�� <u>��ǥ(,) Ȭ����ǥ(') �ֵ���ǥ(")</u>)</td>
</tr>
<% End If %>
<tr>
	<td bgcolor="#F3F3FF" height="30">�������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="useyn" id="useyn_1" value="Y" <%=CHKIIF(vUseYN="Y","checked","")%> /><label for="useyn_1" style="cursor:pointer;">���</label>
		<input type="radio" name="useyn" id="useyn_2" value="N" <%=CHKIIF(vUseYN="N","checked","")%> /><label for="useyn_2" style="cursor:pointer;">������</label>
		&nbsp;�� ���� : <%=vCateCode%> <b>���� depth</b> ī�װ� <b>���</b>, ������ ������ <b>����</b>�˴ϴ�.
	</td>
</tr>
<tr>
	<td bgcolor="#F3F3FF" height="30"><img src="http://fiximage.10x10.co.kr/web2013/common/ico_new2.gif" /> ������ �������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isnew" id="isnew_1" value="o" <%=CHKIIF(vIsNew="o","checked","")%> /><label for="isnew_1" style="cursor:pointer;">���</label>
		<input type="radio" name="isnew" id="isnew_2" value="x" <%=CHKIIF(vIsNew="x","checked","")%> /><label for="isnew_2" style="cursor:pointer;">������</label>
	</td>
</tr>
<tr>
	<td bgcolor="#F3F3FF" height="30">�˻�Ű����</td>
	<td bgcolor="#FFFFFF"><input type="text" name="searchKeywords" style="width:500px;" value="<%=vSearchKeywords%>"> (�޸��α��� ex: Ŀ��,Ƽ����,����)</td>
</tr>
<tr>
	<td bgcolor="#F3F3FF" height="30">�ܺ� ���� Ű����</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="keywords" style="width:700px;" value="<%= vCateKeywords %>"><br><br>
		&nbsp;&nbsp; * Ű���� ��Ͻ� �˻�����(���� ��) �˻������ ���� ���� �����<br><br>
		&nbsp;&nbsp; �� ��� ��:<br>
		&nbsp;&nbsp; ����ũ�� �Ŀ�ġ =&gt; ����ũ�� �Ŀ�ġ/ȭ��ǰ �Ŀ�ġ/����� �Ŀ�ġ<br>
		&nbsp;&nbsp; ������ =&gt; ������/���� ������/���� ������<br>
		&nbsp;&nbsp; ������ =&gt; ������/�ǳ�ȭ/���� ������/���� �ǳ�ȭ/�繫�� ������/�л� ������<br>
		&nbsp;
	</td>
</tr>
<% If vDownCateCount = 0 Then %>
<tr>
	<td bgcolor="#F3F3FF" height="30">�������� ���� ����</td>
	<td bgcolor="#FFFFFF">
		<select name="safetyinfotype">
			<option value="" <%=CHKIIF(vSafetyInfoType="","selected","")%>>���þ���</option>
			<option value="choice" <%=CHKIIF(vSafetyInfoType="choice","selected","")%>>���� �Է� ����</option>
			<option value="necessary" <%=CHKIIF(vSafetyInfoType="necessary","selected","")%>>�ʼ� �Է� ����</option>
		</select>
	</td>
</tr>
<% End If %>
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