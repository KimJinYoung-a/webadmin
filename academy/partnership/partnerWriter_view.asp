<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2016.07.22 ������ ����
'	Description : �۰� ��û ����Ʈ
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/partner_writercls.asp"-->
<%
Dim idx
Dim oWriter, i
idx = RequestCheckvar(request("idx"),10)

Set oWriter = new CWriter
	oWriter.FRectIdx = idx
	oWriter.getWriterViewOneitem
%>
<script language='javascript'>
	function chk_form(frm)
	{
		if(!frm.confirmMemo.value)
		{
			alert("�亯 ������ �ۼ����ֽʽÿ�.");
			frm.confirmMemo.focus();
			return false;
		}
		return true;
	}

	function GotoWriterDel(){
		if (confirm('���� �Ͻðڽ��ϱ�?')){
			document.frm_write.mode.value="DelWriter";
			document.frm_write.submit();
		}
	}

	function NewWindow(v){
	  window.open("http://www.thefingers.co.kr/myfingers/showimage.asp?img=" + v, "imageView", "left=10px,top=10px, width=560,height=400,status=no,resizable=yes,scrollbars=yes");
	}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doPartnerLecture.asp">
<input type="hidden" name="mode" value="AnsWriter">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">����</td>
	<td bgcolor="#FFFFFF" align="LEFT">
	<%
		Select Case oWriter.FOneItem.FGubun
			Case "1"	response.write "����"
			Case "2"	response.write "����"
			Case "3"	response.write "���"
		End Select
	%>
	</td>
</tr>
<tr align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">�۰���</td>
	<td bgcolor="#FFFFFF" align="LEFT"><%= oWriter.FOneItem.FWritername %></td>
</tr>
<tr align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">��ǰ�о�</td>
	<td bgcolor="#FFFFFF" align="LEFT"><%= oWriter.FOneItem.FBunya %></td>
</tr>
<tr align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">�ּ�</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		[ <%= oWriter.FOneItem.FZipcode %> ] <%= oWriter.FOneItem.FAddress1 %>&nbsp;<%= oWriter.FOneItem.FAddress2 %>
	</td>
</tr>
<tr align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">�޴�����ȣ</td>
	<td bgcolor="#FFFFFF" align="LEFT"><%= oWriter.FOneItem.FUsercell %></td>
</tr>
<tr align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">��ȭ��ȣ</td>
	<td bgcolor="#FFFFFF" align="LEFT"><%= oWriter.FOneItem.FUserphone %></td>
</tr>
<tr align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">�̸���</td>
	<td bgcolor="#FFFFFF" align="LEFT"><%= oWriter.FOneItem.FUsermail %></td>
</tr>
<tr align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">Ȩ������</td>
	<td bgcolor="#FFFFFF" align="LEFT">
	<%
		If oWriter.FOneItem.FHomepage<>"" Then
			Response.Write "<a href='"& oWriter.FOneItem.FHomepage & "' target='_blank'>" & oWriter.FOneItem.FHomepage & "</a>"
		End If
	%>
	</td>
</tr>
<tr align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">��ǰ�Ұ�</td>
	<td bgcolor="#FFFFFF" align="LEFT"><%= nl2br(oWriter.FOneItem.FIntroduce) %></td>
</tr>
<tr align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">��Ÿ</td>
	<td bgcolor="#FFFFFF" align="LEFT"><%= nl2br(oWriter.FOneItem.FEtc) %></td>
</tr>
<tr align="center" height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td width="200">÷������</td>
	<td bgcolor="#FFFFFF" align="LEFT">
	<%
		if oWriter.FOneItem.FWritefile<>"" then
			'���������� ���� ���� �ɼ� �߰�
			Select Case getFileExtention(oWriter.FOneItem.FWritefile)
				Case "jpg", "gif", "png"
					Response.Write "<span onClick=""NewWindow('" & imgFingers & oWriter.upfolder & "writer/" & oWriter.FOneItem.FWritefile & "')"" style='cursor:pointer'>" & oWriter.FOneItem.FWritefile & "</span>"
				Case Else
					Response.Write "<a href='" & imgFingers & "/linkweb/download.asp?filepath=" & Server.URLencode(oWriter.upfolder & "writer/") & "&filename=" & Server.URLencode(oWriter.FOneItem.FWritefile) & "'>" & oWriter.FOneItem.FWritefile & "</a>"
			end Select
		end if
	%>
	</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr height="30" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" width="200"><font color="darkred">* </font>�亯 ����</td>
	<td bgcolor="#FFFFFF">
		<textarea name="confirmMemo" rows="10" cols="80" class="textarea"><%= oWriter.FOneItem.FConfirmMemo %></textarea><br>
		�� �亯 ������ ����� ���� ���Դϴ�. ������ ���� ���� �����Ƿ� ��������� ������ֽʽÿ�.
	</td>
</tr>
<tr>
	<td colspan="4" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_reply.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoWriterDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="location.href='/academy/partnership/partnerWriter_list.asp?menupos=<%=menupos%>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->