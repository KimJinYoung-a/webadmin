<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/offshop_noticecls.asp" -->
<%

dim i, j

'==============================================================================
'��������
dim boardnotice
set boardnotice = New CNoticeDetail

boardnotice.read(request("idx"))

%>
<script>
function SubmitForm()
{
        if (document.f.gubun.value == "") {
                alert("�����ֱ� ������ �����ϼ���.");
                return;
        }
        if (document.f.title.value == "") {
                alert("������ �Է��ϼ���.");
                return;
        }
        if (document.f.contents.value == "") {
                alert("������ �Է��ϼ���.");
                return;
        }

        document.f.submit();
}
function SubmitDelete()
{
        if (confirm("�����Ͻðڽ��ϱ�?") == true) {
                document.f.mode.value = "delete";
                document.f.submit();
        }
}
</script>
<table border="1" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="#808080" class="a">
<form method="post" name="f" action="offshop_notice_act.asp" onsubmit="return false" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="idx" value="<%= request("idx") %>">
<input type="hidden" name="userid" value="<%=session("ssBctId")%>">
<input type="hidden" name="username" value="<%=session("ssBctCname")%>">
<input type="hidden" name="mode" value="modify">
<tr>
	<td align="center">�����ֱ� ����</td>
	<td>
		<select name="gubun">
			<option value="" <% if boardnotice.Fgubun = "" then response.write "selected" %>>����</option>
			<option value="00" <% if boardnotice.Fgubun = "00" then response.write "selected" %>>��ü</option>
			<option value="01" <% if boardnotice.Fgubun = "01" then response.write "selected" %>>1F Shop</option>
			<option value="02" <% if boardnotice.Fgubun = "02" then response.write "selected" %>>3F Zoom</option>
			<option value="03" <% if boardnotice.Fgubun = "03" then response.write "selected" %>>3F College</option>
			<option value="04" <% if boardnotice.Fgubun = "04" then response.write "selected" %>>�¶��λ����</option>
			<option value="50" <% if boardnotice.Fgubun = "50" then response.write "selected" %>>����-��ü</option>
			<option value="51" <% if boardnotice.Fgubun = "51" then response.write "selected" %>>����-����</option>
			<option value="52" <% if boardnotice.Fgubun = "52" then response.write "selected" %>>����-������</option>
			<option value="53" <% if boardnotice.Fgubun = "53" then response.write "selected" %>>����-������</option>
		</select>
	</td>
</tr>
<tr>
	<td align="center">����</td>
	<td><input type="text" name="title" size="50" value="<%= boardnotice.Ftitle %>"></td>
</tr>
<tr>
	<td align="center">����</td>
	<td><textarea name="contents" cols="80" rows="18"><%= db2html(boardnotice.Fcontents) %></textarea></td>
</tr>
<tr>
	<td align="center">����</td>
	<td><input type="file" name="file" size="30"></td>
</tr>
<% if boardnotice.Ffile <> "" then %>
<tr>
	<td align="center">���ϻ���</td>
	<td><input type="checkbox" name="dl_file"><%= boardnotice.Ffile %></td>
</tr>
<% end if %>
<tr><td colspan="2" align="right"><input type="button" value=" ���� " onclick="SubmitForm()">
<input type="button" value=" ���� " onclick="SubmitDelete()">&nbsp;&nbsp;</td></tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->