<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/sitemaster/EmoDic/EmoDicCls.asp" -->
<%

dim eNumber
eNumber = request("eno")
dim eType 
eType = request("etp")
dim eTitle
eTitle = request("etlt")
IF eNumber="" or eType="" Then
	response.write "�߸��� �����Դϴ�"
	dbget.close()	:	response.End
End if

dim oWord,iLp
set oWord =	new EmodicCls
oWord.FRectEmoNumber = eNumber
oWord.FRectEmoType = eType
oWord.FRectEmoTitle = eTitle
oWord.getEmoWordsList

IF oWord.FResultCount<0 Then
	response.write "���� �����ڿ��� �������ּ���"
	dbget.close()	:	response.End
End if


dim EmoNo, EmoType, EmoDesc, EmoTitle, EmoImage, EmoUsing,EmoImageUrl
iLp=0
EmoNo = oWord.FList(iLp).EmoNO
EmoType = oWord.FList(iLp).EmoType
EmoDesc = oWord.FList(iLp).EmoDesc
EmoTitle = oWord.FList(iLp).EmoTitle
EmoImage = oWord.FList(iLp).EmoImage
EmoUsing = oWord.FList(iLp).EmoUsing
EmoImageUrl = oWord.FList(iLp).getImgUrl()
SET oWord = nothing

function getETypeStr(eTp)
	dim eStr 
	Select Case etp
	Case "1"
		eStr = "��������"	
	Case"2"
		eStr = "�󷷶׶�"	
	Case "3"
		eStr = "�̼�����"		
	Case "4"
		eStr = "��������"
	End Select
	getETypeStr=eStr
End Function

%>
<table width="350" border="0" class="a" cellpadding="2" cellspacing="1" align="left" bgcolor="<%=adminColor("tablebg") %>">
<form name="" method="post" action="<%= uploadUrl %>/linkweb/Emodic_Upload.asp" Enctype="multiPart/form-data"> 
<input type="hidden" name="eno" value="<%= eNumber %>">
<input type="hidden" name="etp" value="<%= eType %>">
<input type="hidden" name="etlt" value="<%= EmoTitle %>">
<input type="hidden" name="rgflag" value="edit">
<tr bgcolor="#FFFFFF">
	<td colspan="2" bgcolor="<%=adminColor("tablebar")%>">
		<b><%=eNumber %>�� - <%=getETypeStr(eType) %></b> ����
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="40">�ܾ�</td>
	<td width="295"><input type="text" name="stitle" value="<%= EmoTitle%>"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="40">�̹���</td>
	<td><input type="file" name="simage" size="25" value=""><br><% IF EmoImage<>"" Then response.write EmoImageUrl %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="40">����</td>
	<td><textarea name="sdesc" cols="40" rows="8"><%= EmoDesc%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="40">���<br>����</td>
	<td>
		<input type="radio" name="sisusing" value="Y" <% IF EmoUsing="Y" Then Response.write "checked" %>>Y
		<input type="radio" name="sisusing" value="N" <% IF EmoUsing="N" Then Response.write "checked" %>>N
	</td>
</tr>

<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="right"><input type="submit" class="button" value="���"></td>
</tr>
</form>
</table>
	
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->