<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->

<%
dim oip, vIdx, vRealTitle, vTitle, vContents, vRegdate, vIsusing, vMainImg, vSubImg
vIdx = Request("idx")
If vIdx <> "" Then
set oip = new organizerCls
	oip.FOW_IDX = vIdx
	oip.FOrgStory
	
	vIdx = oip.FOneItem.FOW_IDX
	vTitle = oip.FOneItem.FOW_TITLE
	vContents = oip.FOneItem.FOW_CONTENTS
	vRegdate = oip.FOneItem.FOW_REGDATE
	vIsusing = oip.FOneItem.FOW_ISUSING
set oip = nothing
End IF
%>
���ų����� ���丮 �ۼ�
<script language='javascript'>

function SaveMainContents(frm){
    if (frm.title.value.length<1){
        alert('������ �Է� �ϼ���.');
        frm.title.focus();
        return;
    }
    
    if (frm.contents.value.length<1){
        alert('������ �Է� �ϼ���.');
        frm.contents.focus();
        return;
    }
  
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}
</script>
<!-- ����Ʈ ���� -->
<table border="0" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmcontents" method="post" action="story_proc.asp" onsubmit="return false;">
<input type="hidden" name="idx" value="<%=vIdx%>">
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td nowrap align="center" width="100"> ��ȣ</td>
		<td bgcolor="#FFFFFF"><%=vIdx%></td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td nowrap align="center"> ���� </td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="title" value="<%=vTitle%>" size="100">
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td nowrap align="center"> ���� </td>
		<td bgcolor="#FFFFFF">
			<textarea name="contents" cols="100" rows="20"><%=vContents%></textarea>
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td nowrap align="center"> ��뿩�� </td>
		<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <% If vIsusing = "Y" Then Response.Write "checked" End If %>> Y&nbsp;&nbsp;&nbsp;
		<input type="radio" name="isusing" value="N" <% If vIsusing = "N" Then Response.Write "checked" End If %>> N
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td nowrap align="center"> �ۼ��� </td>
		<td align="center" bgcolor="#FFFFFF"><%=vRegdate%></td>
	</tr>
</form>
</table>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" value=" �� �� " onClick="SaveMainContents(frmcontents);" class="button">
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- ����Ʈ �� -->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->