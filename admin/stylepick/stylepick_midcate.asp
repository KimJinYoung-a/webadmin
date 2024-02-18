<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

<%
	Dim ocateone, ocate, vCode, vMidCode, vMidCodeName, vOrderNo, vIsUsing, i
	vCode = Request("code")
	vMidCode = Request("midcode")
	
	If vMidCode <> "" Then
		set ocateone = new cstylepickMenu
		ocateone.frectcd3 = vMidCode
		ocateone.getstylepick_cate_cd3_one()
		
		if ocateone.ftotalcount > 0 then
			vMidCodeName = ocateone.foneitem.fcatename
			vIsUsing = ocateone.foneitem.fisusing
			vOrderNo = ocateone.foneitem.forderno		
		end if
	End If
	
	set ocate = new cstylepickMenu
	ocate.frectcd3 = vCode
	ocate.getstylepick_cate_cd3()
	
	
	If vCode = "1" Then
		Response.Write "�з� : <b>Stationery & Personal</b>"
	ElseIf vCode = "2" Then
		Response.Write "�з� : <b>Home & Living</b>"
	ElseIf vCode = "3" Then
		Response.Write "�з� : <b>Fashion & Beauty</b>"
	ElseIf vCode = "4" Then
		Response.Write "�з� : <b>Kidult & Hobby</b>"
	ElseIf vCode = "5" Then
		Response.Write "�з� : <b>Kids & Baby</b>"
	ElseIf vCode = "6" Then
		Response.Write "�з� : <b>Digital & Camera</b>"
	End If
	
	if vIsUsing = "" then vIsUsing = "Y"
%>

<script language="javascript">
function chkfrm()
{
	if(frm.midcodename.value == "")
	{
		alert("�ߺз��ڵ���� �Է��ϼ���.");
		return false;
	}
	if(frm.orderno.value == "")
	{
		alert("���ļ����� �Է��ϼ���.");
		return false;
	}
	if(frm.isusing.value == "")
	{
		alert("���ļ����� �Է��ϼ���.");
		return false;
	}
	return true;
}
</script>

<form name="frm" method="post" action="stylepick_midcate_proc.asp" onSubmit="return chkfrm()">
<input type="hidden" name="code" value="<%=vCode%>">
<input type="hidden" name="midcode" value="<%=vMidCode%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% If vMidCode <> "" Then %>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�ߺз��ڵ�</td>
    <td><%=vMidCode%></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�ߺз��ڵ��</td>
    <td><input type="text" name="midcodename" value="<%=vMidCodeName%>" maxlength="32" size="32"></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">���ļ���</td>
    <td><input type="text" name="orderno" value="<%=vOrderNo%>" maxlength="2" size="2"> ex) 1 ~ 99 ���ڰ� �������� �켱����</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">��뿩��</td>
    <td>
		<select name="isusing">
			<option value="" <% if vIsUsing="" then response.write " selected" %>>����</option>
			<option value="Y" <% if vIsUsing="Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if vIsUsing="N" then response.write " selected" %>>N</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2">
    	<% If vMidCode = "" Then %>
    		<input type="submit" value="�ű�����"  class="button">
    	<% else %>
    		<input type="submit" value="����"  class="button">
    	<% end if %>
    </td>
</tr>
</table>
</form>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= ocate.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">	
	<td>ī�װ��ڵ�</td>
	<td>ī�װ���</td>
	<td>���ļ���</td>
	<td>��뿩��</td>
	<td>����������</td>
	<td>���</td>
</tr>
<% if ocate.FresultCount>0 then %>
<% for i=0 to ocate.FresultCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">			

<% if ocate.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
<% end if %>

	<td>
		<%= ocate.FItemList(i).fcd1 %><%= ocate.FItemList(i).fcd2 %><%= ocate.FItemList(i).fcd3 %>
	</td>
	<td>
		<%= ocate.FItemList(i).fcatename %>
	</td>
	<td>
		<%= ocate.FItemList(i).forderno %>
	</td>
	<td>
		<%= ocate.FItemList(i).fisusing %>
	</td>
		
	<td>
		<%= ocate.FItemList(i).flastadminid %>
	</td>
	<td>
		<input type="button" onclick="location.href='?code=<%=vCode%>&midcode=<%=ocate.FItemList(i).fcd3%>'" value="����" class="button">
	</td>
</tr>   
</form>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<% 
set ocate = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->