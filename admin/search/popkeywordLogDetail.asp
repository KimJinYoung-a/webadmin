<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/search/itemCls.asp" -->
<%
Dim idx, arrList, oKeyList, i, rowNum
idx		= request("idx")
SET oKeyList = new cItemContent
	oKeyList.FRectIdx = idx
	oKeyList.getKeyWordLogDetailList
	arrList = oKeyList.fnkeywordMaster(idx)
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}
</script>
<table width="100%">
<tr>
	<td align="right"><input type="button" value="���" class="button" onclick="location.href='/admin/search/popkeywordLog.asp';"></td>
</tr>
</table>
<p />
<table width="100%">
<tr>
	<td align="LEFT"><strong>���� �̷� ����</strong></td>
	<td align="RIGHT">*��Ͽ��� �ٷ� ���� ������ Ű���� ������ �����Դϴ�.</td>
</tr>
<tr>
	<td colspan="2">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="15%">���� ����</td>
			<td width="35%" bgcolor="#FFFFFF" align="LEFT">
			<%
				Select Case arrList(1, 0)
					Case "I"		response.write "���"
					Case "U"		response.write "����"
					Case "D"		response.write "����"
				End Select
			%>
			</td>
			<td width="15%">���� Ű����</td>
			<td width="35%" bgcolor="#FFFFFF" align="LEFT">
			<%
				If arrList(1, 0) = "U" Then
					If arrList(4, 0) <> "" Then
						response.write arrList(3, 0) & " �� " & arrList(4, 0)
					Else
						response.write arrList(4, 0)
					End If
				Else
					response.write arrList(4, 0)
				End If
			%>	
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="15%">������</td>
			<td width="35%" bgcolor="#FFFFFF" align="LEFT"><%= arrList(8, 0) %></td>
			<td width="15%">������</td>
			<td width="35%" bgcolor="#FFFFFF" align="LEFT"><%= arrList(7, 0) %></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="15%">���</td>
			<td colspan="3" bgcolor="#FFFFFF" align="LEFT"><%= arrList(5, 0) %></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="15%">���� ����</td>
			<td colspan="3" bgcolor="#FFFFFF" align="LEFT"><%= arrList(2, 0) %></td>
		</tr>
	</td>
</tr>	
</table>
<br/>
<table width="100%">
<tr>
	<td align="LEFT"><strong>���� ��ǰ ����</strong></td>
</tr>
<tr>
	<td colspan="2">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="30">��ȣ</td>
			<td width="100">����ī�װ�</td>
			<td width="50">��ǰ�ڵ�</td>
			<td width="50">�̹���</td>
			<td width="100">�귣��ID</td>
			<td width="250">��ǰ��</td>
			<td>Ű����</td>
		</tr>
	<%
		rowNum = oKeyList.FTotalcount
		For i = 0 to oKeyList.FResultCount - 1
	%>
		<tr align="center" bgcolor="#FFFFFF" height="30">
			<td><%= rowNum %></td>
			<td><%= oKeyList.FItemList(i).FCatename %></td>
			<td><%= oKeyList.FItemList(i).FItemid %></td>
			<td><img src="<%= oKeyList.FItemList(i).Fsmallimage %>" width="50"></td>
			<td><%= oKeyList.FItemList(i).FMakerid %></td>
			<td><%= oKeyList.FItemList(i).FItemname %></td>
			<td><%= oKeyList.FItemList(i).FKeywords %></td>
		</tr>
	<%
			rowNum = rowNum - 1 
		Next
	%>
		<tr align="center" bgcolor="#FFFFFF" height="30">
			<td colspan="7"><input type="button" class="button" value="�ݱ�" onclick="self.close();"></td>
		</tr>			
		</table>
	</td>
</tr>
</table>
<% SET oKeyList = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->