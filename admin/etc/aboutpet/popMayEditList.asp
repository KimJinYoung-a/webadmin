<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim strSQL, arrRows, i, apiaction
Dim itemidarr, kakaoidarr

apiaction = request("apiaction")

strSQL = " exec [db_etcmall].[dbo].[usp_Ten_OutMall_Aboutpet_MayEditList] '"& apiaction &"' "
rsget.CursorLocation = adUseClient
rsget.CursorType=adOpenStatic
rsget.Locktype=adLockReadOnly
rsget.Open strSQL, dbget
If Not(rsget.EOF or rsget.BOF) Then
	arrRows = rsget.getRows()
End If
rsget.Close
%>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�׼� :
			<select name="apiaction" class="select">
				<option value="">-����-</option>
				<option <%= Chkiif(apiaction="SOLDOUT", "selected", "") %> value="SOLDOUT">ǰ�����</option>
				<option <%= Chkiif(apiaction="PRICE", "selected", "") %> value="PRICE">����</option>
				<option <%= Chkiif(apiaction="ITEMNAME", "selected", "") %> value="ITEMNAME">��ǰ��</option>
				<option <%= Chkiif(apiaction="OPTNAME", "selected", "") %> value="OPTNAME">�ɼǸ�</option>
				<option <%= Chkiif(apiaction="SELLCHG", "selected", "") %> value="SELLCHG">�Ǹ���ȯ���</option>
			</select>
		<br>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<p />

<table width="100%" align="center">
<tr align="center">
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>�ٹ����� ��ǰ�ڵ�</td>
			<td>�ɼ��ڵ�</td>
			<td>�׼�</td>
			<td>��ǰ��</td>
			<td>aboutpet��ǰ��</td>
			<td>�ɼǸ�</td>
			<td>aboutpet�ɼǸ�</td>
			<td>�ǸŰ�</td>
			<td>aboutpet�ǸŰ�</td>
			<td>�ǸŻ���</td>
		</tr>
<%
	if isArray(ArrRows) then
		For i = 0 To UBound(ArrRows,2)
%>
		<tr align="center" bgcolor="#FFFFFF">
			<td><%= ArrRows(0, i) %></td>
			<td><%= ArrRows(4, i) %></td>
			<td><%= ArrRows(1, i) %></td>
			<td><%= ArrRows(2, i) %></td>
			<td><%= ArrRows(3, i) %></td>
			<td><%= ArrRows(5, i) %></td>
			<td><%= ArrRows(6, i) %></td>
			<td><%= ArrRows(7, i) %></td>
			<td><%= ArrRows(8, i) %></td>
			<td><%= ArrRows(9, i) %></td>
		</tr>
<%
			i=i+1
		Next
	End If
%>
		</table>
	</td>
</tr>
</table>


<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
