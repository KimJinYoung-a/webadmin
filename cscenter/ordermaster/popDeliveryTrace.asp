<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
response.write "�ý��� ���� ��� : ������" ''eastone
response.end

Dim i

Dim traceUrl	: traceUrl		= req("traceUrl", "")
Dim songjangNo	: songjangNo	= req("songjangNo", "")

Dim strSql
strSql = " SELECT (case DLV_GB when '13' then '����' when '22' then '���' when '12' then '������' when '21' then '�̹��' else '' end)	" & vbCrLf
strSql = strSql & " , (case COM_GB when '1' then '����' when '3' then 'ȸ��' else '' end)	" & vbCrLf
strSql = strSql & " , RTN_PV_NM, BRAN_NM, BRAN_TEL, DLV_EMP, PIC_DT	" & vbCrLf
strSql = strSql & " FROM db_logics.dbo.tbl_V_DIST_DLV_CPL_SE_TENBYTEN	" & vbCrLf
strSql = strSql & " WHERE 1=1	" & vbCrLf
strSql = strSql & " AND WBL_NO = '" & songjangNo & "'	" & vbCrLf
strSql = strSql & " ORDER BY PIC_DT	" & vbCrLf



db3_rsget.Open strSql, db3_dbget, 1

Dim rs 
If Not db3_rsget.EOF Then
	rs = db3_rsget.getRows()
End If 
db3_rsget.close




%>
<div align="center">
<%
If IsArray(rs) Then 
%>
<!-- ����Ʈ ���� -->
<table width="700" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td>��¥</td>
        <td>����</td>
        <td>��ȭ</td>
        <td>���</td>
        <td colspan="2">����</td>
        <td>������Ȳ</td>
	</tr>
	<%For i=0 To UBound(rs,2)%>
    <tr align="center" bgcolor="#FFFFFF">
	<%
	%>
		<td><%=Left(rs(6,i),4) & "-" & Mid(rs(6,i),5,2) & "-" & Right(rs(6,i),2) %></td>
		<td><%=rs(3,i)%></td>
		<td><%=rs(4,i)%></td>
		<td><%=rs(5,i)%></td>
		<td><%=rs(0,i)%></td>
		<td><%=rs(1,i)%></td>
		<td><%=rs(2,i)%></td>
	</tr>
	<%Next%>
</table>
<%
End If 
%>

<iframe width="700" height="570" scrolling="no" frameborder="0" src="<%=traceUrl%><%=songjangNo%>"></iframe>
<br><br>
<input type="button" class="button" value="â�ݱ�" onclick="window.close();">
</div>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
