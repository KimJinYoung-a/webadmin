<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ����ǰ ������Ȳ ���� (�����Ϸ��̻�, �ǽð�)
' History : 2014.10.06 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
 Dim arrGiftCd, strSql

 arrGiftCd		= requestCheckVar(Request("arr"),128)		'����ǰ�ڵ�(��ǥ����)
 if arrGiftCd="" then
 	Call Alert_Close("�μ�����")
 	dbget.close()
 	response.End
 End if

	'�ð��� ���� ���� ������ ���� ;;;
	strSql = "select g.chg_gift_code, g.chg_giftSTR, count(*) cnt "
	strSql = strSql & "from db_order.dbo.tbl_order_master as m "
	strSql = strSql & "	join db_order.dbo.tbl_order_gift as g "
	strSql = strSql & "		on m.orderserial=g.orderserial "
	strSql = strSql & "where m.ipkumdiv>3 "
	strSql = strSql & "	and m.jumundiv<>9 "
	strSql = strSql & "	and m.cancelyn='N' "
	strSql = strSql & "	and g.chg_gift_code in (" & arrGiftCd & ") "
	strSql = strSql & "group by g.chg_gift_code, g.chg_giftSTR"
	rsget.Open strSql, dbget, 1

%>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="3">�˻���� : <b><%=rsget.RecordCount %></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>����ǰ�ڵ�</td>
    	<td>����ǰ��</td>
    	<td>����</td>
    </tr>
<%
	if Not(rsget.EOF or rsget.BOF) then
    	Do Until rsget.EOF
%>
    <tr align="center" bgcolor="#FFFFFF">
    	<td nowrap><%=rsget("chg_gift_code")%></td>
    	<td nowrap><%=rsget("chg_giftSTR")%></td>
    	<td nowrap><%=rsget("cnt")%></td>
    </tr>
<%
		rsget.MoveNext
		Loop
	ELSE
%>
	<tr>
		<td colspan="17" align="center" bgcolor="#FFFFFF">���� ������ �����ϴ�.</td>
	</tr>
<%	END IF %>
</table>
<p class="a">�� ���� �Ϸ��̻�, �����ֹ���, ���� ������ ����ǰ ����</p>
<%
	rsget.Close()
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->