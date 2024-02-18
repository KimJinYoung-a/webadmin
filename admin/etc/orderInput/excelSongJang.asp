<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
	Dim strSql, arrRows, i, mallid
	mallid = request("mallid")

	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename="& mallid & replace(DATE(), "-", "") &"_xl.xls"

	If mallid = "aboutpet" Then
		strSql = ""
		strSql = strSql & " SELECT '어바웃펫' as sitename, '10' as beasonggbn, T.OutMallOrderSerial "
		strSql = strSql & " , T.beasongNum11st "
		strSql = strSql & " , CASE WHEN V.divcd = '1' THEN '05' "
		strSql = strSql & " 		WHEN V.divcd = '2' THEN '37' "
		strSql = strSql & " 		WHEN V.divcd = '3' THEN '01' "
		strSql = strSql & " 		WHEN V.divcd = '4' THEN '01' "
		strSql = strSql & " 		WHEN V.divcd = '8' THEN '04' "
		strSql = strSql & " 		WHEN V.divcd = '9' THEN '07' "
		strSql = strSql & " 		WHEN V.divcd = '18' THEN '02' "
		strSql = strSql & " 		WHEN V.divcd = '21' THEN '11' "
		strSql = strSql & " 		WHEN V.divcd = '26' THEN '10' "
		strSql = strSql & " 		WHEN V.divcd = '29' THEN '32' "
		strSql = strSql & " 		WHEN V.divcd = '31' THEN '13' "
		strSql = strSql & " 		WHEN V.divcd = '34' THEN '09' "
		strSql = strSql & " 		WHEN V.divcd = '35' THEN '14' "
		strSql = strSql & " 		WHEN V.divcd = '37' THEN '12' "
		strSql = strSql & " 		WHEN V.divcd = '38' THEN '08' "
		strSql = strSql & " 		WHEN V.divcd = '39' THEN '03' "
		strSql = strSql & " 		WHEN V.divcd = '42' THEN '14' "
		strSql = strSql & " 		WHEN V.divcd = '46' THEN '15' "
		strSql = strSql & " 		WHEN V.divcd = '90' THEN '27' "
		strSql = strSql & " 		WHEN V.divcd = '91' THEN '29' "
		strSql = strSql & " ELSE '택배사확인' END as divCode "
		strSql = strSql & " , Replace(D.songjangNo, '-', '') as songjangNo "
		strSql = strSql & " , '위탁_텐바이텐' as tenName, T.orderItemName, isNull(T.orderItemOptionName, '') as orderItemOptionName, T.ItemOrderCount, T.ReceiveHpNo, T.ReceiveName "
		strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder T "
		strSql = strSql & " JOIN db_order.dbo.tbl_order_master M on T.orderserial=M.orderserial "
		strSql = strSql & " JOIN db_order.dbo.tbl_order_detail D on T.orderserial=D.orderserial "
		strSql = strSql & " 	and IsNull(T.changeitemid, T.matchitemid)=D.itemid "
		strSql = strSql & " 	and IsNull(T.changeitemoption, T.matchitemoption)=D.itemoption and D.currstate=7 "
		strSql = strSql & " LEFT JOIN db_order.dbo.tbl_songjang_div V on D.songjangDiv=V.divcd "
		strSql = strSql & " WHERE 1=1 "
		strSql = strSql & " and T.sellsite='"& mallid &"' "
		strSql = strSql & " and T.OrgDetailKey is Not NULL "
		strSql = strSql & " GROUP BY T.OutMallOrderSerial,T.OrgDetailKey, V.divcd, D.songjangNo, T.beasongNum11st, T.orderItemName,T.orderItemOptionName, T.ItemOrderCount, T.ReceiveHpNo, T.ReceiveName"
		strSql = strSql & " ORDER BY T.OutMallOrderSerial DESC "
	Else
		strSql = ""
		strSql = strSql & " SELECT T.OutMallOrderSerial, T.OrgDetailKey, V.divname, D.songjangNo "
		strSql = strSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder T "
		strSql = strSql & " JOIN db_order.dbo.tbl_order_master M on T.orderserial=M.orderserial "
		strSql = strSql & " JOIN db_order.dbo.tbl_order_detail D on T.orderserial=D.orderserial "
		strSql = strSql & " 	and IsNull(T.changeitemid, T.matchitemid)=D.itemid	 "
		strSql = strSql & " 	and IsNull(T.changeitemoption, T.matchitemoption)=D.itemoption "
		strSql = strSql & " and D.currstate=7 "
		strSql = strSql & " LEFT JOIN db_order.dbo.tbl_songjang_div V on D.songjangDiv=V.divcd "
		strSql = strSql & " WHERE 1=1"
		strSql = strSql & " and T.sellsite='"& mallid &"' "
		strSql = strSql & " and T.OrgDetailKey is Not NULL "
		strSql = strSql & " GROUP BY T.OutMallOrderSerial,T.OrgDetailKey, V.divname, D.songjangNo  "
		strSql = strSql & " ORDER BY T.OutMallOrderSerial DESC "
	End If
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		arrRows = rsget.getRows()
	End If
	rsget.Close

	If mallid = "aboutpet" Then
%>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>사이트명</td>
	<td>배송 구분</td>
	<td>주문번호</td>
	<td>배송 번호</td>
	<td>택배사</td>
	<td>송장 번호</td>
	<td>업체 명</td>
	<td>상품 명</td>
	<td>단품 명</td>
	<td>배송수량</td>
	<td>휴대폰</td>
	<td>수취인 명</td>
</tr>
<%
		If isarray(arrRows) Then
			For i = 0 to Ubound(arrRows, 2)
%>
<tr align="center" bgcolor="#FFFFFF">
	<td style="mso-number-format:'\@'"><%= CSTR(arrRows(0, i)) %></td>
	<td style="mso-number-format:'\@'"><%= CSTR(arrRows(1, i)) %></td>
	<td style="mso-number-format:'\@'"><%= CSTR(arrRows(2, i)) %></td>
	<td style="mso-number-format:'\@'"><%= CSTR(arrRows(3, i)) %></td>
	<td style="mso-number-format:'\@'"><%= CSTR(arrRows(4, i)) %></td>
	<td style="mso-number-format:'\@'"><%= CSTR(arrRows(5, i)) %></td>
	<td style="mso-number-format:'\@'"><%= CSTR(arrRows(6, i)) %></td>
	<td style="mso-number-format:'\@'"><%= CSTR(arrRows(7, i)) %></td>
	<td style="mso-number-format:'\@'"><%= CSTR(arrRows(8, i)) %></td>
	<td style="mso-number-format:'\@'"><%= CSTR(arrRows(9, i)) %></td>
	<td style="mso-number-format:'\@'"><%= CSTR(arrRows(10, i)) %></td>
	<td style="mso-number-format:'\@'"><%= CSTR(arrRows(11, i)) %></td>
</tr>
<%
			Next
		End If
'''''''''''''''''''''''어바웃팻'''''''''''''''''''''''
	Else
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="150">주문번호</td>
	<td width="250">품목별 주문번호</td>
	<td width="150">택배사명</td>
	<td width="150">송장번호</td>
</tr>
<%
		If isarray(arrRows) Then
			For i = 0 to Ubound(arrRows, 2)
%>
<tr align="center" bgcolor="#FFFFFF">
	<td width="50" style="mso-number-format:'\@'"><%= CSTR(arrRows(0, i)) %></td>
	<td width="50" style="mso-number-format:'\@'"><%= CSTR(arrRows(1, i)) %></td>
	<td width="50" style="mso-number-format:'\@'"><%= CSTR(arrRows(2, i)) %></td>
	<td width="50" style="mso-number-format:'\@'"><%= CSTR(arrRows(3, i)) %></td>
</tr>
<%
			Next
		End If
	End If
%>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->