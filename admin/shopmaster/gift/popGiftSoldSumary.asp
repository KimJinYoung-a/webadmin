<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 사은품 지급현황 보기 (결제완료이상, 실시간)
' History : 2014.10.06 허진원 생성
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

 arrGiftCd		= requestCheckVar(Request("arr"),128)		'상은품코드(쉼표구분)
 if arrGiftCd="" then
 	Call Alert_Close("인수없음")
 	dbget.close()
 	response.End
 End if

	'시간이 없어 대충 빠르게 만듦 ;;;
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
		<td colspan="3">검색결과 : <b><%=rsget.RecordCount %></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>사은품코드</td>
    	<td>사은품명</td>
    	<td>수량</td>
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
		<td colspan="17" align="center" bgcolor="#FFFFFF">지급 내역이 없습니다.</td>
	</tr>
<%	END IF %>
</table>
<p class="a">※ 결제 완료이상, 정상주문건, 고객이 선택한 사은품 기준</p>
<%
	rsget.Close()
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->