<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 정산리스트
' History : 2009.04.07 서동석 생성
'			2012.09.26 한용민 수정
'####################################################
%>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->

<%
dim idx ,ofranchulgojungsan, shopid
	idx=request("idx")

if idx="" then idx="0"

set ofranchulgojungsan = new CFranjungsan
	ofranchulgojungsan.FRectidx = idx
	ofranchulgojungsan.getOneFranJungsan
%>

<table width="760" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_number %></td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Fidx %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_SHOP %></td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Fshopid %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_divide %></td>
	<td bgcolor="#FFFFFF" ><font color="<%= ofranchulgojungsan.FOneItem.GetDivCodeColor %>"><%= ofranchulgojungsan.FOneItem.GetDivCodeName %></font></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_title %></td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Ftitle %>
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_total_consumer_price %></td>
	<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalsellcash,0) %></td>
</tr>

<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_total_Supply_price %></td>
	<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalsuplycash,0) %>
	<font color="#AAAAAA">(샾으로 공급한 상품가격)</font></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_Bill %>&nbsp;<%= CTX_Issue %></td>
	<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalsum,0) %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_tax_Bill_date %></td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Ftaxdate %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_Deposit_Date %></td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Fipkumdate %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_Status %></td>
	<td bgcolor="#FFFFFF" >
		<font color="<%= ofranchulgojungsan.FOneItem.GetStateColor %>"><%= ofranchulgojungsan.FOneItem.GetStateName %></font>
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=120><%= CTX_etc %></td>
	<td bgcolor="#FFFFFF" >
		<textarea name="etcstr" cols=86 rows=8 style="border:1px #999999 solid; "><%= ofranchulgojungsan.FOneItem.Fetcstr %></textarea>
	</td>
</tr>
<!--
<tr>
	<td bgcolor="#DDDDFF" width=100>최초등록자</td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Fregusername %>(<%= ofranchulgojungsan.FOneItem.Freguserid %>)</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF" width=100>최종처리자</td>
	<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Ffinishusername %>(<%= ofranchulgojungsan.FOneItem.Ffinishuserid %>)</td>
</tr>
-->
<tr>
	<td colspan=2 align=center bgcolor="#FFFFFF">
		<input type=button value="CLOSE" onclick="window.close();" class="button">
	</td>
</tr>
</form>
</table>

<%
set ofranchulgojungsan = Nothing
%>
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->