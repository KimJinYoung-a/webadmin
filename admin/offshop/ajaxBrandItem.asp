<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 브랜드 상품
' History : 최초생성자모름
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Response.CharSet = "euc-kr"

Dim makerid		: makerid		= requestCheckVar(request("makerid",""),32)
Dim itemgubun	: itemgubun		= requestCheckVar(request("itemgubun",""),2)
Dim itemid		: itemid		= requestCheckVar(request("itemid",""),10)
Dim itemoption	: itemoption	= requestCheckVar(request("itemoption",""),4)
Dim shopid      : shopid	    = requestCheckVar(request("shopid",""),32)
dim sqlStr
Dim ret

sqlStr = " SELECT i.isUsing, i.shopItemName, i.shopItemOptionName, i.orgSellPrice, i.shopItemPrice " & vbCrLf
sqlStr = sqlStr & " ,CASE WHEN i.shopsuplycash=0 THEN convert(int,i.shopItemPrice*(100-d.defaultmargin)/100) ELSE i.shopsuplycash END as buycash" & vbCrLf
sqlStr = sqlStr & " ,CASE WHEN i.shopbuyprice=0 THEN convert(int,i.shopbuyprice*(100-d.defaultsuplymargin)/100) ELSE i.shopbuyprice END as suplycash" & vbCrLf
sqlStr = sqlStr & " FROM db_shop.dbo.tbl_shop_item i" & vbCrLf
sqlStr = sqlStr & "     left join db_shop.dbo.tbl_shop_designer d" & vbCrLf
sqlStr = sqlStr & "     on i.makerid=d.makerid" & vbCrLf
sqlStr = sqlStr & "     and d.shopid='"&shopid&"'" & vbCrLf
sqlStr = sqlStr & " WHERE 1=1 " & vbCrLf
sqlStr = sqlStr & " AND i.makerid		= '" & makerid		& "'" & vbCrLf
sqlStr = sqlStr & " AND i.itemgubun	= '" & itemgubun	& "'" & vbCrLf
sqlStr = sqlStr & " AND i.shopitemid	= '" & itemid		& "'" & vbCrLf
sqlStr = sqlStr & " AND i.itemoption	= '" & itemoption	& "'" & vbCrLf

'rw sqlStr

rsget.Open sqlStr, dbget, 1
If Not rsget.EOF  Then 
	If rsget("isUsing") = "Y" Then 
		response.write "Y|" & rsget("shopItemName") & "|" & rsget("shopItemOptionName") & "|" & rsget("orgSellPrice")  & "|" & rsget("shopItemPrice") & "|"& rsget("buycash") & "|"& rsget("suplycash") & "|"
	Else
		response.write "N|" & rsget("shopItemName") & "|" & rsget("shopItemOptionName") & "|0|0|0|0|"
	End If 
Else 
	response.write "|||||"
End If 
rsget.close

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->