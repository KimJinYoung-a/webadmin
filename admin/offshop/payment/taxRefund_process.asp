<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 taxRefund 관리
' History : 2014.01.17 서동석
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/payment/taxRefundMngCls.asp"-->
<%
Dim cmdparam, taxrefundkey, idx, refundmonth
Dim sqlstr
cmdparam		= requestCheckVar(request("cmdparam"),10)
taxrefundkey	= requestCheckVar(request("refundkey"),32)
idx				= requestCheckVar(request("midx"),10)
refundmonth     = requestCheckVar(request("refundmonth"),7)

If cmdparam = "U" Then
	If len(taxrefundkey) <> "20" Then
		response.write "<script type='text/javascript'>"
		response.write "	alert('코드는 20자리여야 합니다.다시 확인해주세요');"
		response.write "</script>"
		dbget.close() : response.end
	End If

	sqlstr = ""
	sqlstr = sqlstr & " UPDATE db_shop.dbo.tbl_shopjumun_master SET "
	sqlstr = sqlstr & " taxrefundkey = '"&taxrefundkey&"' "
	sqlstr = sqlstr & " WHERE idx = '"&idx&"' "
	dbget.execute sqlstr

	response.write "<script type='text/javascript'>"
	response.write "	alert('저장 되었습니다');"
	response.write "	parent.location.reload();"
	response.write "</script>"
	dbget.close() : response.end

ElseIf cmdparam = "D" Then
	sqlstr = ""
	sqlstr = sqlstr & " UPDATE db_shop.dbo.tbl_shopjumun_master SET "
	sqlstr = sqlstr & " taxrefundkey = Null "
	sqlstr = sqlstr & " WHERE idx = '"&idx&"' "
	dbget.execute sqlstr

	response.write "<script type='text/javascript'>"
	response.write "	alert('삭제 되었습니다');"
	response.write "	parent.location.reload();"
	response.write "</script>"
	dbget.close() : response.end
End If
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->