<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim orderserial : orderserial = requestCheckVar(request("orderserial"),11)
dim itemid      : itemid = requestCheckVar(request("itemid"),11)
dim didx        : didx = requestCheckVar(request("didx"),11)
dim itemcost    : itemcost = requestCheckVar(request("itemcost"),11)
dim buycash     : buycash = requestCheckVar(request("buycash"),11)
dim onlybuycash : onlybuycash = requestCheckVar(request("onlybuycash"),10)

dim jMonth : jMonth="2020-01-01"

dim sqlStr, AssignedRow

sqlStr = " update db_order.dbo.tbl_order_detail " & VbCRLF
sqlStr = sqlStr & " set buycash="&buycash & VbCRLF
if (onlybuycash="") then
sqlStr = sqlStr & " ,buycashCouponNotApplied="&buycash & VbCRLF
end if
sqlStr = sqlStr & " ,issailitem='Y'" & VbCRLF
sqlStr = sqlStr & " where itemid="&itemid & VbCRLF
sqlStr = sqlStr & " and orderserial='"&orderserial&"'" & VbCRLF
sqlStr = sqlStr & " and idx="&didx
sqlStr = sqlStr & " and isNULL(beasongdate,'2020-01-01')>='"&jMonth&"'"

dbget.Execute sqlStr,AssignedRow
rw AssignedRow&"°Ç ¹Ý¿µµÊ"

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
