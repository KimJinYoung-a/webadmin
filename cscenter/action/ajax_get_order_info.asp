<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<%

dim orderserial
	orderserial = RequestCheckVar(request("orderserial"),11)

dim ojumun
set ojumun = new COrderMaster

if (orderserial <> "") then
    ojumun.FRectOrderSerial = orderserial
    ojumun.QuickSearchOrderMaster
end if

if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster
end if

dim code, msg
dim userid, buyhp, reqhp

if (ojumun.FResultCount < 1) then
	code = "99"
	msg = "주문내역 없음"
else
	code = "00"
	msg = "OK"
	userid = ojumun.FOneItem.FUserID
	buyhp = ojumun.FOneItem.FBuyHp
	reqhp = ojumun.FOneItem.FReqHp
end if

response.write "{"
response.write """code"":""" & code & ""","
response.write """msg"":""" & msg & ""","
response.write """orderserial"":""" & orderserial & ""","
response.write """userid"":""" & userid & ""","
response.write """buyhp"":""" & buyhp & ""","
response.write """reqhp"":""" & reqhp & """"
response.write "}"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
