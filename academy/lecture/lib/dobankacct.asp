<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/lec_bankacctcls.asp"-->
<%
dim mode,orderidx
mode = RequestCheckvar(request("mode"),16)
orderidx = request("orderidx")

'response.write mode + "<br>"
'response.write orderidx + "<br>"
if orderidx <> "" then
	if checkNotValidHTML(orderidx) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if Len(orderidx)>0 then
	orderidx = Left(orderidx,Len(orderidx)-1)
end if

dim sqlStr,i
dim ibank

set ibank = new CBankAcct

if mode="del" then

	''한정수량조정
	sqlStr = "update [db_academy].[dbo].tbl_lec_item" + vbCrlf
	sqlStr = sqlStr + " set limit_sold=limit_sold - T.ttlitemno" + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlStr = sqlStr + " ("
	sqlStr = sqlStr + " select itemid, sum(itemno) as ttlitemno from "
	sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_master m, "
	sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d "
	sqlStr = sqlStr + " where m.idx=d.masteridx "
	sqlStr = sqlStr + " and m.idx in (" + orderidx + ")"
	sqlStr = sqlStr + " and m.ipkumdiv='2'"
	sqlStr = sqlStr + " and m.accountdiv='7'"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	sqlStr = sqlStr + " group by itemid "
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_academy].[dbo].tbl_lec_item.idx=T.Itemid"

''대기자 있는경우.. 수량조절 안해야..
'response.write sqlStr
'''	rsACADEMYget.Open sqlStr,dbACADEMYget,1


	''취소처리
	sqlStr = "update [db_academy].[dbo].tbl_academy_order_master"
	sqlStr = sqlStr + " set cancelyn='Y'"
	sqlStr = sqlStr + " where idx in (" + orderidx + ")"
	sqlStr = sqlStr + " and ipkumdiv='2'"
	sqlStr = sqlStr + " and accountdiv='7'"
	sqlStr = sqlStr + " and cancelyn='N'"

'response.write sqlStr
	rsACADEMYget.Open sqlStr,dbACADEMYget,1



'dbget.close()	:	response.End

	ibank.GetMileageSpendList orderidx
	for i=0 to ibank.FResultCount-1
		response.write CStr(i) + "<br>"
		ibank.FItemList(i).DelMilegelog
		ibank.FItemList(i).RecalcuCurrentMileage
	next

	ibank.GetCardSpendList orderidx
	for i=0 to ibank.FResultCount-1
		response.write CStr(i) + "<br>"
		ibank.FItemList(i).DelCardSpend

	next

elseif mode="mail" then

	sqlStr = "Insert into [110.93.128.72].[db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "

	sqlStr = sqlStr + " select distinct buyhp, '02-741-9070','1',getdate(),'무통장입금유효기간이틀남았습니다.입금미확인건은자동취소됩니다.더핑거스아카데미^^'"
	sqlStr = sqlStr + "  from [db_academy].[dbo].tbl_academy_order_master"
	sqlStr = sqlStr + " where idx in (" + orderidx + ")"
	sqlStr = sqlStr + " and cancelyn='N'"
	sqlStr = sqlStr + " and ipkumdiv='2'"
	sqlStr = sqlStr + " and accountdiv='7'"

	rsACADEMYget.Open sqlStr,dbACADEMYget,1


	sqlStr = " insert into [db_academy].[dbo].tbl_bankmail_sendlist(orderserial)"
	sqlStr = sqlStr + " select distinct orderserial "
	sqlStr = sqlStr + "  from [db_academy].[dbo].tbl_academy_order_master"
	sqlStr = sqlStr + " where idx in (" + orderidx + ")"
	sqlStr = sqlStr + " and cancelyn='N'"
	sqlStr = sqlStr + " and ipkumdiv='2'"
	sqlStr = sqlStr + " and accountdiv='7'"

	rsACADEMYget.Open sqlStr,dbACADEMYget,1

'dbget.close()	:	response.End
end if

set ibank = Nothing

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
<% if mode="mail" then %>
alert('발송 되었습니다.');
<% else %>
alert('취소처리 되었습니다.');
<% end if %>
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->