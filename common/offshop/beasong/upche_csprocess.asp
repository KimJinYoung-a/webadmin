<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 cs내역
' Hieditor : 2011.03.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim masteridx,finishmemo, finishuser,songjangdiv, songjangno ,sqlStr ,oldcurrstate
	masteridx          = request("masteridx")
	finishmemo  = html2db(request("finishmemo"))
	finishuser  = request("finishuser")
	songjangdiv = request("songjangdiv")
	songjangno  = request("songjangno")

sqlStr = "select currstate "
sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_cs_master" + VbCrlf
sqlStr = sqlStr + " where masteridx =" + masteridx

'response.write sqlStr &"<Br>"
rsget.Open sqlStr,dbget,1
    oldcurrstate = rsget("currstate")
rsget.Close

if (oldcurrstate="B007") then
    response.write "<script>alert('이미 처리 완료된 내역입니다. - 완료처리로 진행 할 수 없습니다.');history.back();</script>"
    response.end
end if

sqlStr = "update db_shop.dbo.tbl_shopbeasong_cs_master" + VbCrlf
sqlStr = sqlStr + " set finishuser ='" + finishuser + "'," + VbCrlf
sqlStr = sqlStr + " contents_finish ='" + finishmemo + "'," + VbCrlf
sqlStr = sqlStr + " songjangdiv ='" + songjangdiv + "'," + VbCrlf
sqlStr = sqlStr + " songjangno ='" + songjangno + "'," + VbCrlf
sqlStr = sqlStr + " finishdate=getdate()," + VbCrlf
sqlStr = sqlStr + " currstate='B006'" + VbCrlf
sqlStr = sqlStr + " where masteridx =" + masteridx
sqlStr = sqlStr + " and makerid='" & session("ssBctID") & "'"

'response.write sqlStr &"<Br>"
rsget.Open sqlStr,dbget,1
%>

<script language='javascript'>

	alert('저장되었습니다.');
	location.replace('/common/offshop/beasong/upche_cslist.asp');

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->