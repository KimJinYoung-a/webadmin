<%@ language=vbscript %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/chkItem.asp"-->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/itemOptionLib.asp"-->
<%

function WaitRegDoubleOptionProc(waititemid)
    
    Dim optlimitno, optionyn, itemoption, loopcnt
	Dim arroptlimitno,arroptionyn,arritemoption, TotalOptLimitNo
	TotalOptLimitNo=0
	optlimitno = Request.Form("optlimitno") & ","
	optionyn = Request.Form("optionyn") & ","
	itemoption = Request.Form("itemoption") & ","

	arroptlimitno = split(optlimitno,",")
	arroptionyn = split(optionyn,",")
	arritemoption = split(itemoption,",")
	loopcnt = ubound(arritemoption)
	Dim sqlStr, i
    ''0번은 입력 없음. N까지
    For i=0 To loopcnt-1
		TotalOptLimitNo = TotalOptLimitNo +  Cint(trim(arroptlimitno(i)))
		sqlStr = "update db_academy.dbo.tbl_diy_wait_item_option"
		sqlStr = sqlStr + " set isusing='" + CStr(trim(arroptionyn(i))) + "'"
		sqlStr = sqlStr + " ,optlimitno='" + CStr(trim(arroptlimitno(i))) + "'"
		sqlStr = sqlStr + " where itemid=" + CStr(trim(waititemid))
		sqlStr = sqlStr + " and itemoption='" + CStr(trim(arritemoption(i))) + "'"
		dbACADEMYget.Execute sqlStr
    Next
    ''옵션 총수 저장
	sqlStr = "update db_academy.dbo.tbl_diy_wait_item"
	sqlStr = sqlStr + " set optioncnt=(select count(itemid) as cnt from db_academy.dbo.tbl_diy_wait_item_option where itemid =" + CStr(waititemid) + " and isusing = 'Y')"
	sqlStr = sqlStr + " , limitno=(select sum(optlimitno) as cnt from db_academy.dbo.tbl_diy_wait_item_option where itemid =" + CStr(waititemid) + " and isusing = 'Y')"
	sqlStr = sqlStr + " ,limityn=(case"
	sqlStr = sqlStr + " when 0 < (select sum(optlimitno) as cnt from db_academy.dbo.tbl_diy_wait_item_option where itemid =" + CStr(waititemid) + " and isusing = 'Y') then 'Y'"
	sqlStr = sqlStr + " when 0 >= (select sum(optlimitno) as cnt from db_academy.dbo.tbl_diy_wait_item_option where itemid =" + CStr(waititemid) + " and isusing = 'Y') then 'N'"
	sqlStr = sqlStr + " end)"
	sqlStr = sqlStr + "where itemid = " + CStr(waititemid) + " "
    dbACADEMYget.Execute sqlStr
	WaitRegDoubleOptionProc=TotalOptLimitNo
End Function

Dim waititemid, TotalOptLimitNo, makerid
waititemid = requestCheckVar(request("waititemid"),10)
makerid = request.cookies("partner")("userid")

'If (WaitItemCheckMyItemYN(makerid,waititemid)<>true) Then
'	Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
'	Response.end
'End If
TotalOptLimitNo = WaitRegDoubleOptionProc(waititemid)
if (TotalOptLimitNo<>"") then
%>
<script type="text/javascript">
<!--
	parent.fnOptionStateEditEnd('<%=TotalOptLimitNo%>');
//-->
</script>
<%
end if
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->