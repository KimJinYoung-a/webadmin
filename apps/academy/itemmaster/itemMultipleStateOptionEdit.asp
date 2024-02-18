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

function RegDoubleOptionProc(itemid)
    
    Dim optlimitno, optionyn, itemoption, loopcnt
	Dim arroptlimitno,arroptionyn,arritemoption

	optlimitno = Request.Form("optlimitno") & ","
	optionyn = Request.Form("optionyn") & ","
	itemoption = Request.Form("itemoption") & ","

	arroptlimitno = split(optlimitno,",")
	arroptionyn = split(optionyn,",")
	arritemoption = split(itemoption,",")
	loopcnt = ubound(arritemoption)
'Response.write optlimitno & "<br>"
'Response.write optionyn & "<br>"
'Response.write itemoption & "<br>"
'Response.write loopcnt & "<br>"
	Dim sqlStr, i
    ''0번은 입력 없음. N까지
    For i=0 To loopcnt-1
		sqlStr = "update db_academy.dbo.tbl_diy_item_option"
		sqlStr = sqlStr + " set isusing='" + CStr(trim(arroptionyn(i))) + "'"
		sqlStr = sqlStr + " ,optlimitno='" + CStr(trim(arroptlimitno(i))) + "'"
		sqlStr = sqlStr + " where itemid=" + CStr(trim(itemid))
		sqlStr = sqlStr + " and itemoption='" + CStr(trim(arritemoption(i))) + "'"
	'Response.write sqlStr & "<br>"
		dbACADEMYget.Execute sqlStr
    Next
End Function

Dim itemid, iErrMsg, makerid
itemid = requestCheckVar(request("itemid"),10)
makerid = request.cookies("partner")("userid")

If (ItemCheckMyItemYN(makerid,itemid)<>true) Then
	Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
	Response.end
End If

iErrMsg = RegDoubleOptionProc(itemid)
if (iErrMsg="") then
%>
<script type="text/javascript">
<!--
	parent.fnOptionStateEditEnd();
//-->
</script>
<%
end if
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->