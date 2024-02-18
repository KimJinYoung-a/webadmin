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

<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/itemOptionLib.asp"-->
<%

function WaitRegDoubleOptionProc(waititemid)
    
    Dim optlimitno, optionyn, itemoption, loopcnt
	Dim arroptlimitno,arroptionyn,arritemoption

	optlimitno = Request.Form("optlimitno") & ","
	optionyn = Request.Form("optionyn") & ","
	itemoption = Request.Form("itemoption") & ","

	arroptlimitno = split(optlimitno,",")
	arroptionyn = split(optionyn,",")
	arritemoption = split(itemoption,",")
	loopcnt = ubound(arritemoption)
 
 
    ''0번은 입력 없음. N까지
    For i=0 To loopcnt-1
		sqlStr = "update db_academy.dbo.tbl_diy_wait_item_option"
		sqlStr = sqlStr + " set isusing='" + CStr(arroptionyn(i)) + "'"
		sqlStr = sqlStr + " ,optlimitno='" + CStr(arroptlimitno(i)) + "'"
		sqlStr = sqlStr + " where itemid=" + CStr(waititemid(i))
		sqlStr = sqlStr + " and itemoption='" + CStr(arritemoption(i)) + "'"
		dbACADEMYget.Execute sqlStr
    Next
End Function

Dim waititemid, iErrMsg
waititemid = requestCheckVar(request("waititemid"),10)

iErrMsg = WaitRegDoubleOptionProc(waititemid)
if (iErrMsg<>"") then
%>
<script type="text/javascript">
<!--
	
//-->
</script>
<%
end if
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->