<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/lectureadmin/lib/email/smsLib.asp"-->
<!-- #include virtual="/lectureadmin/lib/email/maillib.asp"-->
<!-- #include virtual="/lectureadmin/lib/email/mailLib2.asp"-->
<!-- #include virtual="/lectureadmin/lib/email/mailFunc_Designer.asp"-->
<%
dim id,finishmemo, finishuser,songjangdiv, songjangno, MakerID
id          = requestCheckVar(request("id"),10)
finishmemo  = html2db(request("finishmemo"))
finishuser  = requestCheckVar(request("finishuser"),32)
songjangdiv = requestCheckVar(request("songjangdiv"),2)
songjangno  = requestCheckVar(request("songjangno"),16)
MakerID		= requestCheckVar(request.cookies("partner")("userid"),32)

dim sqlStr, GetOrderStateNum
dim oldcurrstate, msg
sqlStr = "select currstate "
sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_as_list" + VbCrlf
sqlStr = sqlStr + " where id =" + id
rsACADEMYget.Open sqlStr,dbACADEMYget,1
    oldcurrstate = rsACADEMYget("currstate")
rsACADEMYget.Close

if (oldcurrstate="B007") then
    msg="이미 처리 완료된 내역입니다. - 완료처리로 진행 할 수 없습니다."
Else
	sqlStr = "update [db_academy].[dbo].tbl_academy_as_list" + VbCrlf
	sqlStr = sqlStr + " set finishuser ='" + finishuser + "'," + VbCrlf
	sqlStr = sqlStr + " contents_finish ='" + finishmemo + "'," + VbCrlf
	sqlStr = sqlStr + " songjangdiv ='" + songjangdiv + "'," + VbCrlf
	sqlStr = sqlStr + " songjangno ='" + songjangno + "'," + VbCrlf
	sqlStr = sqlStr + " finishdate=getdate()," + VbCrlf
	sqlStr = sqlStr + " currstate='B006'" + VbCrlf
	sqlStr = sqlStr + " where id =" + id
	sqlStr = sqlStr + " and makerid='" & MakerID & "'"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	msg="CS 처리결과 작성이 완료 되었습니다."

	sqlStr = sqlStr + "update [db_academy].[dbo].[tbl_academy_app_iconbadge_count]" + vbCrlf
	sqlStr = sqlStr + "	set cscnt=cscnt-1" + vbCrlf
	sqlStr = sqlStr + "	where makerid='" + CStr(MakerID) + "'" + vbCrlf
	dbACADEMYget.Execute sqlStr

	sqlStr = "select mibaljucnt, ordercnt, cscnt from [db_academy].[dbo].[tbl_academy_app_iconbadge_count] where makerid='" + CStr(MakerID) + "'" + vbCrlf
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	if not rsACADEMYget.EOF then
		GetOrderStateNum=rsACADEMYget("mibaljucnt")+rsACADEMYget("ordercnt")+rsACADEMYget("cscnt")
	Else
		GetOrderStateNum=0
	End If
end if

%>
<script>
<!--
parent.fnCSInputEnd("<%=msg%>",<%=GetOrderStateNum%>);
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->