<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 해당 아이디나 직원번호의 파트 넘버를 반환(아작스)
' History : 2016.04.07 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim empno, userid
	empno = requestcheckvar(request("empno"),10)
	userid = requestcheckvar(request("userid"),32)

if empno="" and userid="" then
	response.write "구분자가 없습니다"
	dbget.close() : response.end
end if

'//본인의 파트번호
dim part_sn
	part_sn = getpart_sn(empno, userid)

response.write part_sn
dbget.close() : response.end
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->