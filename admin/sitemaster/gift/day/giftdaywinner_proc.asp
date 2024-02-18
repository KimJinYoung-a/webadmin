<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트
' History : 2014.03.19 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftday_cls.asp"-->

<%
dim detailidx, mode, menupos, adminid, sqlStr
	detailidx 	= requestcheckvar(request("detailidx"),10)
	mode 	= requestcheckvar(request("mode"),32)
	menupos 	= requestcheckvar(request("menupos"),10)

adminid = session("ssBctId")

dim refer
	refer = request.ServerVariables("HTTP_REFERER")
if InStr(refer,"10x10.co.kr")<1 then
	Response.Write "잘못된 접속입니다."
	dbget.close()	:	response.End
end if

If mode = "del" Then
	if detailidx="" then
		Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
		dbget.close()	:	response.End			
	end if

	sqlStr = "UPDATE db_board.dbo.tbl_giftday_detail" + VBCRLF
	sqlStr = sqlStr & " SET isusing = 'N' where " + VBCRLF
	sqlStr = sqlStr & " detailidx ='" & Cstr(detailidx) & "'"

	'response.write sqlStr & "<BR>"	
	dbget.execute sqlStr

	Response.Write "<script language='javascript'>alert('OK'); location.href='"&refer&"'</script>"

else
	Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End
End If
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->