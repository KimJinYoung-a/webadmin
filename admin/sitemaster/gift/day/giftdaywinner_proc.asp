<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����Ʈ
' History : 2014.03.19 �ѿ�� ����
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
	Response.Write "�߸��� �����Դϴ�."
	dbget.close()	:	response.End
end if

If mode = "del" Then
	if detailidx="" then
		Response.Write "<script language='javascript'>alert('�����ڰ� �����ϴ�.'); history.back(-1);</script>"
		dbget.close()	:	response.End			
	end if

	sqlStr = "UPDATE db_board.dbo.tbl_giftday_detail" + VBCRLF
	sqlStr = sqlStr & " SET isusing = 'N' where " + VBCRLF
	sqlStr = sqlStr & " detailidx ='" & Cstr(detailidx) & "'"

	'response.write sqlStr & "<BR>"	
	dbget.execute sqlStr

	Response.Write "<script language='javascript'>alert('OK'); location.href='"&refer&"'</script>"

else
	Response.Write "<script language='javascript'>alert('�����ڰ� �����ϴ�.'); history.back(-1);</script>"
	dbget.close()	:	response.End
End If
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->