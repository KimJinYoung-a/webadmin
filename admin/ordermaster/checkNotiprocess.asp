<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/cjmall/cjmallitemcls.asp"-->
<!-- #include virtual="/admin/etc/cjmall/incCJmallFunction.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),30)
Dim cksel : cksel = request("cksel")
Dim sqlStr, AssignedRow

If (cmdparam="I") Then	'�ϰ� Ȯ��ó��
	cksel = Trim(cksel)
	if Right(cksel,1)="," then cksel=Left(cksel,Len(cksel)-1)
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE [db_contents].[dbo].[tbl_check_noti_log] SET " & VBCRLF
	sqlStr = sqlStr & " isconfirmed = '1' " & VBCRLF
	sqlStr = sqlStr & " ,lastconfirmDT = getdate() " & VBCRLF
	sqlStr = sqlStr & " ,lastconfirmUSER = '"&session("ssBctID")&"' " & VBCRLF
	sqlStr = sqlStr & " WHERE itemid in ("&cksel&")"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"�� �ϰ� Ȯ���Ͽ����ϴ�.');parent.location.reload();</script>"
End if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->