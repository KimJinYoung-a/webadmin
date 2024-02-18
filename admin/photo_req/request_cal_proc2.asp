<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim rno, sno, status, query1
Dim scnt, scnt2
rno 	= request("rno")
sno 	= request("sno")
status  = request("req_status")

	query1 = " Update db_partner.dbo.tbl_photo_schedule set "
	query1 = query1 + " status = '"&status&"'"
	query1 = query1 + " where schedule_no = '"&sno&"'"
	dbget.execute query1

	query1 = " select count(*) as cnt from db_partner.dbo.tbl_photo_schedule "
	query1 = query1 + " where req_no = '"&rno&"'"
	rsget.Open query1,dbget,1
	IF not rsget.EOF THEN
		scnt = rsget("cnt")
	End IF
	rsget.Close

	query1 = " select count(*) as cnt from db_partner.dbo.tbl_photo_schedule "
	query1 = query1 + " where req_no = '"&rno&"' and status = '"&status&"' "
	rsget.Open query1,dbget,1
	IF not rsget.EOF THEN
		scnt2 = rsget("cnt")
	End IF
	rsget.Close

	If CInt(scnt) = CInt(scnt2) Then
		query1 = " Update db_partner.dbo.tbl_photo_req set "
		query1 = query1 + " req_status = '"&status&"'"
		query1 = query1 + " where req_no = '"&rno&"'"
		dbget.execute query1		
	End If
%>
<script language="javascript">
	alert('OK');
	window.opener.location.reload();
	self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->