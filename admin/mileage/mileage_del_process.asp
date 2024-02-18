<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  마일리지 구분 
' History : 2007.10.23 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/mileage/mileage_class.asp"-->
<%
dim jukyocd
	jukyocd = request("jukyocd")
	
dim sql
	sql = "delete from db_user.dbo.tbl_mileage_gubun where 1=1 and jukyocd = '" & jukyocd & "'"
	dbget.execute sql
	'response.write sql&"<br>"
	
%>
<script language="javascript">
	opener.location.reload();
	self.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->

