<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  마일리지 수정 저장
' History : 2007.10.23 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/mileage/mileage_class.asp"-->
<%
dim jukyocd,jukyoname,isusing,realjukyocd
	jukyocd = request("jukyocd")
	jukyoname = request("jukyoname")
	isusing = request("isusing")
	realjukyocd = request("realjukyocd")
%>

<% 
dim sql
	sql = "update db_user.dbo.tbl_mileage_gubun set"
	sql = sql & " jukyocd= '" & jukyocd & "',"
	sql = sql & " jukyoname= '" & jukyoname & "',"
	sql = sql & " isusing= '" & isusing & "'"
	sql = sql & " where jukyocd = '" & realjukyocd & "'"
	
	dbget.execute sql
	response.write sql&"<br>"
%>	
	
	<script language="javascript">		
		opener.location.reload();
		self.close();
	</script>


<!-- #include virtual="/lib/db/dbclose.asp" -->

