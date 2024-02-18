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
dim jukyocd,jukyoname,isusing
	jukyocd = request("jukyocd")
	jukyoname = request("jukyoname")
	isusing = request("isusing")
	if isusing = "" then
		isusing = "Y"
	end if	
%>

<% 
dim sql_seach
	sql_seach = "select jukyocd from db_user.dbo.tbl_mileage_gubun"
	sql_seach = sql_seach & " where 1=1 and jukyocd = '" & jukyocd & "'"
	
	rsget.open sql_seach,dbget,1
	'response.write sql_seach&"<br>"
	if not rsget.eof then
%>

<script language="javascript">
	alert('코드값이 이미 등록되어 있습니다.');
	history.go(-1);
</script>
			
<% else %>	

<%
dim sql
	sql = "insert into db_user.dbo.tbl_mileage_gubun (jukyocd,jukyoname,isusing) values"
	sql = sql & " ('" & jukyocd & "',"
	sql = sql & "'" & jukyoname & "'," 
	sql = sql & "'" & isusing & "')"

	dbget.execute sql
	'response.write sql&"<br>"
%>
	<script language="javascript">		
		opener.location.reload();
		self.close();
	</script>
<% 
end if
rsget.close
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

