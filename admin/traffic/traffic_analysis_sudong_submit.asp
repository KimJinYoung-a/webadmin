<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 traffic analysis  수동 저장 페이지
' History : 2007.09.04 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/traffic/traffic_class.asp"-->

<%
dim  yyyymmdd,pageview,totalcount,newcount,recount,realcount
dim sql , sqlselect,rectyyyymmdd , i
yyyymmdd = request("yyyymmdd")
pageview = request("pageview")
totalcount = request("totalcount")
newcount = request("newcount")
recount = request("recount")
realcount = request("realcount")

sqlselect = "select yyyymmdd from db_datamart.dbo.tbl_traffic_analysis where yyyymmdd = '"& yyyymmdd &"'"
	'response.write sqlselect     '삑살시 뿌려본다.
	rsget.open sqlselect,dbget,1
		if not rsget.eof then						'레코드의 첫번째가 아니라면
			do until rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				rectyyyymmdd = rsget("yyyymmdd")
				rsget.movenext
			loop		
		end if
	rsget.close

if rectyyyymmdd = "" then
	sql = "insert into db_datamart.dbo.tbl_traffic_analysis (yyyymmdd,pageview,totalcount,newcount,recount,realcount) values"	& VbCrlf
	sql = sql & " ("& yyyymmdd &","& pageview &","& totalcount &","& newcount &","& recount &","& realcount &")" 	
	'response.write sql     '삑살시 뿌려본다.
	dbget.execute sql
end if 
%>
	<script language="javascript">
	opener.location.reload();
	self.close();
	</script>

