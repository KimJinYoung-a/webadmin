<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 모바일 베스트페이지 기본 구분값 
' Hieditor : 2017.11.02 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->

<%
dim menupos : menupos	= Request("menupos")
dim mode : mode = requestCheckvar(Request("mode"),15)
dim bestgubun : bestgubun = requestCheckvar(Request("bestgubun"),2)

	if bestgubun="" or isNull(bestgubun) then
		bestgubun = "dt"
	end if

	if mode <> "bestgubunupdate" then
		Response.Write	"<script type='text/javascript'>" &_
						"alert('잘못된 접속입니다.');" &_
						"</script>"
		dbget.close()	:	response.end
	end if
	
	dim sqlstr
	if mode = "bestgubunupdate" then
		sqlstr = " update db_sitemaster.dbo.tbl_mobile_best_gubun set "
		sqlstr = sqlstr & " bestgubun = '"& bestgubun &"' "
		'response.write sqlstr
		dbget.execute sqlstr
	end if
%>

<script language = "javascript">
	alert("저장되었습니다.");
	self.location = "index.asp?menupos=<%=menupos%>";
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
