<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����� ����Ʈ������ �⺻ ���а� 
' Hieditor : 2017.11.02 ���¿� ����
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
						"alert('�߸��� �����Դϴ�.');" &_
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
	alert("����Ǿ����ϴ�.");
	self.location = "index.asp?menupos=<%=menupos%>";
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
