<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 북마크
' Hieditor : 2009.11.20 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim bookmarkid, mode 
	bookmarkid = request("bookmarkid")
	mode = request("mode")
dim sql

'// 삭제
if mode = "delete" then
	
	bookmarkid = left(bookmarkid,len(bookmarkid)-1)
	
	sql = "update db_momo.dbo.tbl_bookmark set" + vbcrlf	
	sql = sql & " isusing='N'" + vbcrlf
	sql = sql & " where bookmarkid in("&bookmarkid&")" + vbcrlf	
	
	'response.write sql &"<br>"
	dbget.execute sql
			
end if	
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<script>
	opener.location.reload();
	alert('처리되었습니다');
	self.close();
</script>