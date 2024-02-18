<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 아이디어파일 저장페이지
' Hieditor : 2009.11.18 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim ideafileid, keyword, mainimage, regdate, isusing , detailimage
dim wordimage , ingimage , mode ,wordovimage
	ideafileid = request("ideafileid")
	mode = request("mode")
dim sql

'// 삭제
if mode = "delete" then
	
	ideafileid = left(ideafileid,len(ideafileid)-1)
	
	sql = "update db_momo.dbo.tbl_ideafile set" + vbcrlf	
	sql = sql & " isusing='N'" + vbcrlf
	sql = sql & " where ideafileid in("&ideafileid&")" + vbcrlf	
	
	'response.write sql &"<br>"
	dbget.execute sql
			
elseif mode = "ing" then	

	ideafileid = split(ideafileid,",")

	if ubound(ideafileid) <> "1" then
	response.write "<script>alert('한개만 선택해 주세요'); self.close();</script>"
	rsget.close() : response.end
	end if
	
	'//기존 베스트 를 N 으로 바꿈
	sql = "update db_momo.dbo.tbl_ideafile set" + vbcrlf
	sql = sql & " bestyn = 'N'" + vbcrlf
	sql = sql & " where bestyn = 'Y'" + vbcrlf	

	'response.write sql &"<br>"
	dbget.execute sql
	
	'//선택된 아이디어파일을 선정 함	
	sql = ""	
	sql = "update db_momo.dbo.tbl_ideafile set" + vbcrlf
	sql = sql & " bestyn = 'Y'" + vbcrlf
	sql = sql & " where ideafileid = "&ideafileid(0)&"" + vbcrlf	
	
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