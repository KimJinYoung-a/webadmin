<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 다이어리 저장페이지
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
dim vote_num, title, question, startdate, enddate, isusing 
dim i , mode , contents , mainimage
	vote_num = request("vote_num")
	title = request("title")
	question = request("question")
	startdate = request("startdate")
	enddate = request("enddate")				
	isusing = request("isusing")
	mode = request("mode")
	contents = request("contents")
	mainimage = request("mainimg")
	
dim sql

'//신규 & 수정 
if mode = "add" then
			
	'//신규	
	if vote_num = "" then
		sql = "insert into db_momo.dbo.tbl_vote (title, question, startdate, enddate, isusing,mainimage)" + vbcrlf
		sql = sql & " values (" + vbcrlf
		sql = sql & " '"&html2db(title)&"'"
		sql = sql & " ,'"&html2db(question)&"'"				
		sql = sql & " ,'"&html2db(startdate)&" 00:00:00'"		
		sql = sql & " ,'"&html2db(enddate)&" 23:59:59'"	
		sql = sql & " ,'Y'"
		sql = sql & " ,'"&html2db(mainimage)&" 23:59:59'"											
		sql = sql & " )"		
	
	'response.write sql &"<br>"
	dbget.execute sql
		
	'//수정	
	else 
	
		if vote_num = "" then
			response.write "<script>alert('공감질문 아이디 값이 없습니다.'); self.close();</script>"
			dbget.close() : response.end
		end if		

	sql = "update db_momo.dbo.tbl_vote set" + vbcrlf	
	sql = sql & " title='"&html2db(title)&"'" + vbcrlf	
	sql = sql & " ,question='"&html2db(question)&"'" + vbcrlf		
	sql = sql & " ,startdate='"&html2db(startdate)&" 00:00:00'" + vbcrlf
	sql = sql & " ,enddate='"&html2db(enddate)&" 23:59:59'" + vbcrlf		
	sql = sql & " ,isusing='"&isusing&"'" + vbcrlf	
	sql = sql & " ,mainimage='"&mainimage&"'" + vbcrlf				
	sql = sql & " where vote_num = "&vote_num&"" + vbcrlf	
	
	'response.write sql &"<br>"
	dbget.execute sql
		
	end if			

'//투표등록
elseif mode = "contents" then
			
	if vote_num = "" then
		response.write "<script>alert('공감질문 아이디 값이 없습니다.'); self.close();</script>"
		dbget.close() : response.end
	end if		
	
	contents = contents & ","
	contents = split(contents,",")
	
	'//트랜젝션 시작
	dbget.begintrans
	
		''기존내역삭제
		sql = "update db_momo.dbo.tbl_vote_contents set isusing='N' where vote_num = "&vote_num&""
	
		'response.write sql &"<br>"
		dbget.execute sql
		
		for i = 0 to ubound(contents) -1
		
			sql = ""
			sql = "insert into db_momo.dbo.tbl_vote_contents" + vbcrlf
			sql = sql & " (vote_num,contents_num,contents,isusing) values" + vbcrlf
			sql = sql & " (" + vbcrlf
			sql = sql & " "&vote_num&"" + vbcrlf
			sql = sql & " ,"&i&"" + vbcrlf			
			sql = sql & " ,'"&html2db(contents(i))&"'" + vbcrlf			
			sql = sql & " ,'Y'" + vbcrlf	
			sql = sql & " )" + vbcrlf
						
			'response.write sql &"<br>"
			dbget.execute sql		
		next
	
	'오류검사 및 반영
	If Err.Number = 0 Then   
		dbget.CommitTrans				'커밋(정상)

	Else
	    dbget.RollBackTrans				'롤백(에러발생시)
				
		response.write "<script language='javascript'>"
		response.write "alert('정상적인 접근방식이 아니거나 오류가 발생되었습니다.');"
		response.write "self.close();"	
		response.write "</script>"
		rsget.close : resposne.end
	End If
		
end if	
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<script>
	opener.location.reload();
	alert('처리되었습니다');
	self.close();
</script>