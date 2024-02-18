<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 한줄소설 저장페이지
' Hieditor : 2009.11.17 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim novelid,startdate,enddate,title,prolog,genre,isusing , i , mode
dim wordimage , winner
	novelid = request("novelid")
	startdate = request("startdate")
	enddate = request("enddate")
	title = request("title")
	genre = request("genre")				
	prolog = request("prolog")
	isusing = request("isusing")	
	mode = request("mode")
	wordimage = request("wordimg")
	winner = request("winner")
dim sql

'//신규 & 수정 
if mode = "edit" then
			
	'//신규	
	if novelid = "" then
		sql = "insert into db_momo.dbo.tbl_novel (startdate,enddate,prolog,title,genre,isusing,winner,wordimage)" + vbcrlf
		sql = sql & " values (" + vbcrlf
		sql = sql & " '"&html2db(startdate)&" 00:00:00'"		
		sql = sql & " ,'"&html2db(enddate)&" 23:59:59'"	
		sql = sql & " ,'"&html2db(prolog)&"'"	
		sql = sql & " ,'"&html2db(title)&"'"
		sql = sql & " ,'"&html2db(genre)&"'"		
		sql = sql & " ,'Y'"		
		sql = sql & " ,'"&winner&"'"		
		sql = sql & " ,'"&html2db(wordimage)&"'"							
		sql = sql & " )"		
	
	'response.write sql &"<br>"
	dbget.execute sql
		
	'//수정	
	else 
	
	if novelid = "" then
		response.write "<script>alert('한줄소설 아이디 값이 없습니다.'); self.close();</script>"
		dbget.close() : response.end
	end if		

	sql = "update db_momo.dbo.tbl_novel set" + vbcrlf	
	sql = sql & " startdate='"&html2db(startdate)&" 00:00:00'" + vbcrlf
	sql = sql & " ,enddate='"&html2db(enddate)&" 23:59:59'" + vbcrlf
	sql = sql & " ,prolog='"&html2db(prolog)&"'" + vbcrlf
	sql = sql & " ,title='"&html2db(title)&"'" + vbcrlf
	sql = sql & " ,genre='"&html2db(genre)&"'" + vbcrlf				
	sql = sql & " ,isusing='"&isusing&"'" + vbcrlf
	sql = sql & " ,winner='"&winner&"'" + vbcrlf			
	sql = sql & " ,wordimage='"&html2db(wordimage)&"'" + vbcrlf				
	sql = sql & " where novelid = "&novelid&"" + vbcrlf	
	
	'response.write sql &"<br>"
	dbget.execute sql
		
	end if			
end if	
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<script>
	opener.location.reload();
	alert('처리되었습니다');
	self.close();
</script>