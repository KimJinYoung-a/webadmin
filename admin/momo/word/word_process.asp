<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 감성사전 저장페이지
' Hieditor : 2009.10.28 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim keyid, keyword, mainimage, regdate, isusing , detailimage
dim wordimage , ingimage , mode ,wordovimage , winner, prizedate
	keyid = request("keyid")
	keyword = request("keyword")
	mainimage = request("mainimg")
	isusing = request("isusing")
	detailimage = request("detailimg")				
	wordimage = request("wordimg")
	wordovimage = request("wordovimg")	
	ingimage = request("ingimg")	
	mode = request("mode")
	winner = request("winner")
	prizedate = request("prizedate")
dim sql

if mode = "edit" then
	''신규등록
	if keyid = "" then
		
		sql = "insert into db_momo.dbo.tbl_word (keyword,mainimage,detailimage,wordimage,ingimage,wordovimage,isusing , winner, prizedate) values " + vbcrlf
		sql = sql & "( " + vbcrlf
		sql = sql & " '"&html2db(keyword)&"' " + vbcrlf
		sql = sql & " ,'"&html2db(mainimage)&"' " + vbcrlf
		sql = sql & " ,'"&html2db(detailimage)&"' " + vbcrlf
		sql = sql & " ,'"&html2db(wordimage)&"' " + vbcrlf
		sql = sql & " ,'"&html2db(ingimage)&"' " + vbcrlf
		sql = sql & " ,'"&html2db(wordovimage)&"' " + vbcrlf		
		sql = sql & " ,'"&isusing&"' " + vbcrlf
		sql = sql & " ,'"&html2db(winner)&"' " + vbcrlf
		sql = sql & " ,'"&prizedate&"' " + vbcrlf
		sql = sql & ") "
		
		'response.write sql &"<br>"
		dbget.execute sql
		
	'수정	
	else
	
		sql = "update db_momo.dbo.tbl_word set" + vbcrlf
		sql = sql & " keyword='"&html2db(keyword)&"'" + vbcrlf
		sql = sql & " ,mainimage='"&html2db(mainimage)&"'" + vbcrlf
		sql = sql & " ,detailimage='"&html2db(detailimage)&"'" + vbcrlf		
		sql = sql & " ,wordimage='"&html2db(wordimage)&"'" + vbcrlf
		sql = sql & " ,ingimage='"&html2db(ingimage)&"'" + vbcrlf	
		sql = sql & " ,wordovimage='"&html2db(wordovimage)&"'" + vbcrlf		
		sql = sql & " ,isusing='"&isusing&"'" + vbcrlf
		sql = sql & " ,winner='"&html2db(winner)&"'" + vbcrlf
		sql = sql & " ,prizedate='"&prizedate&"'" + vbcrlf
		sql = sql & " where keyid = "&keyid&"" + vbcrlf	
		
		'response.write sql &"<br>"
		dbget.execute sql
	end if	
	
elseif mode = "ing" then	

	keyid = split(keyid,",")

	if ubound(keyid) <> "1" then
	response.write "<script>alert('한개만 선택해 주세요'); self.close();</script>"
	rsget.close() : response.end
	end if
		
	sql = "update db_momo.dbo.tbl_word set" + vbcrlf
	sql = sql & " regdate=getdate()" + vbcrlf
	sql = sql & " where keyid = "&keyid(0)&"" + vbcrlf	
	
	'response.write sql &"<br>"
	dbget.execute sql	
end if	
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<script>
	opener.location.reload();
	self.close();
</script>