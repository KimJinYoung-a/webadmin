<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ���ٳ���
' Hieditor : 2010.11.23 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim onelineid, startdate,enddate,winnerdate,comment,isusing , i , mode , winner,winnercomment
	onelineid = request("onelineid")
	startdate = request("startdate")
	enddate = request("enddate")
	winnerdate = request("winnerdate")
	comment = request("comment")				
	isusing = request("isusing")	
	mode = request("mode")
	winner = request("winner")	
	winnercomment = request("winnercomment")
	
dim sql

'//�ű� & ���� 
if mode = "edit" then
			
	'//�ű�	
	if onelineid = "" then
		sql = "insert into db_momo.dbo.tbl_oneline (startdate,enddate,winnerdate,comment,isusing,winner,winnercomment)" + vbcrlf
		sql = sql & " values (" + vbcrlf
		sql = sql & " '"&html2db(startdate)&" 00:00:00'"		
		sql = sql & " ,'"&html2db(enddate)&" 23:59:59'"	
		sql = sql & " ,'"&html2db(winnerdate)&" 00:00:00'"	
		sql = sql & " ,'"&html2db(comment)&"'"
		sql = sql & " ,'"&isusing&"'"
		sql = sql & " ,'"&html2db(winner)&"'"
		sql = sql & " ,'"&html2db(winnercomment)&"'"				
		sql = sql & " )"		
	
	'response.write sql &"<br>"
	dbget.execute sql
		
	'//����	
	else 
	
	if onelineid = "" then
		response.write "<script>alert('���̵� ���� �����ϴ�.'); self.close();</script>"
		dbget.close() : response.end
	end if		

	sql = "update db_momo.dbo.tbl_oneline set" + vbcrlf	
	sql = sql & " startdate='"&html2db(startdate)&" 00:00:00'" + vbcrlf
	sql = sql & " ,enddate='"&html2db(enddate)&" 23:59:59'" + vbcrlf
	sql = sql & " ,winnerdate='"&html2db(winnerdate)&"'" + vbcrlf
	sql = sql & " ,comment='"&html2db(comment)&"'" + vbcrlf			
	sql = sql & " ,isusing='"&isusing&"'" + vbcrlf
	sql = sql & " ,winner='"&html2db(winner)&"'" + vbcrlf
	sql = sql & " ,winnercomment='"&html2db(winnercomment)&"'" + vbcrlf					
	sql = sql & " where onelineid = "&onelineid&"" + vbcrlf	
	
	'response.write sql &"<br>"
	dbget.execute sql
		
	end if			
end if	
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<script>
	opener.location.reload();
	alert('ó���Ǿ����ϴ�');
	self.close();
</script>