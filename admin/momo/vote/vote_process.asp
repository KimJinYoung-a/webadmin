<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ���̾ ����������
' Hieditor : 2009.11.20 �ѿ�� ����
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

'//�ű� & ���� 
if mode = "add" then
			
	'//�ű�	
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
		
	'//����	
	else 
	
		if vote_num = "" then
			response.write "<script>alert('�������� ���̵� ���� �����ϴ�.'); self.close();</script>"
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

'//��ǥ���
elseif mode = "contents" then
			
	if vote_num = "" then
		response.write "<script>alert('�������� ���̵� ���� �����ϴ�.'); self.close();</script>"
		dbget.close() : response.end
	end if		
	
	contents = contents & ","
	contents = split(contents,",")
	
	'//Ʈ������ ����
	dbget.begintrans
	
		''������������
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
	
	'�����˻� �� �ݿ�
	If Err.Number = 0 Then   
		dbget.CommitTrans				'Ŀ��(����)

	Else
	    dbget.RollBackTrans				'�ѹ�(�����߻���)
				
		response.write "<script language='javascript'>"
		response.write "alert('�������� ���ٹ���� �ƴϰų� ������ �߻��Ǿ����ϴ�.');"
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
	alert('ó���Ǿ����ϴ�');
	self.close();
</script>