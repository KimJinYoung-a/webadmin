<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ���̾ ����������
' Hieditor : 2009.12.01 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim idx, diary_date, title, contents, mainimage1, isusing
dim odiary ,iResult , arrlist ,ocheck , mode , diary_order
dim mainimage2 , mainimage3 , sql , diarytype
	idx   = request("idx")
	diary_date   = request("diary_date")
	title   = request("title")
	contents   = request("contents")
	mainimage1   = request("mainimage1")
	mainimage2   = request("mainimage2")
	mainimage3   = request("mainimage3")	
	isusing   = request("isusing")
	mode   = request("mode")
	diarytype   = request("diarytype")	
	diary_order = 50
	
	'//������ ��Ϻκ� 	
	if mode = "contents" then	
		
		'// �űԵ��
		if idx = "" then
			
			'//�ߺ� üũ
			sql = "select top 1 idx ,diary_date" + vbcrlf
			sql = sql & " from db_momo.dbo.tbl_diary" + vbcrlf
			sql = sql & " where convert(varchar(10),diary_date,121) = '"&diary_date&"'"			
			
			'response.write sql &"<Br>"
			rsget.open sql,dbget,1
			
			if not(rsget.bof or rsget.eof) then
			%>
			
			<script>
				alert('�̹� ��ϵ� ��¥ �Դϴ�');
				history.go(-1);
			</script>	
			
			<% response.end
			end if	
			
			sql = ""			
			sql = "insert into db_momo.dbo.tbl_diary" + vbcrlf
			sql = sql & " (diary_date,title,contents,isusing,diary_order,diarytype)" + vbcrlf
			sql = sql & " values(" + vbcrlf
			sql = sql & " '"&html2db(diary_date)&"'" + vbcrlf
			sql = sql & " ,'"&html2db(title)&"'" + vbcrlf
			sql = sql & " ,'"&html2db(contents)&"'" + vbcrlf
			sql = sql & " ,'"&isusing&"'" + vbcrlf
			sql = sql & " ,"&diary_order&"" + vbcrlf
			sql = sql & " ,'"&diarytype&"'" + vbcrlf
			sql = sql & " )"
			
			'response.write sql &"<Br>"					
			dbget.execute sql
			
		else
		
			sql = ""
			sql = "update db_momo.dbo.tbl_diary set " + vbcrlf
			sql = sql & " diary_date = '"&html2db(diary_date)&"'" + vbcrlf
			sql = sql & " ,title= '"&html2db(title)&"'" + vbcrlf
			sql = sql & " ,contents='"&html2db(contents)&"'" + vbcrlf
			sql = sql & " ,isusing='"&isusing&"'" + vbcrlf
			sql = sql & " ,diary_order = "&diary_order&"" + vbcrlf
			sql = sql & " ,diarytype = '"&diarytype&"'" + vbcrlf			
			sql = sql & " where idx= "&idx&""

			'response.write sql &"<Br>"					
			dbget.execute sql
					
		end if		
	%>
		<script>			
			opener.location.reload();
			self.close();
		</script>
		
<%
	'//�̹��� ó���κ�
	elseif mode = "image" then	

		if idx = "" then
%>		
			<script>
				alert('�ε��� ���� �����ϴ�. ������ ���ǿ��');
				self.close();
			</script>			
<%
		end if				
	
			sql = "update db_momo.dbo.tbl_diary set " + vbcrlf
			sql = sql & " mainimage1= '"&html2db(mainimage1)&"'" + vbcrlf
			sql = sql & " ,mainimage2='"&html2db(mainimage2)&"' " + vbcrlf
			sql = sql & " ,mainimage3='"&html2db(mainimage3)&"' " + vbcrlf
			sql = sql & " where idx= "&idx&" " + vbcrlf

			'response.write sql &"<Br>"					
			dbget.execute sql
%>		
		<script>		
			opener.location.reload();
			self.close();
		</script>
<%
end if
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

