<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ����������
' Hieditor : 2009.10.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim photoid, photoword, mainimage, regdate, isusing , detailimage
dim wordimage , ingimage , mode ,wordovimage
	photoid = request("photoid")
	photoword = request("photoword")
	mainimage = request("mainimg")
	isusing = request("isusing")
	detailimage = request("detailimg")				
	wordimage = request("wordimg")
	wordovimage = request("wordovimg")	
	ingimage = request("ingimg")	
	mode = request("mode")
dim sql

if mode = "edit" then
	
	''�űԵ��
	if photoid = "" then
		
		sql = "insert into db_momo.dbo.tbl_photo (photoword,mainimage,detailimage,wordimage,ingimage,wordovimage,isusing) values " + vbcrlf
		sql = sql & "( " + vbcrlf
		sql = sql & " '"&html2db(photoword)&"' " + vbcrlf
		sql = sql & " ,'"&html2db(mainimage)&"' " + vbcrlf
		sql = sql & " ,'"&html2db(detailimage)&"' " + vbcrlf
		sql = sql & " ,'"&html2db(wordimage)&"' " + vbcrlf
		sql = sql & " ,'"&html2db(ingimage)&"' " + vbcrlf
		sql = sql & " ,'"&html2db(wordovimage)&"' " + vbcrlf			
		sql = sql & " ,'"&isusing&"' " + vbcrlf	
		sql = sql & ") "
		
		'response.write sql &"<br>"
		dbget.execute sql
		
	'����	
	else
	
		sql = "update db_momo.dbo.tbl_photo set" + vbcrlf
		sql = sql & " photoword='"&html2db(photoword)&"'" + vbcrlf
		sql = sql & " ,mainimage='"&html2db(mainimage)&"'" + vbcrlf
		sql = sql & " ,detailimage='"&html2db(detailimage)&"'" + vbcrlf	
		sql = sql & " ,wordimage='"&html2db(wordimage)&"'" + vbcrlf
		sql = sql & " ,ingimage='"&html2db(ingimage)&"'" + vbcrlf
		sql = sql & " ,wordovimage='"&html2db(wordovimage)&"'" + vbcrlf				
		sql = sql & " ,isusing='"&isusing&"'" + vbcrlf
		sql = sql & " where photoid = "&photoid&"" + vbcrlf	
		
		'response.write sql &"<br>"
		dbget.execute sql
	end if	

elseif mode = "ing" then	

	photoid = split(photoid,",")

	if ubound(photoid) <> "1" then
	response.write "<script>alert('�Ѱ��� ������ �ּ���'); self.close();</script>"
	rsget.close() : response.end
	end if
		
	sql = "update db_momo.dbo.tbl_photo set" + vbcrlf
	sql = sql & " regdate=getdate()" + vbcrlf
	sql = sql & " where photoid = "&photoid(0)&"" + vbcrlf	
	
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