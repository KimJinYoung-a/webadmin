<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� yesno ����������
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
dim yesnoid, yesnoword, mainimage, regdate, isusing , detailimage
dim wordimage , ingimage , mode ,wordovimage
	yesnoid = request("yesnoid")
	yesnoword = request("yesnoword")
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
	if yesnoid = "" then
		
		sql = "insert into db_momo.dbo.tbl_yesno (yesnoword,mainimage,detailimage,wordimage,ingimage,wordovimage,isusing) values " + vbcrlf
		sql = sql & "( " + vbcrlf
		sql = sql & " '"&html2db(yesnoword)&"' " + vbcrlf
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
	
		sql = "update db_momo.dbo.tbl_yesno set" + vbcrlf
		sql = sql & " yesnoword='"&html2db(yesnoword)&"'" + vbcrlf
		sql = sql & " ,mainimage='"&html2db(mainimage)&"'" + vbcrlf
		sql = sql & " ,detailimage='"&html2db(detailimage)&"'" + vbcrlf		
		sql = sql & " ,wordimage='"&html2db(wordimage)&"'" + vbcrlf
		sql = sql & " ,ingimage='"&html2db(ingimage)&"'" + vbcrlf	
		sql = sql & " ,wordovimage='"&html2db(wordovimage)&"'" + vbcrlf					
		sql = sql & " ,isusing='"&isusing&"'" + vbcrlf
		sql = sql & " where yesnoid = "&yesnoid&"" + vbcrlf	
		
		'response.write sql &"<br>"
		dbget.execute sql
	end if	

elseif mode = "ing" then	

	yesnoid = split(yesnoid,",")

	if ubound(yesnoid) <> "1" then
	response.write "<script>alert('�Ѱ��� ������ �ּ���'); self.close();</script>"
	rsget.close() : response.end
	end if
		
	sql = "update db_momo.dbo.tbl_yesno set" + vbcrlf
	sql = sql & " regdate=getdate()" + vbcrlf
	sql = sql & " where yesnoid = "&yesnoid(0)&"" + vbcrlf	
	
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