<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ���̵������ ����������
' Hieditor : 2009.11.18 �ѿ�� ����
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

'// ����
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
	response.write "<script>alert('�Ѱ��� ������ �ּ���'); self.close();</script>"
	rsget.close() : response.end
	end if
	
	'//���� ����Ʈ �� N ���� �ٲ�
	sql = "update db_momo.dbo.tbl_ideafile set" + vbcrlf
	sql = sql & " bestyn = 'N'" + vbcrlf
	sql = sql & " where bestyn = 'Y'" + vbcrlf	

	'response.write sql &"<br>"
	dbget.execute sql
	
	'//���õ� ���̵�������� ���� ��	
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
	alert('ó���Ǿ����ϴ�');
	self.close();
</script>