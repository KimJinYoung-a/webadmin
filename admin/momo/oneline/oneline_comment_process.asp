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
dim idx
	idx = request("idx")

if idx = "" then
	response.write "<script>"
	response.write "alert('IDX���� �����ϴ�. ������ �����ϼ���');"
	response.write "self.close();"
	response.write "</script>"
	rsget.close
	response.end
end if

dim sql

sql = "update db_momo.dbo.tbl_oneline_comment set" + vbcrlf
sql = sql & " isusing='N'" + vbcrlf
sql = sql & " where idx = "&idx&"" + vbcrlf	

'response.write sql &"<br>"
dbget.execute sql
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<script>
	opener.location.reload();
	alert('OK');
	self.close();
</script>