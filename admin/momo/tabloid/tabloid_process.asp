<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� Ÿ����̵�
' Hieditor : 2009.11.17 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim tabloid, keyword, mainimage, regdate, isusing , detailimage
dim wordimage , ingimage , mode ,wordovimage
	tabloid = request("tabloid")
	keyword = request("keyword")
	mainimage = request("mainimg")
	isusing = request("isusing")
	detailimage = request("detailimg")				
	wordimage = request("wordimg")
	wordovimage = request("wordovimg")	
	ingimage = request("ingimg")	
	mode = request("mode")
dim sql

'// ����
if mode = "delete" then
	
	tabloid = left(tabloid,len(tabloid)-1)
	
	'//Ʈ������ 
	dbget.begintrans
	
	sql = "update db_momo.dbo.tbl_tabloid set" + vbcrlf	
	sql = sql & " isusing='N'" + vbcrlf
	sql = sql & " where tabloid in("&tabloid&")" + vbcrlf	
	
	'response.write sql &"<br>"
	dbget.execute sql
	
	sql = ""	
	sql = "update db_momo.dbo.tbl_tabloid_item set" + vbcrlf	
	sql = sql & " isusing='N'" + vbcrlf
	sql = sql & " where tabloid in ("&tabloid&")" + vbcrlf	
	
	'response.write sql &"<br>"
	dbget.execute sql
	
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
	
elseif mode = "ing" then	

	tabloid = split(tabloid,",")

	if ubound(tabloid) <> "1" then
	response.write "<script>alert('�Ѱ��� ������ �ּ���'); self.close();</script>"
	rsget.close() : response.end
	end if
		
	sql = "update db_momo.dbo.tbl_tabloid set" + vbcrlf
	sql = sql & " best = best + 50" + vbcrlf
	sql = sql & " where tabloid = "&tabloid(0)&"" + vbcrlf	
	
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