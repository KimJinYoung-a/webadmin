<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������������ �űԵ�� ��¥�� �� , �׷�� ��� �ߺ� �˻�������
' History : 2007.09.28 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<%
dim yyyy,mm,groupname , yyyymm
	yyyy = request("yyyy")
	mm = request("mm")
	groupname = request("groupname")
	yyyymm = yyyy&mm
%>

<%
dim sql,i

sql = "select groupname , yyyymm" 
sql = sql & " from db_datamart.dbo.tbl_month_wonga"
sql = sql & " where 1=1 and	groupname = '"& groupname &"' and yyyymm = '"& yyyymm &"'"
'response.write sql&"<br>"
db3_rsget.open sql,db3_dbget,1

	if not db3_rsget.eof then				'���ڵ尡 �ִٸ�

		db3_rsget.close
	%>		
		<script language="javascript">
			alert('�Է��Ͻ� ���а� ������ �⵵�� ���� ��ϵǾ� �ֽ��ϴ�. Ȯ���Ͻ��� �ٽ� �Է��ϼ���.');
			opener.location.reload();
			self.close();
		</script>
	<%
		dbget.close()	:	response.End
	else
		db3_rsget.close
	%>		
		<script language="javascript">
			alert('��밡��');
			self.close();
		</script>
	<% end if %>
		
	
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->