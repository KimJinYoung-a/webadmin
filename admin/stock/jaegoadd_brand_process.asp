<%@ language = vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�庰 ��Ը� ����
' History : 2007.07.31 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/jaegostock.asp"-->

<%
dim drawitemid , drawitemid_re
drawitemid = request("drawitemid")			'������ ��ǰid�� ���� �޾ƿ´�.
drawitemid_re = split(drawitemid,",")		' �޸��� �������� 2�����迭�� �����Ѵ�.
dim lens
lens = UBound(drawitemid_re)		'�迭�� ���� üũ�Ѵ�.

dim fnow, fmode , order		'��������
dim itemgubun,jisiid,stats			'��������
	fmode = html2db(request("mode")	)						'��屸��
	jisiid = html2db(session("ssBctId"))					'������ ���� ���� id�� �޾ƿ´�.
	order = now()											'�۾�������
	itemgubun = "10"				'��ǰ��,��������
	stats = 1												'���� �⺻�� 1

%>


<%
dim oip,sql, i,sql12
dim a
for a = 0 to lens -1 		
%>	
	<% 
	sql12 = "select * from [db_summary].[dbo].tbl_req_realstock" 
	sql12 = sql12 & " where itemid = '"& drawitemid_re(a) &"' order by statecd asc"
	rsget.open sql12,dbget,1
		
	if not rsget.eof then				'���ڵ尡 �ִٸ�
		if rsget("statecd") = 1 then	'��ǰ���� ���°��� �۾�������(1) �̶��	
			rsget.close
		%>		
		<script language="javascript">
			alert('��ǰ��ȣ(<%=drawitemid_re(a)%>) �̹� ����ľ����Դϴ�. ���� ���� ��ǰ ��ϿϷ�! <%=drawitemid_re(a)%> ������ǰ���� �ٽ� ����ϼ���');
			opener.opener.location.reload();
			opener.frm.drawitemid.value = '';		
			self.close();
		</script>
		<%
		dbget.close()	:	response.End	
		end if
	end if
	rsget.close
	%>
	
<%
		set oip = new Cfitemlist        	'Ŭ���� ����
		oip.Frectitemid = drawitemid_re(a)		'��ǰ���� ������ ���鼭 sql������ ��ǰid�� �ְ� ��ǰ������ �޾ƿ´�.
		oip.fbrandinsert()					'Ŭ������ ����

	for i=0 to oip.FTotalCount - 1 	' �����ؼ� �޾ƿ� ��ǰ������ �ɼ� ��,���� ���� �Ѹ���. 

	sql = "INSERT INTO [db_summary].[dbo].tbl_req_realstock (itemgubun,itemid,itemoption,reguserid,statecd) VALUES "	& VbCrlf
	sql = sql & "('" & itemgubun & "'"		& VbCrlf
	sql = sql & ",'" & drawitemid_re(a) & "'"			& VbCrlf
	sql = sql & ",'" & oip.flist(i).fitemoption & "'"		& VbCrlf
	sql = sql & ",'" & jisiid & "'"			& VbCrlf
	sql = sql & ",'" & stats & "')"
	'response.write sql&"<br>"			'������ ȭ�鿡 �ѷ�����
	dbget.execute sql
	next	
next
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

<script language="javascript">
	alert('����Ǿ����ϴ�');
	opener.opener.location.reload();
	opener.frm.drawitemid.value = '';
	self.close();
</script>