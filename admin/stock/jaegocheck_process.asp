<%@ language = vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����ľ�
' History : 2007.07.13 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/jaegostock.asp"-->

<%
dim fnow,idx, fmode , jisiname , order , smallimage,itemid,makerid,itemname,itemoption,imagesrc,realstock	'��������
dim errstock,actionstartdate,itemgubun,jisiid,stats			'��������
	idx = html2db(request("idx"))							'���̺��� �ε������� �޾ƿ´�
	fmode = html2db(request("mode")	)						'��屸��
	jisiname = html2db(request("jisiname"))					'�۾��������̸�
	jisiid = html2db(session("ssBctId"))					'������ ���� ���� id�� �޾ƿ´�.
	order = now()											'�۾�������
	smallimage = html2db(request("smallimage"))				'�̹���
	itemid = html2db(request("itemid"))						'��ǰid
	makerid = html2db(request("makerid"))					'�귣��id
	itemname = html2db(request("itemname"))					'��ǰ��
	itemoption = html2db(request("itemoption"))				'��ǰ�ɼ��ڵ�
		if itemoption = "" then								'��ǰ�ɼ��ڵ尡 ������
			itemoption = "0000"								'�⺻��0 �Է�
		end if
	realstock = html2db(request("realstock"))				'����ľ������
	errstock = html2db(request("errstock"))					'����
	actionstartdate = Left(request("actionstartdate"),10)	'����ľ��Ͻ�
	itemgubun = html2db(request("itemgubun"))				'��ǰ��,��������
	stats = 1												'���� �⺻�� 1
	imagesrc = request("imagesrc")							'��ǰ�̹���
%>
	
<% 
dim sql , refer , sql111			'��������

if fmode = "" then					'����ľ� ���� ���
	%>

	<%
	dim sql12
		sql12 = "select * from [db_summary].[dbo].tbl_req_realstock" 
		sql12 = sql12 & " where itemid = '"& itemid &"' order by statecd asc"
		rsget.open sql12,dbget,1
	
	if not rsget.eof then				'���ڵ尡 �ִٸ�
		if rsget("statecd") = 1 then	'��ǰ���� ���°��� �۾�������(1) �̶��	
			rsget.close
	%>		
		<script language="javascript">
			alert('������ ��ǰ�� ����ľ����Դϴ�. Ȯ���Ͻ��� �ٽ��Է��ϼ���');
			opener.location.reload();
			self.close();
			</script>
	<%
		dbget.close()	:	response.End	
		end if
	end if
	rsget.close
	
	sql = "INSERT INTO [db_summary].[dbo].tbl_req_realstock" 		'��ǰ�ڵ�� �ɼ��� �����ؼ� ���� �Ѵ�.
	sql = sql & " (itemgubun,itemid,itemoption)"
	sql = sql & " select a.itemgubun,a.itemid,isnull(b.itemoption,'0000')"
	sql = sql & " from [db_item].[dbo].tbl_item a"
	sql = sql & " left join [db_item].[dbo].tbl_item_option b" 
	sql = sql & " on a.itemid = b.itemid"
	sql = sql & " where a.itemid = '" & itemid &"'"
	'response.write sql			'������ ȭ�鿡 �ѷ�����
	dbget.execute sql
	
	sql = ""
	sql = "update [db_summary].[dbo].tbl_req_realstock set"
	sql = sql & " itemgubun='" & itemgubun & "'"		& VbCrlf
	sql = sql & " ,reguserid='" & jisiid & "'"			& VbCrlf
	sql = sql & " ,statecd='" & stats & "'"
	sql = sql & " where 1=1 and itemid = '" & itemid & "' and statecd is null" 
	'response.write sql			'������ ȭ�鿡 �ѷ�����
	dbget.execute sql

	%>
					
	<script language="javascript">
		opener.location.reload();
		self.close();
	</script>	

<!--������� ����-->
	<% 
	elseif fmode = "del" then				
	sql = "delete from [db_summary].[dbo].tbl_req_realstock where idx=" & idx
	'response.write sql			'������ ȭ�鿡 �ѷ����� 	
	dbget.execute sql
	refer = request.ServerVariables("HTTP_REFERER")			'������������ ������ �����´�
	%>
	<script language="javascript">
	location.replace('<%= refer %>');
	</script>
<!--������� ��-->


<!--�ݿ�������-->
	<% elseif fmode = "banyoung" then 
	
'	response.write "������..."
'	dbget.close()	:	response.End
	
	sql = "exec db_summary.dbo.ten_realchekErr_Input '"& actionstartdate &"', '"& itemgubun &"', '"& itemid &"' , '"& itemoption &"', "& errstock &", '"& jisiid &"'"
	dbget.execute sql
	'response.write sql			'���� �ѷ�����
	
	sql111 = "update [db_summary].[dbo].tbl_req_realstock set finishdate = '"& order &"' , statecd = '7' , finishuserid = '"& jisiid &"'" 	& VbCrlf
	sql111 = sql111 & " where idx = '"& idx &"'"
	dbget.execute sql111
	'response.write sql			'���� �ѷ�����
	
	refer = request.ServerVariables("HTTP_REFERER")			'������������ ������ �����´�
	%>
	
	<script language="javascript">
	{
		alert('����Ǿ����ϴ�.���ÿ���,�Ǹſ���,��뿩��,��������,�������θ� �����Ͻʽÿ�');
		location.replace('<%= refer %>');		
		var edit = window.open("itemcurrentstock.asp?itemgubun=<%=itemgubun%>&itemid=<%=itemid%>&itemoption=<%=itemoption%>", "jaegoadd" , 'width=1024,height=768,scrollbars=yes,resizable=yes');
		edit.focus();
	}		
	</script>
	
<!--�ݿ���峡-->		

<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->