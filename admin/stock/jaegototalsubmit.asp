<%@ language = vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����ľǵ���Ϲ� ����
' History : 2007.07.13 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<% 
dim mode , order , errstock , basicstock , jaego , error , idx , sql,itemid,itemoption	'��������
	mode = request("mode")				'�޾ƿ¸��
	order = now()						'��¥
	errstock = request("errstock")		'����
	jaego = request("jaego")			'�����ľ������
			if jaego = "" then			'�����ľ������ ���� ������
				jaego = 0				'�⺻�� 0
			end if
	
	idx = request("idx") 				'�ε������� �޾ƿ´�.
	itemid = request("itemid")
	itemoption = request("itemoption")

'	sql = "select basicstock from [db_summary].[dbo].tbl_req_realstock"
'	sql = sql & " where 1=1 and itemid = '" & itemid & "' and itemoption = '" & itemoption & "'"
    
    sql = "select (realstock + ipkumdiv5+ offconfirmno) as basicstock from [db_summary].[dbo].tbl_current_logisstock_summary"
    sql = sql & " where itemgubun='10' and itemid = '" & itemid & "' and itemoption = '" & itemoption & "'"

	rsget.open sql,dbget,1
	
	if not rsget.eof then
		basicstock = rsget("basicstock")
	end if
	
	sql = ""
	rsget.close
	
	if basicstock = "" then
	basicstock = 0
	end if
	error = jaego - basicstock			'������ �����ľ�������� ����ľǿ���� ����.


%>

<!--����������-->
<%
if mode = "edit" then
	dim sql50
		sql50 = "update [db_summary].[dbo].tbl_req_realstock set realstock = "& jaego &" , errstock = "& error &""	& VbCrlf
		sql50 = sql50 & " where idx = " & idx 
		'response.write sql50
		dbget.execute sql50
	%>
	<script language="javascript">
	opener.location.reload();
	self.close();
	</script>
<!--������峡-->

<!--����ľ��Ѽ��� ��ϸ�����-->
<% 
elseif mode = "" then
 
	dim sql12 , sql2
	sql12 = "select * from [db_summary].[dbo].tbl_req_realstock" 
	sql12 = sql12 & " where itemid = '"& itemid &"' and itemoption = '" & itemoption & "'" 
	sql12 = sql12 & " order by statecd asc"
	'response.write sql12&"<br>"
	rsget.open sql12,dbget,1
		
	if not rsget.eof then				'���ڵ尡 �ִٸ�		
	
		sql2 = "update [db_summary].[dbo].tbl_req_realstock set"	& VbCrlf
		sql2 = sql2 & " errstock = "& error &" , actiondate = '"& order &"', realstock = "& jaego &" , basicstock = "& basicstock &", statecd = 5" 	& VbCrlf
		sql2 = sql2 & " where 1=1 and itemid = '" & itemid & "' and itemoption = '" & itemoption & "'" 
		''response.write sql2     '���� �ѷ�����.
		dbget.execute sql2		
	else
%>
<script language="javascript">
	if (confirm('����ľ� ���û����� �����ϴ�. ����ľ� ���� �Ͻðڽ��ϱ�?')){
	location.href = '/admin/stock/jaegoadd.asp?itemid=<%=itemid%>&submitview=yes';
	}
</script>	
<%	
	end if
	rsget.close

%>

	<script language="javascript">
	opener.location.reload();
	self.close();
	</script>
<% end if %>
<!--����ľ��Ѽ��� ��ϸ�峡-->

<!-- #include virtual="/lib/db/dbclose.asp" -->