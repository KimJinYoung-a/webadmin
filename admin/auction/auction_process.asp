<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��� �ϰ� ��� ������
' History : 2007.09.28 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/auction/auctionclass.asp"-->
<%
dim idx, mode, itemid , ing_rectitemid
dim sql , i ,j, ing_itemid , arr_ing_itemid , arr_itemid
	itemid = request("itemid")
	idx = html2db(request("idx"))							'���̺��� �ε������� �޾ƿ´�
	mode = html2db(request("mode"))						'��屸��
	ing_rectitemid = left(itemid,len(itemid)-1)

	'response.write ing_rectitemid&"<br>"
%>

<%
	if mode = "" then 
		
	'// �ߺ��Ǵ� ��ǰ ������Ʈ ó��
	sql = "update db_item.dbo.tbl_auction set regdate= getdate()"
	sql = sql & " where 1=1 and" 
	sql = sql & " ten_itemid in ("
	sql = sql & " select ten_itemid" 
	sql = sql & " 	from db_item.dbo.tbl_auction"
	sql = sql & " 	where ten_itemid in ("& ing_rectitemid &")"
	sql = sql & " )"
	
	response.write sql&"<br>"			'������ ȭ�鿡 �ѷ�����
	dbget.execute sql	
	sql = ""	
		
	'//�ߺ��� ���� �ʴ� ��ǰ �˻��� ����
	sql = "insert into db_item.dbo.tbl_auction (ten_itemid ,ten_option)" 
	sql = sql & "(select a.itemid , b.optionname"
	sql = sql & "	from db_item.dbo.tbl_item a" 
	sql = sql & "	left join db_item.dbo.tbl_item_option b" 
	sql = sql & "	on a.itemid = b.itemid" 
	sql = sql & "	where 1=1" 
	sql = sql & "	and a.itemid in ("& ing_rectitemid &")" 
	sql = sql & "	and a.itemid not in (" 
	sql = sql & "		select ten_itemid " 
	sql = sql & "		from db_item.dbo.tbl_auction" 
	sql = sql & "		where 1=1 " 
	sql = sql & "		and ten_itemid in ("& ing_rectitemid &")" 
	sql = sql & "		)" 
	sql = sql & ")" 
	
	response.write sql&"<br>"			'������ ȭ�鿡 �ѷ�����
	dbget.execute sql	
	sql = ""
%>

	<script language="javascript">
		parent.opener.location.reload();
		parent.close();
	</script>

<%	
elseif mode="del" then
	
	sql = "delete from db_item.dbo.tbl_auction where idx = "& idx &""
	response.write sql&"<br>"			'������ ȭ�鿡 �ѷ�����
	dbget.execute sql	
	sql = ""
%>

	<script language="javascript">
		parent.location.reload();
	</script>
	
<% end if %>		
	
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->


