<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doCdChange.asp
' Discription : ī�װ� ��ǰ ���� ó�� ������
' History : 2008.03.20 ������ : ���� Admin���� ����/����
'###############################################

dim cd1,cd2,cd3,itemidarr, ocd1, ocd2, codeDiv
dim cd2slice,cd3slice

cd1 = request("cd1")
cd2 = request("cd2")
cd3 = request("cd3")
cd2slice = split(cd2,",")
cd2 = cd2slice(1)
cd3slice = split(cd3,",")
cd3 = cd3slice(2)

'// ī�װ� ���� �߰�
codeDiv = request("codeDiv")
if codeDiv="" then codeDiv="D"

itemidarr = request("itemidarr")
itemidarr = Left(itemidarr,Len(itemidarr)-1)

dim sqlStr

if cd1 <> "" and cd2 <> "" and cd3 <> "" then

	Select Case codeDiv
		Case "D"
		'// ### �⺻ ī�װ� �߰� ###

		'// ��ǰ-ī�װ� �Ӽ� ������� Ȯ�� �� ����
		'������ �⺻ ī�װ� ����
		sqlStr = "Select top 1 code_large, code_mid " &_
				"from db_item.dbo.tbl_Item_category " &_
				"where code_div='D' " &_
				"	and itemid in (" + itemidarr + ")"
		rsget.Open sqlStr,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			ocd1 = rsget(0)
			ocd2 = rsget(1)
		end if
		rsget.Close
	
		'�ߺз��� ������ ������쿡 ���� ī�װ� �Ӽ� ����
		if (ocd1<>cd1 or ocd2<>cd2) then
			sqlStr = "Select attrib_Code " &_
					" from db_item.dbo.tbl_Cate_Attrib_div " &_
					" Where code_large='" + ocd1 + "'" &_
					" 	and code_mid='" + ocd2 + "'"
			rsget.Open sqlStr,dbget,1
			if Not(rsget.EOF or rsget.BOF) then
				sqlStr = "Delete From db_item.dbo.tbl_Item_Attribute " &_
						" Where attrib_Code=" & rsget(0) &_
						"	and itemid in (" + itemidarr + ")"
				dbget.Execute(sqlStr)
			end if
			rsget.Close
		end if
	
		'// ��ǰ �⺻ ī�װ� ����   '''" 	lastupdate=getdate()" &_   ::tbl_Item_category ���� Ʈ���� �۵��ϹǷ� ��..
		sqlStr = "update [db_item].dbo.tbl_item" &_
				" set cate_large='" + cd1 + "'," &_
				"	cate_mid='" + cd2 + "'," &_
				"	cate_small='" + cd3 + "'" &_
				" where itemid in (" + itemidarr + ") " & vbCrLf
	
		'// ��ǰ-ī�װ� ���� ���̺� ����(�⺻ �ڵ� ������)
		sqlStr = sqlStr & "Update db_item.dbo.tbl_Item_category " &_
				" set code_large='" + cd1 + "'," &_
				"	code_mid='" + cd2 + "'," &_
				"	code_small='" + cd3 + "'" &_
				" where code_div='D' " &_
				"	and itemid in (" + itemidarr + ")"
		dbget.Execute(sqlStr)

		Case "A"
		'// ### �߰� ī�װ� �߰� ###
			''���� ī�װ��� ���°�츸 �Է�
			sqlStr = "Insert into [db_item].dbo.tbl_Item_category "
			sqlStr = sqlStr & " (itemid,code_large,code_mid,code_small,code_div)  " 
			sqlStr = sqlStr & " select i.itemid" 
			sqlStr = sqlStr & " ,'" & cd1 & "'" 
			sqlStr = sqlStr & " ,'" & cd2 & "'" 
			sqlStr = sqlStr & " ,'" & cd3 & "'" 
			sqlStr = sqlStr & " ,'A'"
			sqlStr = sqlStr & " from [db_item].dbo.tbl_Item i"
			sqlStr = sqlStr & "     left join [db_item].dbo.tbl_Item_category c"
			sqlStr = sqlStr & "     on i.itemid=c.itemid"
			sqlStr = sqlStr & "     and c.code_large='" & cd1 & "'" 
			sqlStr = sqlStr & "     and c.code_mid='" & cd2 & "'" 
			sqlStr = sqlStr & "     and c.code_small='" & cd3 & "'" 
			sqlStr = sqlStr & " where i.itemid in (" & itemidarr & ")"
			sqlStr = sqlStr & " and c.itemid Is NULL"
			
			dbget.execute(sqlStr)

		Case "DelA"
		'// ### �߰� ī�װ� ���� ###
			sqlStr = "Delete From [db_item].dbo.tbl_Item_category "
			sqlStr = sqlStr & "Where code_large='" & cd1 & "'" 
			sqlStr = sqlStr & " and code_mid='" & cd2 & "'" 
			sqlStr = sqlStr & " and code_small='" & cd3 & "'" 
			sqlStr = sqlStr & " and code_div='A'" 
			sqlStr = sqlStr & " and itemid in (" & itemidarr & ")"
			
			dbget.execute(sqlStr)
	end Select
else
	response.write "<script language='javascript'>alert('ī�װ��� �����Ͻ� �����̽��ϴ�.');history.back();</script>"
	dbget.close()	:	response.End
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('���� �Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->