<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doCdChange.asp
' Discription : ī�װ� ��ǰ ���� ó�� ������
' History : 2008.03.20 ������ : ���� Admin���� ����/����
'###############################################

dim cd1,cd2,cd3,itemidarr
dim cd2slice,cd3slice

cd1 = RequestCheckvar(request("cd1"),10)
cd2 = request("cd2")
cd3 = request("cd3")
cd2slice = split(cd2,",")
cd2 = RequestCheckvar(cd2slice(1),10)
cd3slice = split(cd3,",")
cd3 = RequestCheckvar(cd3slice(2),10)

'response.write cd1 + "<br>"
'response.write cd2 + "<br>"
'response.write cd3 + "<br>"
'dbACADEMYget.close()	:	response.End

itemidarr = request("itemidarr")
itemidarr = Left(itemidarr,Len(itemidarr)-1)
  	if itemidarr <> "" then
		if checkNotValidHTML(itemidarr) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end if
dim sqlStr

if cd1 <> "" and cd2 <> "" and cd3 <> "" then

	'// ��ǰ �⺻ ī�װ� ����   '''" 	lastupdate=getdate()" &_   ::tbl_Item_category ���� Ʈ���� �۵��ϹǷ� ��..
	sqlStr = "update db_academy.dbo.tbl_diy_item" &_
			" set cate_large='" + cd1 + "'," &_
			"	cate_mid='" + cd2 + "'," &_
			"	cate_small='" + cd3 + "'" &_
			" where itemid in (" + itemidarr + ") " & vbCrLf

	'// ��ǰ-ī�װ� ���� ���̺� ����(�⺻ �ڵ� ������)
	sqlStr = sqlStr & "Update db_academy.dbo.tbl_diy_item_category " &_
			" set code_large='" + cd1 + "'," &_
			"	code_mid='" + cd2 + "'," &_
			"	code_small='" + cd3 + "'" &_
			" where code_div='D' " &_
			"	and itemid in (" + itemidarr + ")"
	dbACADEMYget.Execute(sqlStr)

else
	response.write "<script language='javascript'>alert('ī�װ��� �����Ͻ� �����̽��ϴ�.');history.back();</script>"
	dbACADEMYget.close()	:	response.End
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('���� �Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->