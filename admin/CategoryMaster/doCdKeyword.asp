<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doCdKeyword.asp
' Discription : ī�װ� Ű���� ����ó�� ������
' History : 2008.03.20 ������ : ���� Admin���� ����/����
'###############################################

dim cdl,cdm,cds,keyword

cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")
keyword = trim(html2db(request("keyword")))

dim sqlStr

if cdl <> "" and cdm <> "" and cds <> "" then

	sqlStr = "update [db_item].dbo.tbl_Cate_small"
	sqlStr = sqlStr + " set keyword ='" + keyword + "'"
	sqlStr = sqlStr + " where code_large='" + cdl + "'"
	sqlStr = sqlStr + " and code_mid='" + cdm + "'"
	sqlStr = sqlStr + " and code_small='" + cds + "'"
	rsget.Open sqlStr,dbget,1

else
response.write "<script language='javascript'>alert('ī�װ� �����Ͱ� �����ϴ�.\nī�װ��� �ٽ� �����ϰ� Ű���带 �Է����ּ���');history.back();</script>"
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