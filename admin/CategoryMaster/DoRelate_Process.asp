<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : doRelate_Process.asp
' Discription : ī�װ� ���� Ű���� ó�� ������
' History : 2008.03.29 ������ ����
'			2022.07.05 �ѿ�� ����(isms�������ġ, ǥ���ڵ����κ���)
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode, linkCode, cdl,cdm,cds, linkKeyword, linkURL, menupos
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	mode = requestcheckvar(request("mode"),32)
	LinkCode = requestcheckvar(getNumeric(request("rid")),10)
cdl			= request("cdl")
cdm			= request("cdm")
cds			= request("cds")
linkKeyword = trim(html2db(request("linkKeyword")))
linkURL		= trim(html2db(request("linkURL")))

dim sqlStr

'// ó�� �б�
Select Case mode
	Case "add"
		if linkKeyword <> "" and not(isnull(linkKeyword)) then
			linkKeyword = ReplaceBracket(linkKeyword)
		end If
		if linkURL <> "" and not(isnull(linkURL)) then
			linkURL = ReplaceBracket(linkURL)
		end If

		'### �ű� ��� ###
		sqlStr = "insert into [db_item].dbo.tbl_Cate_RelateLink"
		sqlStr = sqlStr + " (code_large, code_mid, code_small, linkKeyword, linkURL) values "
		sqlStr = sqlStr + " ('" + cdl + "'"
		sqlStr = sqlStr + " ,'" + cdm + "'"
		sqlStr = sqlStr + " ,'" + cds + "'"
		sqlStr = sqlStr + " ,'" + linkKeyword + "'"
		sqlStr = sqlStr + " ,'" + linkURL + "')"
		rsget.Open sqlStr,dbget,1

	Case "modify"
		if linkKeyword <> "" and not(isnull(linkKeyword)) then
			linkKeyword = ReplaceBracket(linkKeyword)
		end If
		if linkURL <> "" and not(isnull(linkURL)) then
			linkURL = ReplaceBracket(linkURL)
		end If

		'### ���� ###
		sqlStr = "update [db_item].dbo.tbl_Cate_RelateLink"
		sqlStr = sqlStr + " set code_large ='" + cdl + "'"
		sqlStr = sqlStr + "		,code_mid ='" + cdm + "'"
		sqlStr = sqlStr + "		,code_small ='" + cds + "'"
		sqlStr = sqlStr + "		,linkKeyword ='" + linkKeyword + "'"
		sqlStr = sqlStr + "		,linkURL ='" + linkURL + "'"
		sqlStr = sqlStr + " where linkCode='" + linkCode + "'"
		rsget.Open sqlStr,dbget,1

	Case "delete"
		'### ���� ###
		sqlStr = "delete [db_item].dbo.tbl_Cate_RelateLink"
		sqlStr = sqlStr + " where linkCode='" + linkCode + "'"
		rsget.Open sqlStr,dbget,1

	Case Else
		response.write "<script type='text/javascript'>alert('�������� ������ �ƴմϴ�.\n�����ڿ��� �����ּ���.');history.back();</script>"
		dbget.close()	:	response.End
end Select

%>
<script type='text/javascript'>
alert('���� �Ǿ����ϴ�.');
location.replace("RelateKeywordLink_list.asp?cdl=<%=cdl%>&cdm=<%=cdm%>&cds=<%=cds%>&menupos=<%= menupos %>");
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->