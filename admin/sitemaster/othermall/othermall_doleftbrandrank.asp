<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2007.11.09 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->

<%
dim mode,makerid,idx,cdl , i
	mode = request("mode")
	cdl = request("cdl")
	makerid = request("makerid")
	idx = request("itemid")
	
	If idx <> "" then
	idx = Left(idx,Len(idx)-1)
	End if

dim sqlStr
if mode="del" then
	
	if idx = "" then
%>
	<script language="javascript">
		alert('�ε��� ���� �����ϴ�.');
		location.replace('<%= refer %>');
	</script>
<%	
	end if
	
	sqlStr = "delete from [db_contents].[dbo].tbl_category_left_brand_rank"
	sqlStr = sqlStr + " where idx in (" + idx + ")"
	'response.write sqlStr
	
	rsget.Open sqlStr,dbget,1
	
elseif mode="add" then

	if makerid = "" or cdl = "" then
%>
	<script language="javascript">
		alert('��ü[<%=makerid%>]�� ī�װ���[<%=cdl%>]���� �����ϴ�.');
		location.replace('<%= refer %>');
	</script>
<%	
	end if

	sqlStr = "insert into [db_contents].[dbo].tbl_category_left_brand_rank(cdl,makerid)"
	sqlStr = sqlStr + " values('" + Cstr(cdl) + "','" + makerid + "')"

	response.write sqlStr	
	rsget.Open sqlStr,dbget,1
	
end if

dim refer
	refer = request.ServerVariables("HTTP_REFERER")
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

	<script language="javascript">
		alert('���� �Ǿ����ϴ�.');
		location.replace('<%= refer %>');
	</script>
	