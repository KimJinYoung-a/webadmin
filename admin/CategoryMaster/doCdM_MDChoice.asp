<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,itemid,cdl,cdm
dim viewidx,disptitle,allusing
dim i

mode = request("mode")
cdl = request("cdl")
cdm = request("cdm")
itemid = request("itemid")
viewidx = request("viewidx")
disptitle = request("disptitle")
allusing = request("allusing")

'// ���۵� ������ �ڵ尪 Ȯ��
if Right(itemid,1)="," then
	itemid = Left(itemid,Len(itemid)-1)
end if

dim sqlStr,msg

on error resume  next 

dbget.BeginTrans
if mode="del" then
	sqlStr = "delete from [db_sitemaster].[dbo].tbl_category_MDChoice" &_
				" where  itemid in (" + itemid + ") and cdl = '"&cdl&"' and cdm ='"&cdm&"'"
elseif mode="add" then
	sqlStr = "insert into [db_sitemaster].[dbo].tbl_category_MDChoice" &_
				" (cdl, cdm, itemid)" &_
				" select cate_large,cate_mid , itemid" &_
				" from [db_item].[dbo].tbl_item" &_
				" where itemid in (" + itemid + ") "
elseif mode="isUsingValue" then
	sqlStr = " update [db_sitemaster].[dbo].tbl_category_MDChoice " &_
				" set isusing='" & allusing & "'" &_
				" where itemid in (" & itemid & ") and cdl= '"&cdl&"' and cdm ='"&cdm&"'"
end if

dbget.execute(sqlStr)

if err.number<>0 then
	dbget.rollback
	msg ="���� �߻�, �����ڹ��� ���"
else
	dbget.committrans
	msg ="���� �Ǿ����ϴ�."
end if
dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('<%= msg %>');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->