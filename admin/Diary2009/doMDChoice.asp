<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,itemid,sortNo,cdl
dim viewidx,disptitle,allusing
dim i

mode = request("mode")
cdl = request("cdl")
itemid = request("itemid")
sortNo = request("sortNo")
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
	sqlStr = "delete from [db_diary2010].[dbo].tbl_category_MDChoice" &_
				" where  itemid in (" + itemid + ") and cdm = '0' "

elseif mode="add" then
	sqlStr = "insert into [db_diary2010].[dbo].tbl_category_MDChoice" &_
				" (cdl, itemid)" &_
				" select '" + Cstr(cdl) + "', itemid" &_
				" from [db_item].[dbo].tbl_item" &_
				" where itemid in (" + itemid + ")" 

elseif mode="isUsingValue" then
	sqlStr = " update [db_diary2010].[dbo].tbl_category_MDChoice " &_
				" set isusing='" & allusing & "'" &_
				" where itemid in (" & itemid & ") and cdm = '0'"

elseif mode="ChangeSort" then
	itemid = split(itemid,",")
	sortNo = split(sortNo,",")
	sqlStr = ""
	for i=0 to ubound(itemid)
		sqlStr = sqlStr & " update [db_diary2010].[dbo].tbl_category_MDChoice " &_
					" set sortNo='" & sortNo(i) & "'" &_
					" where itemid='" & itemid(i) & "' and cdm = '0';" & vbCrLf
	next

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
