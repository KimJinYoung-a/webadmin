<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2010.04.07 ������ ����
'	Description : Favorite Colore ����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,itemid,sortNo,category,colorCD,idx
dim viewidx,disptitle,allusing
dim i

mode = request("mode")
category = request("category")
colorCD = request("colorCD")
itemid = request("itemid")
idx = request("idx")
sortNo = request("sortNo")
viewidx = request("viewidx")
disptitle = request("disptitle")
allusing = request("allusing")

'// ���۵� ������ �ڵ尪 Ȯ��
if Right(itemid,1)="," then
	itemid = Left(itemid,Len(itemid)-1)
end if
if Right(idx,1)="," then
	idx = Left(idx,Len(idx)-1)
end if

dim sqlStr,msg

on error resume  next 

dbget.BeginTrans

if mode="del" then
	sqlStr = "delete from db_sitemaster.dbo.tbl_favoriteColor" &_
				" where  idx in (" + idx + ") "

elseif mode="add" then
	if Not(category="" or colorCD="") then
		sqlStr = "insert into db_sitemaster.dbo.tbl_favoriteColor" &_
					" (category, colorCD, itemid)" &_
					" select '" + Cstr(category) + "','" & colorCD & "', itemid" &_
					" from [db_item].[dbo].tbl_item" &_
					" where itemid in (" + itemid + ")" 
	end if

elseif mode="isUsingValue" then
	sqlStr = " update db_sitemaster.dbo.tbl_favoriteColor " &_
				" set isusing='" & allusing & "'" &_
				" where idx in (" & idx & ") "

elseif mode="ChangeSort" then
	idx = split(idx,",")
	sortNo = split(sortNo,",")
	sqlStr = ""
	for i=0 to ubound(idx)
		sqlStr = sqlStr & " update db_sitemaster.dbo.tbl_favoriteColor " &_
					" set sortNo='" & sortNo(i) & "'" &_
					" where idx='" & idx(i) & "';" & vbCrLf
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
