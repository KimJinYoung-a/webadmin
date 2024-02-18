<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,itemid,sortNo,cate
dim viewidx,disptitle,allusing
dim i, k
dim itemids
dim idx
mode = request("mode")
cate = request("cate")
itemid = request("itemid")

idx = request("idx")
sortNo = request("sortNo")
viewidx = request("viewidx")
disptitle = request("disptitle")
allusing = request("allusing")


'// 전송된 아이템 코드값 확인
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
	sqlStr = "delete from [db_diary2010].[dbo].tbl_diary_withbuy " &_
				" where  itemid in (" + itemid + ") "
	dbget.execute(sqlStr)
elseif mode="add" then
	itemid = split(itemid,",")
	For k = 0 to Ubound(itemid)
	sqlStr = "insert into [db_diary2010].[dbo].tbl_diary_withbuy " &_
			" (cate, itemid) values ("& cate &", "& itemid(k) &")"
	dbget.execute(sqlStr)
	Next
elseif mode="isUsingValue" then
	sqlStr = " update [db_diary2010].[dbo].tbl_diary_withbuy " &_
				" set isusing='" & allusing & "'" &_
				" where itemid in (" & itemid & ") "
	dbget.execute(sqlStr)
elseif mode="ChangeSort" then
	itemid = split(itemid,",")
	sortNo = split(sortNo,",")
	sqlStr = ""
	for i=0 to ubound(itemid)
		sqlStr = sqlStr & " update [db_diary2010].[dbo].tbl_diary_withbuy " &_
					" set arrayno='" & sortNo(i) & "'" &_
					" where itemid='" & itemid(i) & "' ;" & vbCrLf
	next
	dbget.execute(sqlStr)
elseif mode="modify" then
		sqlStr = sqlStr & " update [db_diary2010].[dbo].tbl_diary_withbuy " &_
					" set cate='" & cate & "'" &_
					" where idx in (" & idx & ") "

	dbget.execute(sqlStr)
end if



if err.number<>0 then
	dbget.rollback
	msg ="오류 발생, 관리자문의 요망"
else
	dbget.committrans
	msg ="적용 되었습니다."
end if
dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('<%= msg %>');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
