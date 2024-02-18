<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,itemid,sortNo,cdl, cdm, vIdx
dim viewidx,disptitle,allusing
dim i

mode = request("mode")
cdl = request("cdl")
cdm = request("cdm")
itemid = Trim(Replace(request("itemid")," ",""))
sortNo = request("sortNo")
viewidx = request("viewidx")
disptitle = request("disptitle")
allusing = request("allusing")
vIdx = request("idx")

'// 전송된 아이템 코드값 확인
if Right(itemid,1)="," then
	itemid = Left(itemid,Len(itemid)-1)
end if

dim sqlStr,msg

on error resume  next 

dbget.BeginTrans

if mode="del" then
	sqlStr = "delete from [db_sitemaster].[dbo].tbl_category_contents_brand where itemid in (" + itemid + ") and tidx = '" & vIdx & "' "

elseif mode="add" then
	sqlStr = "insert into [db_sitemaster].[dbo].tbl_category_contents_brand" &_
				" (tidx, itemid, sortNo, isusing)" &_
				" select '" & vIdx & "', itemid, '99', 'Y' " &_
				" from [db_item].[dbo].tbl_item" &_
				" where itemid in (" & itemid & ")" 

elseif mode="isUsingValue" then
	sqlStr = "update [db_sitemaster].[dbo].tbl_category_contents_brand set isusing = '" & allusing & "' where itemid in (" & itemid & ") and tidx = '" & vIdx & "'"

elseif mode="ChangeSort" then
	itemid = split(itemid,",")
	sortNo = split(sortNo,",")
	sqlStr = ""
	for i=0 to ubound(itemid)
		sqlStr = sqlStr & " update [db_sitemaster].[dbo].tbl_category_contents_brand " &_
					" set sortNo='" & sortNo(i) & "'" &_
					" where itemid='" & itemid(i) & "' and tidx = '" & vIdx & "';" & vbCrLf
	next

end if

dbget.execute(sqlStr)


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
//alert('<%= msg %>');
opener.location.reload();
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
