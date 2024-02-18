<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.18 한용민 카테고리md픽 이동/ 추가/수정
'	Description : 메인페이지 탭관리
'#######################################################
%>
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



'// 전송된 아이템 코드값 확인
if Right(itemid,1)="," then
	itemid = Left(itemid,Len(itemid)-1)
end if

dim sqlStr,msg

on error resume  next 

dbget.BeginTrans

if mode="del" then
	sqlStr = "delete from db_sitemaster.dbo.tbl_main_tabitem" &_
				" where  itemid in (" + itemid + ") and cdm = '0' "

elseif mode="add" then
	sqlStr = "insert into db_sitemaster.dbo.tbl_main_tabitem" &_
				" (cdl, itemid)" &_
				" select '" + Cstr(cdl) + "', itemid" &_
				" from [db_item].[dbo].tbl_item" &_
				" where itemid in (" + itemid + ")" 

elseif mode="isUsingValue" then
	sqlStr = " update db_sitemaster.dbo.tbl_main_tabitem " &_
				" set isusing='" & allusing & "'" &_
				" where itemid in (" & itemid & ") and cdm = '0'"

elseif mode="ChangeSort" then
	itemid = split(itemid,",")
	sortNo = split(sortNo,",")
	sqlStr = ""
	for i=0 to ubound(itemid)
		sqlStr = sqlStr & " update db_sitemaster.dbo.tbl_main_tabitem " &_
					" set sortNo='" & sortNo(i) & "'" &_
					" where itemid='" & itemid(i) & "' and cdm = '0';" & vbCrLf
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
alert('<%= msg %>');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
