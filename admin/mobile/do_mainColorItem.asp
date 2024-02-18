<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	Description : 모바일 사이트 컬러별 상품 목록 처리
'	History	:  2010.02.258 허진원
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,itemid,sortNo,ccd
dim viewidx,disptitle,allusing
dim i

mode = request("mode")
ccd = request("ccd")
itemid = request("itemid")
sortNo = request("sortNo")
viewidx = request("viewidx")
disptitle = request("disptitle")
allusing = request("allusing")

'// 전송된 아이템 코드값 확인
if Right(trim(itemid),1)="," then
	itemid = Left(itemid,Len(itemid)-1)
end if
if Right(trim(ccd),1)="," then
	ccd = Left(ccd,Len(ccd)-1)
end if

dim sqlStr,msg

on error resume  next 

dbget.BeginTrans

if mode="del" then
	itemid = split(itemid,",")
	ccd = split(ccd,",")
	sqlStr = ""
	for i=0 to ubound(itemid)
		sqlStr = sqlStr & " Delete From db_sitemaster.dbo.tbl_mobile_main_colorItem " &_
					" where itemid='" & itemid(i) & "'" &_
					"	and colorCode='" & ccd(i) & "';" & vbCrLf
	next


elseif mode="add" then
	sqlStr = "insert into db_sitemaster.dbo.tbl_mobile_main_colorItem" &_
				" (colorCode, itemid)" &_
				" select '" + Cstr(ccd) + "', itemid" &_
				" from [db_item].[dbo].tbl_item" &_
				" where itemid in (" + itemid + ")" 

elseif mode="ChangeSort" then
	ccd = split(ccd,",")
	itemid = split(itemid,",")
	sortNo = split(sortNo,",")
	sqlStr = ""
	for i=0 to ubound(itemid)
		sqlStr = sqlStr & " update db_sitemaster.dbo.tbl_mobile_main_colorItem " &_
					" set sortNo='" & sortNo(i) & "'" &_
					" where itemid='" & itemid(i) & "'" &_
					"	and colorCode='" & ccd(i) & "';" & vbCrLf
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
