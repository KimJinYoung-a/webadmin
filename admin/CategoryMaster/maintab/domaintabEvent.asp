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
dim mode,evt_code,sortNo,cdl
dim viewidx,allusing
dim i

mode = request("mode")
cdl = request("cdl")
evt_code = request("evt_code")
sortNo = request("sortNo")
viewidx = request("viewidx")
allusing = request("allusing")



'// 전송된 아이템 코드값 확인
if Right(evt_code,1)="," then
	evt_code = Left(evt_code,Len(evt_code)-1)
end if

dim sqlStr,msg

on error resume  next 

dbget.BeginTrans

if mode="del" then
	sqlStr = "delete from db_sitemaster.dbo.tbl_main_tabEvent" &_
				" where  evt_code in (" + evt_code + ") "

elseif mode="add" then
	sqlStr = "insert into db_sitemaster.dbo.tbl_main_tabEvent" &_
				" (cdl, evt_code)" &_
				" select '" + Cstr(cdl) + "', evt_code" &_
				" from db_event.dbo.tbl_event" &_
				" where evt_code in (" + evt_code + ")" 

elseif mode="isUsingValue" then
	sqlStr = " update db_sitemaster.dbo.tbl_main_tabEvent " &_
				" set isusing='" & allusing & "'" &_
				" where evt_code in (" & evt_code & ") "

elseif mode="ChangeSort" then
	evt_code = split(evt_code,",")
	sortNo = split(sortNo,",")
	sqlStr = ""
	for i=0 to ubound(evt_code)
		sqlStr = sqlStr & " update db_sitemaster.dbo.tbl_main_tabEvent " &_
					" set sortNo='" & sortNo(i) & "'" &_
					" where evt_code='" & evt_code(i) & "';" & vbCrLf
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
