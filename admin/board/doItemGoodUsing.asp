<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,lineSel
dim menupos, page, selStatus, SearchKey1, SearchKey2
dim vSCMChangeSQL, vChangeContents

mode = request("mode")
lineSel = request("lineSel")
menupos = request("menupos")
page = request("page")
selStatus = request("selStatus")
SearchKey1 = request("SearchKey1")
SearchKey2 = request("SearchKey2")

dim SQL

'후기 상태 변경
if (mode = "delete") then
	SQL =	"Update db_board.dbo.tbl_Item_Evaluate Set " &_
			"	isUsing='N' " &_
			"Where IDX in (" & lineSel & ")"
	 dbget.execute(SQL)

	vChangeContents = "- 후기 삭제" & vbCrLf
else
	SQL =	"Update db_board.dbo.tbl_Item_Evaluate Set " &_
			"	isUsing='Y' " &_
			"Where IDX in (" & lineSel & ")"
	dbget.execute(SQL)

	vChangeContents = "- 후기 복원" & vbCrLf
end if

'상품 정보 업데이트
SQL =	"update i " &_
		"set evalcnt=t.cnt " &_
		"	, evalCnt_photo=t.CntP " &_
		"	, evaloffcnt=t.cnt2 " &_
		"from db_item.[dbo].tbl_item as i " &_
		"	join ( " &_
		"		select t1.itemid, sum(t1.cnt) as cnt, sum(t1.CntP) as CntP, sum(t1.cnt2) as cnt2 " &_
		"		from ( " &_
		"			select eo.itemid, isnull(sum(case when eo.isusing='Y' then 1 else 0 end),0) as cnt " &_
		"			,isnull(sum(case when eo.isusing='Y' and (isnull(eo.file1,'')<>'' or isnull(eo.file2,'')<>'' or isnull(eo.file3,'')<>'') then 1 else 0 end),0) as CntP " &_
		"			, isnull(sum(case when eo.isusing='Y' then 1 else 0 end),0) as cnt2 " &_
		"			from db_board.[dbo].tbl_item_evaluate_Offshop as eo " &_
		"				join db_board.[dbo].tbl_item_evaluate as ev1 " &_
		"					on eo.itemid=ev1.itemid " &_
		"			where eo.isusing='Y' " &_
		"				and ev1.idx in (" & lineSel & ") " &_
		"			group by eo.itemid " &_
		"		union all " &_
		"			select ev.itemid, isnull(sum(case when ev.isusing='Y' then 1 else 0 end),0) as cnt " &_
		"			,isnull(sum(case when ev.isusing='Y' and (isnull(ev.file1,'')<>'' or isnull(ev.file2,'')<>'' or isnull(ev.file3,'')<>'') then 1 else 0 end),0) as CntP " &_
		"			,0 as cnt2 " &_
		"			from db_board.[dbo].tbl_item_evaluate as ev " &_
		"				join db_board.[dbo].tbl_item_evaluate as ev2 " &_
		"					on ev.itemid=ev2.itemid " &_
		"			where ev.isusing='Y' " &_
		"				and ev2.idx in (" & lineSel & ") " &_
		"			group by ev.itemid " &_
		"		) as t1 " &_
		"		group by t1.itemid " &_
		"	) as t " &_
		"		on i.itemid=t.ItemID "
dbget.execute(SQL)


'### 수정 로그 저장(review)
vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
vSCMChangeSQL = vSCMChangeSQL & "SELECT '" & session("ssBctId") & "', 'review', IDX, '" & Request("menupos") & "',"
vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "- 주문번호: '+OrderSerial+' / 상품코드: '+ Cast(itemid as varchar(10)) + ' / 옵션코드: ' + itemoption, '" & Request.ServerVariables("REMOTE_ADDR") & "' "
vSCMChangeSQL = vSCMChangeSQL & "FROM db_board.dbo.tbl_Item_Evaluate "
vSCMChangeSQL = vSCMChangeSQL & "WHERE IDX in (" & lineSel & ")"
dbget.execute(vSCMChangeSQL)

Response.Redirect request.ServerVariables("HTTP_REFERER")
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->