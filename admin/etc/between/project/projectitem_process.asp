<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim mode, pjt_code, itemidarr, sGroup, itemid, page, strG
Dim sqlStr
mode		= Request("mode")
itemidarr	= Request("itemidarr")
sGroup		= Request("selGroup")
pjt_code	= request("pjt_code")
itemid      = request("itemid")
page		= request("page")
strG		= Request("selG")

dbCTget.beginTrans
Select Case mode
	Case "I" '// 상품추가
		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO [db_outmall].[dbo].tbl_between_project_groupItem "
		sqlStr = sqlStr & " (pjt_code, itemid, pjtgroup_code, pjtitem_sort)"
		sqlStr = sqlStr & " SELECT " & CStr(pjt_code) & ",i.itemid , '"&sGroup&"', 50 "
		sqlStr = sqlStr & " FROM [db_AppWish].[dbo].tbl_item i"
		sqlStr = sqlStr & " WHERE itemid not in ( "
		sqlStr = sqlStr & " 	SELECT itemid from [db_outmall].[dbo].tbl_between_project_groupItem "
		sqlStr = sqlStr & "		WHERE pjt_code=" & pjt_code
		sqlStr = sqlStr & " )	and itemid in ("&trim(itemidarr)&") "
		dbCTget.execute sqlStr
	Case "D" '// 선택상품 삭제
		sqlStr = ""
		sqlStr = sqlStr & " Delete FROM [db_outmall].[dbo].tbl_between_project_groupItem WHERE pjt_code = '"&pjt_code&"' and itemid in ("&itemidarr&") "
		dbCTget.execute sqlStr
	Case "G" '//그룹이동
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE [db_outmall].[dbo].tbl_between_project_groupItem SET "
		sqlStr = sqlStr & " pjtgroup_code = '"&sGroup&"'"
		sqlStr = sqlStr & "	WHERE pjt_code = '"&pjt_code&"' and itemid in ( "&itemidarr&") "
		dbCTget.execute sqlStr
	Case "S" '//상품순서 저장
		Dim tmpSort, sortarr, cnt, i
		sortarr = Request("sortarr")

		If sortarr="" Then
			dbCTget.RollBackTrans
			Response.Write "<script language='javascript'>history.back(-1);</script>"
			dbCTget.close()	:	response.End
		End if

		'선택상품 파악
		itemidarr = split(itemidarr,",")
		cnt = ubound(itemidarr)

		'// 정렬순서 저장
		If sortarr<>"" THEN
			sortarr =  split(sortarr,",")

			For i = 0 to cnt
				If sortarr(i) = "" Then
					 tmpSort = "NULL"
				Else
					 tmpSort = sortarr(i)
				End If
				sqlStr = ""
				sqlStr = sqlStr & " UPDATE [db_outmall].[dbo].tbl_between_project_groupItem SET "
				sqlStr = sqlStr & " pjtitem_sort = '"&tmpSort&"'"
				sqlStr = sqlStr & "	WHERE pjt_code = "&pjt_code&" and itemid =" & itemidarr(i)
				dbCTget.execute sqlStr
			Next
		End If
End Select

IF Err.Number = 0 THEN
	dbCTget.CommitTrans

	If mode= "I" then
%>
	<script langauge="javascript">
		location.href ="about:blank";
		parent.history.go(0);
	</script>
<%
	Else
		response.redirect("projectitem_regist.asp?pjt_code="&pjt_code&"&menupos="&menupos&"&selG="&strG&"&page="&page)
	End If
	dbCTget.close()	:	response.End
Else
   	dbCTget.RollBackTrans
%>
	<script language="javascript">
	alert("데이터 처리에 문제가 발생하였습니다.");
	history.back(-1);
	</script>
<%
	dbCTget.close()	:	response.End
End IF
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->