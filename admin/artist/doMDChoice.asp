<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim mode,itemid,sortNo,cdl
Dim viewidx,disptitle,allusing
Dim i, idx

mode = request("mode")
cdl = request("cdl")
itemid = request("itemid")
sortNo = request("sortNo")
viewidx = request("viewidx")
disptitle = request("disptitle")
allusing = request("allusing")
idx = request("idx")


'// 전송된 아이템 코드값 확인
If Right(itemid,1)="," Then
	itemid = Left(itemid,Len(itemid)-1)
End If

If Right(idx,1)="," Then
	idx = Left(idx,Len(idx)-1)
End If

Dim sqlStr,msg
on error resume  next 

dbget.BeginTrans

If mode="del" Then
	sqlStr = " delete from db_contents.dbo.tbl_artist_banner " &_
			 " where  idx in (" + idx + ") "
ElseIf mode="add" Then
	sqlStr = "insert into db_contents.dbo.tbl_artist_banner " &_
			" (itemid, gubun, sortNo, isusing)" &_
			" select itemid, 0, 0, 'Y' " &_
			" from [db_item].[dbo].tbl_item" &_
			" where itemid in (" + itemid + ")" 
ElseIf mode="isUsingValue" Then
	sqlStr = " update db_contents.dbo.tbl_artist_banner set " &_
			 " isusing='" & allusing & "'" &_
			 " where idx in (" & idx & ") "
ElseIf mode="ChangeSort" Then
	itemid = split(itemid,",")
	sortNo = split(sortNo,",")
	idx = split(idx,",")
	sqlStr = ""
	For i=0 to ubound(itemid)
		sqlStr = sqlStr & " update db_contents.dbo.tbl_artist_banner set " &_
						  " sortNo='" & sortNo(i) & "'" &_
						  " where idx='" & idx(i) & "' ;" & vbCrLf
	Next
ElseIf mode="hot_ChangeSort" Then
	sortNo = split(sortNo,",")
	idx = split(idx,",")
	sqlStr = ""
	For i=0 to ubound(idx)
		sqlStr = sqlStr & " update db_contents.dbo.tbl_artist_brand set " &_
						  " sortNo='" & sortNo(i) & "'" &_
						  " where idx='" & idx(i) & "' ;" & vbCrLf
	Next
ElseIf mode="hot_isUsingValue" Then
	Dim idxcnt
	idxcnt = split(idx,",")
	sqlStr = sqlStr & " update db_contents.dbo.tbl_artist_brand set " &_
			 " mainHOT='" & allusing & "'" &_
			 " where idx in (" & idx & ") "
End If
dbget.execute(sqlStr)

If err.number<>0 Then
	dbget.rollback
	msg ="오류 발생, 관리자문의 요망"
Else
	dbget.committrans
	msg ="적용 되었습니다."
End If

Dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('<%= msg %>');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
