<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : mdpick_Process.asp
' Discription : mdpick 처리 페이지
' History : 2016.08.02 유태욱
'###############################################

dim menupos, mode, sqlStr, lp
dim idx, itemid, title, startdate, enddate, isusing, sortno	''img1
idx			= RequestCheckvar(Request("idx"),10)
mode		= RequestCheckvar(Request("mode"),16)
sortno		= RequestCheckvar(Request("sortno"),10)
'img1		= Request("image1")
menupos	= RequestCheckvar(Request("menupos"),10)
enddate	= RequestCheckvar(Request("enddate"),10)
isusing	= RequestCheckvar(Request("isusing"),1)
startdate	= RequestCheckvar(Request("startdate"),10)
itemid		= getNumeric(Request("itemid"))
title		= html2db(Request("mdpicktitle"))
  	if title <> "" then
		if checkNotValidHTML(title) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
'// 트랜젝션 시작
dbACADEMYget.beginTrans

'// 모드에 따른 분기
Select Case mode
	Case "add"
		'// 신규 등록
		sqlStr = "Insert Into [db_academy].[dbo].tbl_mdpick " &_
				" (itemid, title, startdate, enddate, isusing, sortno, adminid) values " &_
				" (" & itemid & "" &_
				" ,'" & title & "' " &_
				" ,'" & startdate & "' " &_
				" ,'" & enddate & "' " &_
				" ,'" & isusing & "' " &_
				" ,'" & sortno & "' " &_
				" ,'" & session("ssBctId") & "')"

		dbACADEMYget.Execute(sqlStr)

	Case "edit"
		'// 내용 수정
		sqlStr = "Update [db_academy].[dbo].tbl_mdpick " &_
				" Set title='" & title & "'" &_
				" 	,startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,sortno='" & sortno & "'" &_
				" Where idx='" & idx & "'"
'response.write sqlStr
'response.end
		dbACADEMYget.Execute(sqlStr)

	Case "delete"
'		if idx <> "" then
'			sqlStr = sqlStr & "delete [db_sitemaster].[dbo].tbl_mdpick " &_
'					" Where idx='" & idx & "';" & vbCrLf
'			dbACADEMYget.Execute(sqlStr)
'		end if
End Select


'// 트랜젝션 검사 및 실행
If Err.Number = 0 Then
        dbACADEMYget.CommitTrans
Else
        dbACADEMYget.RollBackTrans
		Alert_return("데이타를 저장하는 도중에 에러가 발생하였습니다.")
		dbACADEMYget.close()	:	response.End
End If

%>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "mdpick_list.asp?menupos=<%=menupos%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
