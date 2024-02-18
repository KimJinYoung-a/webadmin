<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : 스페셜 다이어리 등록 처리 페이지
'	History		: 2015.10.05 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
'// 다이어리 저장폼
dim i
dim iitemid(), detailitemimage()
dim idx, mode, isUsing, sortnum
dim itemid1, itemid2, itemid3, itemid4, itemid5
dim pcmainimage, pcoverimage, pctext
dim mobileimage, mobiletext
dim linkgubun, linkcode

idx = request.Form("did")
mode= request.Form("mode")
isUsing= request.Form("ius")
sortnum = trim(request.Form("sortnum"))

itemid1= trim(request.Form("iid1"))
itemid2= trim(request.Form("iid2"))
itemid3= trim(request.Form("iid3"))
itemid4= trim(request.Form("iid4"))
itemid5= trim(request.Form("iid5"))

pctext		= request.Form("pctext")
pcmainimage = request.Form("pcmainimage")
pcoverimage = request.Form("pcoverimage")

mobileimage = request.Form("mobileimage")
mobiletext = request.Form("mobiletext")

linkgubun = request.Form("linkgubun")
linkcode = trim(request.Form("linkcode"))

dim strSQL,msg

dbget.beginTrans
IF mode="add" Then

	strSQL =" INSERT INTO [db_diary2010].[dbo].[tbl_diaryspecial] "&_
			" (pcmainimage, pcoverimage, pctext, mobileimage, mobiletext, linkgubun, linkcode, sortnum, isusing) "&_
			" VALUES(" &_
			"'" & pcmainimage & "' " &_
			",'" & pcoverimage & "' " &_
			",'" & pctext & "' " &_
			",'" & mobileimage & "' " &_
			",'" & mobiletext & "' " &_
			",'" & linkgubun & "' " & _
			"," & linkcode & " " & _
			"," & sortnum & " " & _
			",'" & isusing & "' "
			strSQL = strSQL & " )"

'	msg = "저장 되었습니다"

'	response.write strSQL&"<br>"
'	response.end

	dbget.execute(strSQL)

	strSQL ="select @@identity "

	rsget.open strSQL,dbget,2
	IF not rsget.Eof Then
		idx = rsget(0)
	End IF
	rsget.close

	idx = idx
	
	ReDim iitemid(4), detailitemimage(4)
	for i = 0 to 4
		iitemid(i) = trim(request.Form("iid"&i+1))
		detailitemimage(i) = trim(request.Form("detailitemimage"&i+1))

		strSQL =" INSERT INTO [db_diary2010].[dbo].[tbl_diaryspecial_detail] "&_
				" (midx, itemid, itemordernum, detailitemimage) "&_
				" VALUES(" &_
				"'" & idx & "' " &_
				",'" & iitemid(i) & "' " &_
				",'" & i+1 & "' " &_
				",'" & detailitemimage(i) & "' " &_
				" )"

'		response.write strSQL & "<br>"
		dbget.execute(strSQL)
	next
'	response.write strSQL
'	response.end

'	dbget.execute(strSQL)

	msg = "저장 되었습니다"

	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)

		Alert_move msg,"/admin/diaryspecial/DiaryspecialReg.asp?mode=edit&idx="& idx

	Else
		dbget.RollBackTrans				'롤백(에러발생시)
		Alert_return "처리중 에러가 발생했습니다."
	End If

ELSEIF mode="edit" Then

	strSQL =" UPDATE [db_diary2010].[dbo].[tbl_diaryspecial] "&_
		" SET pcmainimage= '"& pcmainimage &"'" &_
		", pcoverimage = '"& pcoverimage &"'" &_
		", pctext = '"& pctext &"'" &_
		", mobileimage='" & mobileimage & "' " &_
		", mobiletext='" & mobiletext & "' " &_
		", linkgubun = '" & linkgubun & "' " & _
		", linkcode = '" & linkcode & "' " & _
		", sortnum = '" & sortnum & "' " & _
		", isusing = '" & isusing & "' "
		strSQL = strSQL & " WHERE idx = "& idx
	
	'response.write strSQL&"<br>"	
	
	dbget.execute(strSQL)

	ReDim iitemid(4), detailitemimage(4)
	for i = 0 to 4
		iitemid(i) = trim(request.Form("iid"&i+1))
		detailitemimage(i) = trim(request.Form("detailitemimage"&i+1))

		strSQL =" UPDATE [db_diary2010].[dbo].[tbl_diaryspecial_detail] "&_
			" SET itemid= '"& iitemid(i) & "' " &_
			", detailitemimage = '" & detailitemimage(i) & "' "
			strSQL = strSQL & " WHERE midx = "& idx & " and itemordernum = " & i+1

		dbget.execute(strSQL)
	next
	msg = "저장 되었습니다"

'	response.write strSQL&"<br>"
'	response.end

	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)
		Alert_move msg,"/admin/diaryspecial/DiaryspecialReg.asp?mode=edit&idx="& idx
		
	Else
		dbget.RollBackTrans				'롤백(에러발생시)
		Alert_return "처리중 에러가 발생했습니다."
	End If
End IF
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->