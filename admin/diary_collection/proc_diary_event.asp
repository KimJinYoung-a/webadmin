<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->

<%

'=====================================
'본문 시작
'=====================================


dim mode
dim bannerid,bannerType,bannerUrl,evtcode,mapusing
dim imagename
dim multiname,leftname,powername,todayname,quizname,isusing,othermall_leftname,othermall_multiname,othermall_rightname

mode= request("mode")
bannerid = request("bannerid")
bannerType = request("bannerType")
bannerUrl = html2db(request("bannerUrl"))
evtcode = request("evtcode")

mapusing= request("mapusing")
imagename = html2db(request("imagename"))


if mapusing="on" or mapusing="Y" then
	mapusing="Y"
else
	mapusing="N"
end if

multiname = request("multiname")
leftname = request("leftname")
powername = request("powername")
todayname = request("todayname")
quizname = request("quizname")
isusing = request("isusing")
othermall_leftname = request("othermall_leftname")
othermall_multiname = request("othermall_multiname")
othermall_rightname = request("othermall_rightname")


dim strSQL


'SELECT CASE bannerType
'	Case "multi"
'		imageName= multiname
'	Case "left"
'		imageName= leftname
'	Case "power"
'		imageName = powername
'	Case "today"
'		imageName = todayname
'	Case "quiz"
'		imageName = quizname
'	Case "othermall_left"
'		imageName = othermall_leftname
'	Case "othermall_multi"
'		imageName = othermall_multiname
'	Case "othermall_right"
'		imageName = othermall_rightname
'End Select
'response.write othermall_leftname


if mode="edit" then
	strSQL =" UPDATE [db_diary_collection].[dbo].[tbl_diary_banner] " &_
			" SET bannerType='" & bannerType & "' " &_
			" ,bannerUrl='" & bannerUrl & "' " &_
			" ,bannerImg='" & imageName & "'" &_
			" ,evt_code ='" & evtcode & "'" &_
			" ,bannerMapUsing = '" & mapusing & "'" &_
			" ,isUsing='" & isusing & "'" &_
			" ,regdate = getdate() " &_
			" WHERE bannerid ='" & bannerid & "'"

else
	strSQL =" INSERT INTO db_diary_collection.[dbo].[tbl_diary_banner](bannerType,bannerUrl,bannerImg,evt_code,bannerMapUsing,isusing) " &_
			" VALUES('" & bannerType & "','" & bannerUrl & "','" & imageName &"','" & evtcode & "','" & mapusing & "','" & isusing & "')"
end if

	'response.write strSQL
	dbget.BeginTrans

	dbget.execute(strSQL)


'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)

		response.write	"<script language='javascript'>"
		response.write	" alert('저장되었습니다.'); document.location.href='pop_diary_event_List.asp?BannerType=" & bannerType & "'"
		response.write	"</script>"
		dbget.close()	:	response.End
	Else
		dbget.RollBackTrans				'롤백(에러발생시)

		response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.go(-1);" &_
					"</script>"


	End If



%>
<!-- #include virtual="/lib/db/dbclose.asp" -->