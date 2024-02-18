<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<!-- #include virtual="/admin/diary2009/lib/include_event_code.asp"-->

<%
'// 다이어리 저장폼
dim cnt, i

dim Diaryid
dim mode
dim CateCode
dim isUsing
dim itemid
dim basicimgName
dim idx
dim giftYn
dim commentyn
dim Weight

dim basicimg2Name
dim basicimg3Name
dim MDPick
dim limited
dim soonseo
dim storyimgName
dim storytext, nanumimgName, reservdate, mdpicksort

mode= request.Form("mode")
CateCode= request.Form("cate")
isUsing= request.Form("ius")
itemid= request.Form("iid")
Diaryid = request.Form("did")
Weight = request.Form("wt")
basicimgName = request.Form("basicimgName")
commentyn = request.Form("commentyn")

basicimg2Name = request.Form("basicimgName2")
basicimg3Name = request.Form("basicimgName3")
MDPick = request.Form("mdpick")
limited = request.Form("limited")
soonseo = request.Form("soonseo")
storyimgName = request.Form("storyimgName")
storytext = request.Form("storytext")
nanumimgName = request.Form("nanumimgName")
reservdate = request.Form("reservdate")
mdpicksort = request.Form("mdpicksort")

dim Diaryid_newinsert , info_gubun_newinsert ,info_name_newinsert , mode_newinsert , info_gubun_delete , eventgroup_code
dim event_code , commentimgName , event_start , event_end
	info_gubun_delete = request("info_gubun_delete")
	Diaryid_newinsert= request("Diaryid_newinsert")
	info_gubun_newinsert= request("info_gubun_newinsert")
	info_name_newinsert= request("info_name_newinsert")
	mode_newinsert = request("mode_newinsert")
	commentimgName = request("commentimgName")
	event_code = request("event_code")
	eventgroup_code = request("eventgroup_code")
	event_start = request("event_start")
	event_end = request("event_end")
	
	
dim strSQL,msg
	
dbget.beginTrans
IF mode="add" Then
	strSQL =" INSERT INTO db_diary2010.[dbo].tbl_DiaryMaster "&_
			" (Cate,Itemid,BasicImg,isusing,commentyn,event_code,eventgroup_code,comment_img ,weight, BasicImg2, BasicImg3, StoryImg, mdpick, limited, soonseo, storytext , nanumimg, reservdate, mdpicksort, event_start,event_end) "&_
			" VALUES(" &_
			"'" & CateCode & "' " &_
			",'" & itemid & "' " &_
			",'" & basicimgName & "' " &_
			",'" & isUsing & "' " &_
			",'" & commentyn & "' " &_
			",'" & event_code & "' " &_
			",'" & eventgroup_code & "' " &_
			",'" & commentimgName & "' " &_
			",'" & weight & "' " & _
			
			",'" & basicimg2Name & "' " & _
			",'" & basicimg3Name & "' " & _
			",'" & storyimgName & "' " & _
			",'" & MDPick & "' " & _
			",'" & limited & "' " & _
			",'" & soonseo & "' " & _
			",'" & storytext & "' " & _
			",'" & nanumimgName & "' " & _
			",'" & reservdate & "' " & _
			",'" & mdpicksort & "' "
			
			if event_start <>""  then
				strSQL = strSQL & ",'" & event_start & "' "
			else
				strSQL = strSQL & ",null "
			end	if
			if event_end <> "" then
				strSQL = strSQL &  ",'" & event_end & "' "
			else
			 strSQL = strSQL & ",null "
			end if

			strSQL = strSQL & " )"

	msg = "저장 되었습니다"

	'response.write strSQL&"<br>"
	dbget.execute(strSQL)

	strSQL ="select @@identity "

	rsget.open strSQL,dbget,2
	IF not rsget.Eof Then
		Diaryid = rsget(0)
	End IF
	rsget.close

	idx = Diaryid


		If idx<>"" Then

			'2019 다이어리 내지 구성
			strSQL = " INSERT INTO [db_diary2010].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " & vbcrlf
			strSQL = strSQL & " VALUES " & vbcrlf
			'strSQL = strSQL & "('" & idx & "','22','1개월','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','23','분기별','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','24','6개월','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','25','1년','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','26','1년 이상','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','27','연간스케줄','0') ," & vbcrlf
			'strSQL = strSQL & "('" & idx & "','28','월간스케줄','0') ," & vbcrlf
			'strSQL = strSQL & "('" & idx & "','29','주간스케줄','0') ," & vbcrlf
			'strSQL = strSQL & "('" & idx & "','30','일스케줄','0') ," & vbcrlf
			'strSQL = strSQL & "('" & idx & "','31','캐시북','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','32','포켓','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','33','밴드','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','34','펜홀더','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','35','만년형','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','36','2019 날짜형','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','37','먼슬리','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','38','위클리','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','39','데일리','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','40','다이어리','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','41','스터디','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','42','가계부','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','43','자기계발','0') " & vbcrlf			
			dbget.execute(strSQL)
		End If

	'### 이벤트 1
	strSQL =" INSERT INTO db_event.dbo.tbl_eventitem " &_
			" (evt_code,itemid,evtgroup_code,evtitem_sort,evtitem_linkurl) " &_
			" SELECT '" & vEventCode & "' ,i.itemid ,0, 50,null " &_
			" FROM db_item.dbo.tbl_item i " &_
			" JOIN db_diary2010.dbo.tbl_DiaryMaster d " &_
			"	on i.itemid = d.itemid " &_
			" WHERE deliverytype in (1,4) " &_
			" and d.isUsing='Y' " &_
			" and i.isUsing='Y' " &_
			" and i.itemid not in (SELECT itemid  " &_
			"	FROM db_event.dbo.tbl_eventitem " &_
			" WHERE evt_code='" & vEventCode & "') "
''' 조건삭제			" and i.sellyn in ('Y','S') " &_
	'dbget.execute(strSQL)


	'### 이벤트 2
	If vEventCode2 <> "" Then
	strSQL =" INSERT INTO db_event.dbo.tbl_eventitem " &_
			" (evt_code,itemid,evtgroup_code,evtitem_sort,evtitem_linkurl) " &_
			" SELECT '" & vEventCode2 & "' ,i.itemid ,0, 50,null " &_
			" FROM db_item.dbo.tbl_item i " &_
			" JOIN db_diary2010.dbo.tbl_DiaryMaster d " &_
			"	on i.itemid = d.itemid " &_
			" WHERE deliverytype in (1,4) " &_
			" and d.isUsing='Y' " &_
			" and i.isUsing='Y' " &_
			" and i.itemid not in (SELECT itemid  " &_
			"	FROM db_event.dbo.tbl_eventitem " &_
			" WHERE evt_code='" & vEventCode2 & "') "
''' 조건삭제			" and i.sellyn in ('Y','S') " &_
	'dbget.execute(strSQL)
	End If

	IF (vGiftCode1<>"") then

        strSQL = " insert into db_event.dbo.tbl_giftitem" & VbCRLF
        strSQL = strSQL & " (gift_code,itemid,regdate,giftitem_using)" & VbCRLF
        strSQL = strSQL & " SELECT "&vGiftCode1&" ,i.itemid ,getdate(), 1" & VbCRLF
        strSQL = strSQL & "  FROM db_item.dbo.tbl_item i " & VbCRLF
        strSQL = strSQL & "  JOIN db_diary2010.dbo.tbl_DiaryMaster d " & VbCRLF
        strSQL = strSQL & " 	on i.itemid = d.itemid " & VbCRLF
        strSQL = strSQL & "  WHERE deliverytype in (1,4) " & VbCRLF
        strSQL = strSQL & "  and d.isUsing='Y' " & VbCRLF
        strSQL = strSQL & "  and i.isUsing='Y' " & VbCRLF
        strSQL = strSQL & "  and i.itemid not in (SELECT itemid  " & VbCRLF
        strSQL = strSQL & " 	FROM db_event.dbo.tbl_giftitem " & VbCRLF
        strSQL = strSQL & "  WHERE gift_code="&vGiftCode1&")" & VbCRLF

        dbget.execute(strSQL)
	end if

	IF (vGiftCode2<>"") then

        strSQL = " insert into db_event.dbo.tbl_giftitem" & VbCRLF
        strSQL = strSQL & " (gift_code,itemid,regdate,giftitem_using)" & VbCRLF
        strSQL = strSQL & " SELECT "&vGiftCode2&" ,i.itemid ,getdate(), 1" & VbCRLF
        strSQL = strSQL & "  FROM db_item.dbo.tbl_item i " & VbCRLF
        strSQL = strSQL & "  JOIN db_diary2010.dbo.tbl_DiaryMaster d " & VbCRLF
        strSQL = strSQL & " 	on i.itemid = d.itemid " & VbCRLF
        strSQL = strSQL & "  WHERE deliverytype in (1,4) " & VbCRLF
        strSQL = strSQL & "  and d.isUsing='Y' " & VbCRLF
        strSQL = strSQL & "  and i.isUsing='Y' " & VbCRLF
        strSQL = strSQL & "  and i.itemid not in (SELECT itemid  " & VbCRLF
        strSQL = strSQL & " 	FROM db_event.dbo.tbl_giftitem " & VbCRLF
        strSQL = strSQL & "  WHERE gift_code="&vGiftCode2&")" & VbCRLF

        dbget.execute(strSQL)
	end if

	IF (vGiftCode3<>"") then

        strSQL = " insert into db_event.dbo.tbl_giftitem" & VbCRLF
        strSQL = strSQL & " (gift_code,itemid,regdate,giftitem_using)" & VbCRLF
        strSQL = strSQL & " SELECT "&vGiftCode3&" ,i.itemid ,getdate(), 1" & VbCRLF
        strSQL = strSQL & "  FROM db_item.dbo.tbl_item i " & VbCRLF
        strSQL = strSQL & "  JOIN db_diary2010.dbo.tbl_DiaryMaster d " & VbCRLF
        strSQL = strSQL & " 	on i.itemid = d.itemid " & VbCRLF
        strSQL = strSQL & "  WHERE deliverytype in (1,4) " & VbCRLF
        strSQL = strSQL & "  and d.isUsing='Y' " & VbCRLF
        strSQL = strSQL & "  and i.isUsing='Y' " & VbCRLF
        strSQL = strSQL & "  and i.itemid not in (SELECT itemid  " & VbCRLF
        strSQL = strSQL & " 	FROM db_event.dbo.tbl_giftitem " & VbCRLF
        strSQL = strSQL & "  WHERE gift_code="&vGiftCode3&")" & VbCRLF

        dbget.execute(strSQL)
	end if
	
	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)

		Alert_move msg,"/admin/diary2009/DiaryReg.asp?mode=edit&id="& Diaryid

	Else
		dbget.RollBackTrans				'롤백(에러발생시)
		Alert_return "처리중 에러가 발생했습니다."
	End If

ELSEIF mode="edit" Then

	strSQL =" UPDATE db_diary2010.dbo.tbl_DiaryMaster "&_
			" SET Cate= '"& CateCode &"'" &_
			", itemid = "& itemid &_
			", event_code = '"& event_code &"'" &_
			", eventgroup_code='" & eventgroup_code & "' " &_
			", comment_img='" & commentimgName & "' " &_
			", isUsing = '"& isUsing &"'" &_
			", BasicImg='" & basicimgName & "' " &_
			", commentyn='" & commentyn & "' " &_
			", weight='" & weight & "' " &_

			", BasicImg2 = '" & basicimg2Name & "' " & _
			", BasicImg3 = '" & basicimg3Name & "' " & _
			", StoryImg = '" & storyimgName & "' " & _
			", mdpick = '" & MDPick & "' " & _
			", limited = '" & limited & "' " & _
			", soonseo = '" & soonseo & "' " & _
			", storytext = '" & storytext & "' " & _
			", nanumimg = '" & nanumimgName & "' " & _
			", reservdate = '" & reservdate & "' " & _
			", mdpicksort = '" & mdpicksort & "' "

			if event_start <> "" then
				strSQL = strSQL & ", event_start='" & event_start & "' "
			else
				strSQL = strSQL & ", event_start=null "
			end if
			if event_end <> "" then
				 strSQL = strSQL & ", event_end='" & event_end & "' "
			else
				strSQL = strSQL & ", event_end=null "
			end if

			strSQL = strSQL & " WHERE Diaryid = "& Diaryid


	msg = "저장 되었습니다"

	dbget.execute(strSQL)


	'### 이벤트 1
	strSQL =" INSERT INTO db_event.dbo.tbl_eventitem " &_
			" (evt_code,itemid,evtgroup_code,evtitem_sort,evtitem_linkurl) " &_
			" SELECT '" & vEventCode & "' ,i.itemid ,0, 50,null " &_
			" FROM db_item.dbo.tbl_item i " &_
			" JOIN db_diary2010.dbo.tbl_DiaryMaster d " &_
			"	on i.itemid = d.itemid " &_
			" WHERE deliverytype in (1,4) " &_
			" and d.isUsing='Y' " &_
			" and i.isUsing='Y' " &_
			" and i.itemid not in (SELECT itemid  " &_
			"	FROM db_event.dbo.tbl_eventitem " &_
			" WHERE evt_code='" & vEventCode & "') "
'''	조건 삭제.		" and i.sellyn in ('Y','S') " &_

	'dbget.execute(strSQL)

	'### 이벤트 2
	If vEventCode2 <> "" Then
	strSQL =" INSERT INTO db_event.dbo.tbl_eventitem " &_
			" (evt_code,itemid,evtgroup_code,evtitem_sort,evtitem_linkurl) " &_
			" SELECT '" & vEventCode2 & "' ,i.itemid ,0, 50,null " &_
			" FROM db_item.dbo.tbl_item i " &_
			" JOIN db_diary2010.dbo.tbl_DiaryMaster d " &_
			"	on i.itemid = d.itemid " &_
			" WHERE deliverytype in (1,4) " &_
			" and d.isUsing='Y' " &_
			" and i.isUsing='Y' " &_
			" and i.itemid not in (SELECT itemid  " &_
			"	FROM db_event.dbo.tbl_eventitem " &_
			" WHERE evt_code='" & vEventCode2 & "') "
'''	조건 삭제.		" and i.sellyn in ('Y','S') " &_
	'dbget.execute(strSQL)
	End If

    IF (vGiftCode1<>"") then

        strSQL = " insert into db_event.dbo.tbl_giftitem" & VbCRLF
        strSQL = strSQL & " (gift_code,itemid,regdate,giftitem_using)" & VbCRLF
        strSQL = strSQL & " SELECT "&vGiftCode1&" ,i.itemid ,getdate(), 1" & VbCRLF
        strSQL = strSQL & "  FROM db_item.dbo.tbl_item i " & VbCRLF
        strSQL = strSQL & "  JOIN db_diary2010.dbo.tbl_DiaryMaster d " & VbCRLF
        strSQL = strSQL & " 	on i.itemid = d.itemid " & VbCRLF
        strSQL = strSQL & "  WHERE deliverytype in (1,4) " & VbCRLF
        strSQL = strSQL & "  and d.isUsing='Y' " & VbCRLF
        strSQL = strSQL & "  and i.isUsing='Y' " & VbCRLF
        strSQL = strSQL & "  and i.itemid not in (SELECT itemid  " & VbCRLF
        strSQL = strSQL & " 	FROM db_event.dbo.tbl_giftitem " & VbCRLF
        strSQL = strSQL & "  WHERE gift_code="&vGiftCode1&")" & VbCRLF

        dbget.execute(strSQL)
	end if

	IF (vGiftCode2<>"") then

        strSQL = " insert into db_event.dbo.tbl_giftitem" & VbCRLF
        strSQL = strSQL & " (gift_code,itemid,regdate,giftitem_using)" & VbCRLF
        strSQL = strSQL & " SELECT "&vGiftCode2&" ,i.itemid ,getdate(), 1" & VbCRLF
        strSQL = strSQL & "  FROM db_item.dbo.tbl_item i " & VbCRLF
        strSQL = strSQL & "  JOIN db_diary2010.dbo.tbl_DiaryMaster d " & VbCRLF
        strSQL = strSQL & " 	on i.itemid = d.itemid " & VbCRLF
        strSQL = strSQL & "  WHERE deliverytype in (1,4) " & VbCRLF
        strSQL = strSQL & "  and d.isUsing='Y' " & VbCRLF
        strSQL = strSQL & "  and i.isUsing='Y' " & VbCRLF
        strSQL = strSQL & "  and i.itemid not in (SELECT itemid  " & VbCRLF
        strSQL = strSQL & " 	FROM db_event.dbo.tbl_giftitem " & VbCRLF
        strSQL = strSQL & "  WHERE gift_code="&vGiftCode2&")" & VbCRLF

        dbget.execute(strSQL)
	end if

	IF (vGiftCode3<>"") then

        strSQL = " insert into db_event.dbo.tbl_giftitem" & VbCRLF
        strSQL = strSQL & " (gift_code,itemid,regdate,giftitem_using)" & VbCRLF
        strSQL = strSQL & " SELECT "&vGiftCode3&" ,i.itemid ,getdate(), 1" & VbCRLF
        strSQL = strSQL & "  FROM db_item.dbo.tbl_item i " & VbCRLF
        strSQL = strSQL & "  JOIN db_diary2010.dbo.tbl_DiaryMaster d " & VbCRLF
        strSQL = strSQL & " 	on i.itemid = d.itemid " & VbCRLF
        strSQL = strSQL & "  WHERE deliverytype in (1,4) " & VbCRLF
        strSQL = strSQL & "  and d.isUsing='Y' " & VbCRLF
        strSQL = strSQL & "  and i.isUsing='Y' " & VbCRLF
        strSQL = strSQL & "  and i.itemid not in (SELECT itemid  " & VbCRLF
        strSQL = strSQL & " 	FROM db_event.dbo.tbl_giftitem " & VbCRLF
        strSQL = strSQL & "  WHERE gift_code="&vGiftCode3&")" & VbCRLF

        dbget.execute(strSQL)
	end if
	
	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)
		Alert_move msg,"/admin/diary2009/DiaryReg.asp?mode=edit&id="& Diaryid
		
	Else
		dbget.RollBackTrans				'롤백(에러발생시)
		Alert_return "처리중 에러가 발생했습니다."
	End If

Elseif mode = "mdpickreg" Then

	Diaryid = split(Diaryid,",")
	MDPick = split(MDPick,",")
	cnt = ubound(Diaryid)

	For i = 0 to cnt
		strSQL =" UPDATE db_diary2010.dbo.tbl_DiaryMaster "&_
				" SET mdpick = '" & MDPick(i) & "' "
				strSQL = strSQL & " WHERE Diaryid = "& Diaryid(i)

		dbget.execute(strSQL)
	Next
		msg = "저장 되었습니다"

	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)
		Alert_move msg,"/admin/diary2009/"
		
	Else
		dbget.RollBackTrans				'롤백(에러발생시)
		Alert_return "처리중 에러가 발생했습니다."
	End If
End IF

	

'// 내지 추가 삭제
if mode_newinsert ="newinsert" Then

	if Diaryid_newinsert = "" or info_gubun_newinsert = "" or info_name_newinsert = "" then
	response.write "<script>"
	response.write "alert('코드이상..시스템팀에 문의하세요');"
	response.write "history.go(-1);"
	response.write "</script>"
	dbget.close()	:	response.End
	end if

	strSQL =" INSERT INTO [db_diary2010].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt,search_view) " &_
			" VALUES('" & Diaryid_newinsert & "','"& info_gubun_newinsert &"','" & info_name_newinsert &"','0','N')"
	response.write strSQL
	dbget.execute(strSQL)

	msg = "저장 되었습니다"

	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)
		Alert_move msg,"/admin/diary2009/option/pop_diary_info_reg.asp?diaryid="& Diaryid_newinsert

	Else
		dbget.RollBackTrans				'롤백(에러발생시)
		Alert_return "처리중 에러가 발생했습니다."
	End If

elseif mode_newinsert ="vardelete" Then

	if Diaryid_newinsert = ""  or info_gubun_delete = "" then
	response.write "<script>"
	response.write "alert('코드이상..시스템팀에 문의하세요');"
	response.write "history.go(-1);"
	response.write "</script>"
	dbget.close()	:	response.End
	end if

	strSQL =" delete from [db_diary2010].[dbo].tbl_diary_Info" &_
			" where idx = '" & Diaryid_newinsert & "' and info_gubun = '" & info_gubun_delete &"'"
	dbget.execute(strSQL)

	msg = "삭제 되었습니다"
	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)
		Alert_move msg,"/admin/diary2009/option/pop_diary_info_reg.asp?diaryid="& Diaryid_newinsert

	Else
		dbget.RollBackTrans				'롤백(에러발생시)
		Alert_return "처리중 에러가 발생했습니다."
	End If
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->