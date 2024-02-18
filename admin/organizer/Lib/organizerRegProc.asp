<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<!-- #include virtual="/admin/organizer/Lib/include_event_code.asp"-->
<%
'// 오거나이저 저장폼

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
dim organizer_order

mode= request.Form("mode")
CateCode= request.Form("cate")
isUsing= request.Form("ius")
itemid= request.Form("iid")
Diaryid = request.Form("did")
Weight = request.Form("wt")
basicimgName = request.Form("basicimgName")
commentyn = request.Form("commentyn")

dim Diaryid_newinsert , info_gubun_newinsert ,info_name_newinsert , mode_newinsert , info_gubun_delete , eventgroup_code 
dim event_code , commentimgName , event_start , event_end , color
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
	color = request("color")
	organizer_order = request("organizer_order")	
	'response.write color
	'dbget.close()	:	response.End
dim strSQL,msg

dbget.beginTrans		 
IF mode="add" Then
	strSQL =" INSERT INTO db_diary2010.[dbo].tbl_organizermaster "&_
			" (Cate,Itemid,organizer_order,BasicImg,isusing,commentyn,event_code,eventgroup_code,comment_img ,weight ,event_start,event_end,color) "&_
			" VALUES(" &_
			"'" & CateCode & "' " &_
			",'" & itemid & "' " &_
			",'" & organizer_order & "' " &_
			",'" & basicimgName & "' " &_
			",'" & isUsing & "' " &_
			",'" & commentyn & "' " &_
			",'" & event_code & "' " &_
			",'" & eventgroup_code & "' " &_
			",'" & commentimgName & "' " &_
			",'" & weight & "' " 
			
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

				strSQL = strSQL & ",'" & color & "' " 				
			strSQL = strSQL & ")"
			
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

			strSQL =" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','1','yearly','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','2','monthly','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','3','weekly','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','4','daily','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','5','free note','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','6','line/square note','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','7','pocket/case','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','8','function note','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','9','check list','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','10','review','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','11','calendar','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','12','18개월','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','13','only 2011','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','14','만년','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','15','gift','0')"

			dbget.execute(strSQL)
		End If
	
'		strSQL =" INSERT INTO db_event.dbo.tbl_eventitem " &_
'				" (evt_code,itemid,evtgroup_code,evtitem_sort,evtitem_linkurl) " &_
'				" SELECT '" & vEventCode & "' ,i.itemid ,0, 50,null " &_
'				" FROM db_item.dbo.tbl_item i " &_
'				" JOIN db_diary2010.dbo.tbl_organizerMaster d " &_
'				"	on i.itemid = d.itemid " &_
'				" WHERE deliverytype in (1,4) " &_
'				" and d.isUsing='Y' " &_
'				" and i.isUsing='Y' " &_
'				" and i.itemid not in (SELECT itemid  " &_
'				"	FROM db_event.dbo.tbl_eventitem " &_
'				" WHERE evt_code='" & vEventCode & "') " 
'
'		dbget.execute(strSQL)
	
	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)
		Alert_move msg,"/admin/organizer/organizerReg.asp?mode=edit&id="& Diaryid
		
	Else
		dbget.RollBackTrans				'롤백(에러발생시)
		Alert_return "처리중 에러가 발생했습니다."
	End If
	

ELSEIF mode="edit" Then
'response.write organizer_order
	strSQL =" UPDATE db_diary2010.dbo.tbl_organizermaster "&_
			" SET Cate= '"& CateCode &"'" &_
			", itemid = "& itemid &_
			", event_code = '"& event_code &"'" &_
			", organizer_order = '"& organizer_order &"'" &_
			", eventgroup_code='" & eventgroup_code & "' " &_
			", comment_img='" & commentimgName & "' " &_
			", isUsing = '"& isUsing &"'" &_
			", BasicImg='" & basicimgName & "' " &_
			", commentyn='" & commentyn & "' " &_
			", weight='" & weight & "' "

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
			
			strSQL = strSQL & ", color='" & color & "' "	
					
			strSQL = strSQL & " WHERE organizerid = "& Diaryid
			
	msg = "저장 되었습니다"

	dbget.execute(strSQL)
	
'	strSQL =" INSERT INTO db_event.dbo.tbl_eventitem " &_
'			" (evt_code,itemid,evtgroup_code,evtitem_sort,evtitem_linkurl) " &_
'			" SELECT '" & vEventCode & "' ,i.itemid ,0, 50,null " &_
'			" FROM db_item.dbo.tbl_item i " &_
'			" JOIN db_diary2010.dbo.tbl_organizerMaster d " &_
'			"	on i.itemid = d.itemid " &_
'			" WHERE deliverytype in (1,4) " &_
'			" and d.isUsing='Y' " &_
'			" and i.isUsing='Y' " &_
'			" and i.itemid not in (SELECT itemid  " &_
'			"	FROM db_event.dbo.tbl_eventitem " &_
'			" WHERE evt_code='" & vEventCode & "') " 
'response.write strSQL
'	dbget.execute(strSQL)
	
	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)
		Alert_move msg,"/admin/organizer/organizerReg.asp?mode=edit&id="& Diaryid
		
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

	strSQL =" INSERT INTO [db_diary2010].[dbo].tbl_organizer_Info (idx ,info_gubun,info_name,info_pageCnt,search_view) " &_
			" VALUES('" & Diaryid_newinsert & "','"& info_gubun_newinsert &"','" & info_name_newinsert &"','0','N')"	
	response.write strSQL
	dbget.execute(strSQL)

	msg = "저장 되었습니다"
	
	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)
		Alert_move msg,"/admin/organizer/option/pop_organizer_info_reg.asp?diaryid="& Diaryid_newinsert
		
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

	strSQL =" delete from [db_diary2010].[dbo].tbl_organizer_Info" &_
			" where idx = '" & Diaryid_newinsert & "' and info_gubun = '" & info_gubun_delete &"'"	
	dbget.execute(strSQL)

	msg = "삭제 되었습니다"
	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)
		Alert_move msg,"/admin/organizer/option/pop_organizer_info_reg.asp?diaryid="& Diaryid_newinsert
		
	Else
		dbget.RollBackTrans				'롤백(에러발생시)
		Alert_return "처리중 에러가 발생했습니다."
	End If	
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->