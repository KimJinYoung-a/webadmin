<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<!-- #include virtual="/admin/diary2009/lib/include_event_code.asp"-->

<%
'// ���̾ ������
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

	msg = "���� �Ǿ����ϴ�"

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

			'2019 ���̾ ���� ����
			strSQL = " INSERT INTO [db_diary2010].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " & vbcrlf
			strSQL = strSQL & " VALUES " & vbcrlf
			'strSQL = strSQL & "('" & idx & "','22','1����','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','23','�б⺰','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','24','6����','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','25','1��','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','26','1�� �̻�','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','27','����������','0') ," & vbcrlf
			'strSQL = strSQL & "('" & idx & "','28','����������','0') ," & vbcrlf
			'strSQL = strSQL & "('" & idx & "','29','�ְ�������','0') ," & vbcrlf
			'strSQL = strSQL & "('" & idx & "','30','�Ͻ�����','0') ," & vbcrlf
			'strSQL = strSQL & "('" & idx & "','31','ĳ�ú�','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','32','����','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','33','���','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','34','��Ȧ��','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','35','������','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','36','2019 ��¥��','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','37','�ս���','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','38','��Ŭ��','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','39','���ϸ�','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','40','���̾','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','41','���͵�','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','42','�����','0') ," & vbcrlf
			strSQL = strSQL & "('" & idx & "','43','�ڱ���','0') " & vbcrlf			
			dbget.execute(strSQL)
		End If

	'### �̺�Ʈ 1
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
''' ���ǻ���			" and i.sellyn in ('Y','S') " &_
	'dbget.execute(strSQL)


	'### �̺�Ʈ 2
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
''' ���ǻ���			" and i.sellyn in ('Y','S') " &_
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
	
	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)

		Alert_move msg,"/admin/diary2009/DiaryReg.asp?mode=edit&id="& Diaryid

	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)
		Alert_return "ó���� ������ �߻��߽��ϴ�."
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


	msg = "���� �Ǿ����ϴ�"

	dbget.execute(strSQL)


	'### �̺�Ʈ 1
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
'''	���� ����.		" and i.sellyn in ('Y','S') " &_

	'dbget.execute(strSQL)

	'### �̺�Ʈ 2
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
'''	���� ����.		" and i.sellyn in ('Y','S') " &_
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
	
	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)
		Alert_move msg,"/admin/diary2009/DiaryReg.asp?mode=edit&id="& Diaryid
		
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)
		Alert_return "ó���� ������ �߻��߽��ϴ�."
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
		msg = "���� �Ǿ����ϴ�"

	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)
		Alert_move msg,"/admin/diary2009/"
		
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)
		Alert_return "ó���� ������ �߻��߽��ϴ�."
	End If
End IF

	

'// ���� �߰� ����
if mode_newinsert ="newinsert" Then

	if Diaryid_newinsert = "" or info_gubun_newinsert = "" or info_name_newinsert = "" then
	response.write "<script>"
	response.write "alert('�ڵ��̻�..�ý������� �����ϼ���');"
	response.write "history.go(-1);"
	response.write "</script>"
	dbget.close()	:	response.End
	end if

	strSQL =" INSERT INTO [db_diary2010].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt,search_view) " &_
			" VALUES('" & Diaryid_newinsert & "','"& info_gubun_newinsert &"','" & info_name_newinsert &"','0','N')"
	response.write strSQL
	dbget.execute(strSQL)

	msg = "���� �Ǿ����ϴ�"

	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)
		Alert_move msg,"/admin/diary2009/option/pop_diary_info_reg.asp?diaryid="& Diaryid_newinsert

	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)
		Alert_return "ó���� ������ �߻��߽��ϴ�."
	End If

elseif mode_newinsert ="vardelete" Then

	if Diaryid_newinsert = ""  or info_gubun_delete = "" then
	response.write "<script>"
	response.write "alert('�ڵ��̻�..�ý������� �����ϼ���');"
	response.write "history.go(-1);"
	response.write "</script>"
	dbget.close()	:	response.End
	end if

	strSQL =" delete from [db_diary2010].[dbo].tbl_diary_Info" &_
			" where idx = '" & Diaryid_newinsert & "' and info_gubun = '" & info_gubun_delete &"'"
	dbget.execute(strSQL)

	msg = "���� �Ǿ����ϴ�"
	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)
		Alert_move msg,"/admin/diary2009/option/pop_diary_info_reg.asp?diaryid="& Diaryid_newinsert

	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)
		Alert_return "ó���� ������ �߻��߽��ϴ�."
	End If
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->