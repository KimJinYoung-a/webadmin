<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->

<%
dim YearUse,diaryType,itemid,hitYn,giftYn,onlyyearYn,freeBaeSongYn,isusing,basicimgName,idx,mode
YearUse = request("YearUse")
diaryType = request("diaryType")
itemid = request("itemid")
hitYn = request("hitYn")
giftYn = request("giftYn")
onlyyearYn = request("onlyyearYn")
freeBaeSongYn = request("freeBaeSongYn")
isusing = request("isusing")
basicimgName = request("basicimgName")
idx =request("idx")
mode = request("mode")
if hitYn="on" then hitYn="Y" else hitYn ="N" end if
if giftYn="on" then giftYn="Y" else giftYn ="N" end if
if onlyyearYn="on" then onlyyearYn="Y" else onlyyearYn ="N" end if



dim strSQL,msg

dbget.begintrans

	if mode="edit" then

		strSQL =" UPDATE [db_diary_collection].[dbo].tbl_diary_master " &_
				" SET diaryType='" & diaryType & "' " &_
				" ,itemid='" & itemid & "' " &_
				" ,isusing='" & isusing & "' " &_
				" ,giftYn='" & giftYn & "' " &_
				" ,onlyYearYn='" & onlyYearYn & "' " &_
				" ,hitYn='" & hitYn & "' " &_
				" ,basic_img='" & basicimgName & "' " &_
				" ,list_img='" & basicimgName & "' " &_
				" ,icon_img='" & basicimgName & "' " &_
				" WHERE idx ='" & idx & "'"
		msg = "���� �Ǿ����ϴ�"

		'response.write strSQL
		dbget.execute(strSQL)

	else
		strSQL =" INSERT INTO [db_diary_collection].[dbo].tbl_diary_master(yearuse,diaryType,itemid,isusing,giftYn,onlyYearYn,hitYn,basic_img,list_img,icon_img) " &_
				" VALUES(" &_
				"'" & YearUse & "', " &_
				"'" & diaryType & "', " &_
				"'" & itemid & "', " &_
				"'" & isusing & "', " &_
				"'" & giftYn & "', " &_
				"'" & onlyYearYn & "', " &_
				"'" & hitYn & "', " &_
				"'" & basicimgName & "', " &_
				"'" & basicimgName & "', " &_
				"'" & basicimgName & "' " &_
				")"
		msg = "���� �Ǿ����ϴ�"

		dbget.execute(strSQL)

		'// ���̾ ���������� �⺻Ʋ ����

		'strSQL=" select Scope_identity() as idx from [db_diary_collection].[dbo].tbl_diary_master"		'/������.��ü ���� ���� �ѷ���. '/2016.06.02 �ѿ��
		strSQL=" select Scope_identity() as idx"

		rsget.open strSQL,dbget,1

		If not rsget.eof Then
			idx = rsget("idx")
		End If
		rsget.close

		If idx<>"" Then

			strSQL =" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','1','2008 calendar','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','2','yearly plan','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','3','monthly','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','4','weekly','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','5','daily','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','6','my list','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','7','culture&movie','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','8','freenote','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','9','cash pages','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','10','address','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','11','profile','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','12','book\0x2Fweb reference','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','13','bookbinding(��������)','0')"

			if giftYn="Y" then
				strSQL = strSQL &_
					" UPDATE	[db_diary_collection].[dbo].tbl_diary_master " &_
					" SET giftYN='Y'" &_
					" where idx='" & idx  & "'"
			end if

			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','14','gift(����ǰ ����)','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','15','TotalPages','0')"
			strSQL = strSQL &_
					" INSERT INTO [db_diary_collection].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " &_
					" VALUES('" & idx & "','16','','0')"
			dbget.execute(strSQL)
		End If

		strSQL =" exec db_diary_collection.dbo.ten_IMSI_diary_eventPrize"

		dbget.execute(strSQL)

	end if


	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)

		response.write	"<script language='javascript'>"
		response.write	" alert('" & msg & "'); opener.location.reload(true);self.close();"
		response.write	"</script>"
		dbget.close()	:	response.End
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)

		response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.go(-1);" &_
					"</script>"


	End If

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->