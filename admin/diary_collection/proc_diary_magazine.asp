<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->

<%

'=====================================
'���� ����
'=====================================


dim mode
dim magazineid,title,isusing

dim imagename,magazinetxt,imagename2,magazinetxt2

mode= request("mode")
magazineid = request("magazineid")
title = request("title")
isusing = request("isusing")


imagename = request("imagename")
magazinetxt = request("magazinetxt")
imagename2 = request("imagename2")
magazinetxt2 = request("magazinetxt2")

dim strSQL

if mode="edit" then
	strSQL =" UPDATE [db_diary_collection].dbo.tbl_diary_magazine " &_
			" SET magazinetitle ='" & title & "'" &_
			" ,magazineImg1 ='" & imagename & "'" &_
			" ,magazineTxt1 ='" & magazinetxt & "'" &_
			" ,magazineImg2 ='" & imagename2 & "'" &_
			" ,magazineTxt2 ='" & magazinetxt2 & "'" &_
			" ,isusing='" & isusing & "'" &_
			" WHERE magazineid ='" & magazineid & "'"

else
	strSQL =" INSERT INTO db_diary_collection.dbo.tbl_diary_magazine(magazineTitle,magazineImg1,magazineTxt1,magazineImg2,magazineTxt2,isusing) " &_
			" VALUES('" & title & "','" & imagename & "','" & magazinetxt & "','" & imagename2 & "','" & magazinetxt2 & "','" & isusing & "')"
end if

	dbget.BeginTrans

	dbget.execute(strSQL)

'	if magazineid="" then
'		'strSQL =" SELECT SCOPE_IDENTITY() AS magazineid from db_diary_collection.dbo.tbl_diary_magazine "		'/������.��ü ���� ���� �ѷ���. '/2016.06.02 �ѿ��
'		strSQL =" SELECT SCOPE_IDENTITY() AS magazineid"
'
'		rsget.open strSQL,dbget,1
'
'		if not rsget.eof then
'			magazineid = rsget("magazineid")
'		end if
'
'		rsget.close
'	end if
'
'
'	strSQL = "DELETE FROM db_diary_collection.dbo.tbl_diary_magazine_sub where magazineid=" & magazineid
'
'	if imagename<>"" or magazinetxt<>"" then
'		strSQL = strSQL &_
'				" INSERT INTO db_diary_collection.dbo.tbl_diary_magazine_sub(magazineid,magazineImg,magazinetxt) " &_
'				" VALUES('" & magazineid & "','" & imagename & "','" & magazinetxt & "')"
'	end if
'
'	if imagename2<>"" or magazinetxt2<>"" then
'		strSQL = strSQL &_
'				" INSERT INTO db_diary_collection.dbo.tbl_diary_magazine_sub(magazineid,magazineImg,magazinetxt) " &_
'				" VALUES('" & magazineid & "','" & imagename2 & "','" & magazinetxt2 & "')"
'	end if
'
'
'	dbget.execute(strSQL)

	'response.write strSQL



'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)

		response.write	"<script language='javascript'>"
		response.write	" alert('����Ǿ����ϴ�.'); document.location.href='pop_diary_Magazine_List.asp'"
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