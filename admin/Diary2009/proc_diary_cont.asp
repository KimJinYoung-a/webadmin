<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->
<%

'=====================================
'���� ����
'=====================================

dim diaryid,cont_idx,cont_text,contImgName,mode
dim sql
diaryid = request("diaryid")
cont_idx =request("cont_idx")
cont_text = request("cont_text")
contImgName = request("contImgName")
mode = request("mode")

'===========	�ű� ��� ����	===========
if mode="write" then
	'// ������ ����

	'rsget.open "[db_contents].[dbo].tbl_diary_master", adoCon, adOpenKeyset, adLockPessimistic, adCmdTable

	rsget.Source	= "select cont_file,cont_idx,idx,cont_text from [db_diary2010].[dbo].tbl_diary_content where 1=0"

	rsget.ActiveConnection=dbget
	rsget.CursorType=adOpenKeyset
	rsget.LockType=adLockPessimistic

	rsget.Open
		rsget.AddNew
		rsget.Fields("cont_file") 	= contImgName
		rsget.Fields("idx")			= diaryid
		rsget.Fields("cont_text")	= cont_text
	rsget.update

	'//  ��� ��ϵ� cont_idx ���� �����´�
	cont_idx = rsget("cont_idx")

	rsget.close

'===========	���� ����		===========
elseif mode="modify" then

	rsget.Source=" select top 1 cont_text from [db_diary2010].dbo.tbl_diary_content where cont_idx=" & cont_idx

	rsget.ActiveConnection= dbget
	rsget.Cursortype=adOpenStatic
	rsget.LockType=adLockOptimistic

	rsget.open
		rsget.Fields("cont_text")		=	cont_text
	rsget.update


	rsget.close
'===========	���� ��		===========

'===========	���� ���� ===========
elseif mode="del" then
	'response.write cont_idx
'dbget.close()	:	response.End
	sql = "delete from [db_diary2010].dbo.tbl_diary_content where cont_idx= "& cont_idx&""
	
	response.write sql
	dbget.execute sql



	response.write "<script language='javascript'>alert('�����Ǿ����ϴ�.')</script>"
	response.write "<script language='javascript'>location.replace('pop_diary_cont_reg.asp?diaryid=" & diaryid & "')</script>"
	dbget.close()	:	response.End
'===========	���� �� ===========
end if


response.write "<script language='javascript'>alert('�����Ͽ����ϴ�.')</script>"
response.write "<script language='javascript'>location.replace('pop_diary_cont_reg.asp?diaryid=" & diaryid & "')</script>"
dbget.close()	:	response.End


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->