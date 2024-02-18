<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->
<%

'=====================================
'본문 시작
'=====================================

dim diaryid,cont_idx,cont_text,contImgName,mode
dim sql
diaryid = request("diaryid")
cont_idx =request("cont_idx")
cont_text = request("cont_text")
contImgName = request("contImgName")
mode = request("mode")

'===========	신규 등록 시작	===========
if mode="write" then
	'// 데이터 삽입

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

	'//  방금 등록된 cont_idx 값을 가져온다
	cont_idx = rsget("cont_idx")

	rsget.close

'===========	수정 시작		===========
elseif mode="modify" then

	rsget.Source=" select top 1 cont_text from [db_diary2010].dbo.tbl_diary_content where cont_idx=" & cont_idx

	rsget.ActiveConnection= dbget
	rsget.Cursortype=adOpenStatic
	rsget.LockType=adLockOptimistic

	rsget.open
		rsget.Fields("cont_text")		=	cont_text
	rsget.update


	rsget.close
'===========	수정 끝		===========

'===========	삭제 시작 ===========
elseif mode="del" then
	'response.write cont_idx
'dbget.close()	:	response.End
	sql = "delete from [db_diary2010].dbo.tbl_diary_content where cont_idx= "& cont_idx&""
	
	response.write sql
	dbget.execute sql



	response.write "<script language='javascript'>alert('삭제되었습니다.')</script>"
	response.write "<script language='javascript'>location.replace('pop_diary_cont_reg.asp?diaryid=" & diaryid & "')</script>"
	dbget.close()	:	response.End
'===========	삭제 끝 ===========
end if


response.write "<script language='javascript'>alert('저장하였습니다.')</script>"
response.write "<script language='javascript'>location.replace('pop_diary_cont_reg.asp?diaryid=" & diaryid & "')</script>"
dbget.close()	:	response.End


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->