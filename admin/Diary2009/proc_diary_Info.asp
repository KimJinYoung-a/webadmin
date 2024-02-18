<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->

<%

dim diaryid, infoImage ,mode ,infocnt  ,image_count ,idx , contents_idx
	diaryid = request("diaryid")
	infoImage = request("infoImage")
	mode = request("mode")	
	infocnt = request("infocnt")	
	idx = request("idx")	
	image_count = request("image_count")
	contents_idx = request("contents_idx")

	'이미지와 내지갯수 배열선언
	infoImage= split(infoImage,",")
	infocnt= split(infocnt,",")


dim Referer
	Referer = Request.ServerVariables("HTTP_REFERER")

dim strSQL , i



dbget.beginTrans

	'/ 공통 부분
	For i=0 to ubound(infoImage) -1
		response.write image_count &"<br>"
		response.write i &"<br>"
		response.write infocnt(i) &"<br>"		
		if cint(image_count) = cint(i) then
		strSQL = ""	
		strSQL = "UPDATE [db_diary2010].dbo.tbl_contents_master set" + vbcrlf
		strSQL = strSQL & " diaryid = "& diaryid &"," + vbcrlf
		strSQL = strSQL & " image = '"& infoImage(i) &"'," + vbcrlf
		strSQL = strSQL & " PageCnt = "& infocnt(i) &"" + vbcrlf
		strSQL = strSQL & " where idx = "& idx &"" + vbcrlf
		
		end if
		
		response.write strSQL&"<br>"	
	Next


	dbget.execute strSQL


If Err.Number = 0 Then
	dbget.CommitTrans

else
	dbget.RollbackTrans
End If

'response.write "<script language='javascript'>alert('적용하였습니다.')</script>"
response.write "<script language='javascript'>document.location.replace('" &Referer &"');</script>"
dbget.close()	:	response.End


%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
