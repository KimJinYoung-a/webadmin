<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->

<%

public function chkarray(strArr)
	dim tmp
	dim tmparray
	dim intLoop

	if (len(replace(strArr,",",""))<1) or (len(trim(strArr))<1) then

		exit function
	end if

	tmparray = split(strArr,",")

	for intLoop = 0 to ubound(tmparray)

		if trim(tmparray(intLoop)) <>"" then
			tmp = tmp  & tmparray(intLoop) & ","
		end if
	next
	chkarray = left(tmp,len(tmp)-1)
end function
dim Referer
Referer = Request.ServerVariables("HTTP_REFERER")




dim idx,mode, infoname , infogubun ,infoImage ,infocnt
idx = request("idx")
mode= request("mode")

If Idx="" Then

	response.write "경고"
	dbget.close()	:	response.End
End If
'/ 공통 내용

infoname = request("infoname")
infogubun = request("infogubun")
infoImage = request("infoImage")
infocnt = request("infocnt")

infoname= split(infoname,",")
infogubun = split(infogubun,",")
infoImage= split(infoImage,",")
infocnt= split(infocnt,",")

'/고정 내용
dim TotalPageName,TotalPagepageCnt,etcname

TotalPageName = request("TotalPageName")
TotalPagepageCnt = request("TotalPagepageCnt")
etcname= request("etcname")



dim strSQL,i

dbget.beginTrans

	'/ 공통 부분
	For i=0 to ubound(infoname)
	strSQL=	strSQL &_
			" UPDATE [db_diary_collection].[dbo].tbl_diary_Info "&_
			" SET Info_Name ='" & infoname(i) & "' " &_
			" ,info_img ='" & infoImage(i) & "'" &_
			" ,info_PageCnt ='" & infocnt(i) & "'" &_
			" WHERE idx='" & Idx & "' and info_gubun='" & infogubun(i) & "'"

	Next

	dbget.execute(strSQL)

	'// 내지구성 TotalPages 부분 저장 고정값 15
	if trim(TotalPageName)<>"" then

		strSQL=	" UPDATE [db_diary_collection].[dbo].tbl_diary_Info "&_
				" SET Info_Name ='" & html2db(TotalPageName) & "' " &_
				" ,info_PageCnt ='" & TotalPagepageCnt & "'" &_
				" WHERE idx='" & Idx & "' and info_gubun='15'"
		dbget.execute(strSQL)


	end if
	'// 내지구성 기타 사항  부분 저장 고정값 16
	strSQL=	" UPDATE [db_diary_collection].[dbo].tbl_diary_Info "&_
			" SET Info_Name ='" & html2db(etcname) & "' " &_
			" WHERE idx='" & Idx & "' and info_gubun='16'"

	dbget.execute(strSQL)


If Err.Number = 0 Then
	dbget.CommitTrans

else
	dbget.RollbackTrans
End If

response.write "<script language='javascript'>alert('적용하였습니다.')</script>"
response.write "<script language='javascript'>document.location.replace('" &Referer &"');</script>"
dbget.close()	:	response.End


%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

