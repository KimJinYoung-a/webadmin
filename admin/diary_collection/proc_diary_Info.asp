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




dim diaryid,mode, infoname , infogubun ,infoImage ,infocnt
diaryid = request("diaryid")
mode= request("mode")

If diaryid="" Then

	response.write "���"
	dbget.close()	:	response.End
End If
'/ ���� ����

infoname = request("infoname")
infogubun = request("infogubun")
infoImage = request("infoImage")
infocnt = request("infocnt")

infoname= split(infoname,",")
infogubun = split(infogubun,",")
infoImage= split(infoImage,",")
infocnt= split(infocnt,",")

'/���� ����
dim TotalPageName,TotalPagepageCnt,etcname

TotalPageName = request("TotalPageName")
TotalPagepageCnt = request("TotalPagepageCnt")
etcname= request("etcname")



dim strSQL,i

dbget.beginTrans

	'/ ���� �κ�
	For i=0 to ubound(infoname)
		strSQL=	strSQL &_
				" UPDATE [db_diary_collection].[dbo].tbl_diary_Info "&_
				" SET Info_Name ='" & infoname(i) & "' " &_
				" ,info_img ='" & infoImage(i) & "'" &_
				" ,info_PageCnt ='" & infocnt(i) & "'" &_
				" WHERE idx='" & diaryid & "' and info_gubun='" & infogubun(i) & "'"

		IF infogubun(i)="14" then
			if infoImage(i)<>"" then
				strSQL=	strSQL &_
						" UPDATE [db_diary_collection].[dbo].tbl_diary_master "&_
						" SET giftYn='Y'" &_
						" WHERE idx='" & diaryid & "'"
			Else
				strSQL=	strSQL &_
						" UPDATE [db_diary_collection].[dbo].tbl_diary_master "&_
						" SET giftYn='N'" &_
						" WHERE idx='" & diaryid & "'"
			end if
		end if

	Next


	dbget.execute(strSQL)

	'// �������� TotalPages �κ� ���� ������ 15
	if trim(TotalPageName)<>"" then

		strSQL=	" UPDATE [db_diary_collection].[dbo].tbl_diary_Info "&_
				" SET Info_Name ='" & html2db(TotalPageName) & "' " &_
				" ,info_PageCnt ='" & TotalPagepageCnt & "'" &_
				" WHERE idx='" & diaryid & "' and info_gubun='15'"
		dbget.execute(strSQL)


	end if
	'// �������� ��Ÿ ����  �κ� ���� ������ 16
	strSQL=	" UPDATE [db_diary_collection].[dbo].tbl_diary_Info "&_
			" SET Info_Name ='" & html2db(etcname) & "' " &_
			" WHERE idx='" & diaryid & "' and info_gubun='16'"

	dbget.execute(strSQL)


If Err.Number = 0 Then
	dbget.CommitTrans

else
	dbget.RollbackTrans
End If

response.write "<script language='javascript'>alert('�����Ͽ����ϴ�.')</script>"
response.write "<script language='javascript'>document.location.replace('" &Referer &"');</script>"
dbget.close()	:	response.End


%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
