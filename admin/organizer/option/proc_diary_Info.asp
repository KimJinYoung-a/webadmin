<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->
<%
'#######################################################
'	History	:  2008.10.23 한용민 생성
'	Description : 오거나이저
'#######################################################
%>
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
		strSQL = "UPDATE [db_diary2010].[dbo].tbl_organizer_Info" +vbcrlf
		strSQL = strSQL & " SET Info_Name ='" & infoname(i) & "' " +vbcrlf
		strSQL = strSQL & "  ,info_img ='" & html2db(infoImage(i)) & "'" +vbcrlf
		strSQL = strSQL & "  ,info_PageCnt ='" & infocnt(i) & "'" +vbcrlf
		strSQL = strSQL & "  WHERE idx='" & diaryid & "' and info_gubun='" & infogubun(i) & "'" +vbcrlf
		
		'response.write strSQL &"공통부분<br>"
		dbget.execute strSQL
	Next

	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)	
	Else
		dbget.RollBackTrans				'롤백(에러발생시)
	End If

response.write "<script language='javascript'>alert('적용하였습니다.')</script>"
response.write "<script language='javascript'>document.location.replace('" &Referer &"');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
