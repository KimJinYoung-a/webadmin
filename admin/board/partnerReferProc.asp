<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###########################################################
' Description : 매뉴얼  등록
' History : 2016.08.19 정윤정  생성
'###########################################################

dim sMode
dim sTitle, irefType, tContents , irefidx
dim fileName,tmpeCode, i
dim strparm
dim strSql
dim stType, selSearch,strSearch,iCurrpage
sMode		= requestCheckvar(Request("hidM"),2)

irefidx= requestCheckVar(Request("fidx"),10 )
irefType	= requestCheckVar(Request("selrefT"),4) 
sTitle	= requestCheckVar(Request("sT"),60)  
tContents	= ReplaceRequestSpecialChar(Request("tC")) 
fileName 	= ReplaceRequestSpecialChar(Request("sFileP")) 

 '--리스트 검색 파라미터================================================================
 iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호 
 stType 		= requestCheckVar(Request("strefT"),4)
 selSearch = requestCheckVar(Request("selSearch"),10)
 strSearch = requestCheckVar(Request("strSearch"),200)
  
  strParm = "iC="&iCurrpage&"&selrefT="&stType&"&selSearch="&selSearch&"&strSearch="&strSearch
 '--================================================================
  
SELECT CASE sMode
Case "I" 
	strSql = "INSERT INTO [db_board].[dbo].[tbl_partnerA_reference] (reftype,title,contents,regid, lastupdate) "& VBCRLF
	strSql =	strSql	&" VALUES ("&irefType&",'"&sTitle&"','"&tContents&"','"&session("ssBctId")&"',getdate())" 
	dbget.execute strSql
	
	IF fileName <> "" then
	strSql = "select SCOPE_IDENTITY()" 
	rsget.Open strSql, dbget, 0
	tmpeCode = rsget(0)
	rsget.Close
	 
	
	'첨부파일 등록
		fileName = split(fileName,",")
		For i = 0 To UBound(fileName)
		if (trim(fileName(i)) <> "") then
			strSql = " INSERT INTO db_board.dbo.tbl_partnerA_reference_attachfile(refIdx, fileLink) "
			strSql = strSql & " VALUES ("&tmpeCode&",'"&trim(fileName(i))&"' ) "
			dbget.execute strSql
		end if
		Next
	end if	 
	Response.Write "<script>alert('등록되었습니다');location.href='/admin/board/partnerReferList.asp?menupos="&menupos&"&"&strparm&"';</script>"
	dbget.close()
	session.codePage = 949  
	Response.End 
Case "U" 
	strSql = "UPDATE [db_board].[dbo].[tbl_partnerA_reference]"& VBCRLF
	strSql = strSql & " SET reftype="&irefType&", title ='"&sTitle&"',contents='"&tContents&"',regid='"&session("ssBctId")&"',lastupdate=getdate() "& VBCRLF
	strSql = strSql & " WHERE refidx = "&irefidx 
	dbget.execute strSql
	
	IF fileName <> "" then
	'첨부파일 등록
	strSql = "DELETE FROM db_board.dbo.tbl_partnerA_reference_attachfile	where refidx = "&irefidx
	dbget.execute strSql
	
		fileName = split(fileName,",")
		For i = 0 To UBound(fileName)
		if (trim(fileName(i)) <> "") then
			strSql = " INSERT INTO db_board.dbo.tbl_partnerA_reference_attachfile(refIdx, fileLink) "
			strSql = strSql & " VALUES ("&irefidx&",'"&trim(fileName(i))&"' ) "
			dbget.execute strSql
		end if
		Next
	END IF	
		
	Response.Write "<script>alert('수정되었습니다');location.href='/admin/board/partnerReferReg.asp?menupos="&menupos&"&fidx="&irefidx&"';</script>"
	dbget.close()
	session.codePage = 949  
	Response.End 
Case "D" 
	strSql = "UPDATE [db_board].[dbo].[tbl_partnerA_reference]"& VBCRLF
	strSql = strSql & " SET isusing=0,regid='"&session("ssBctId")&"',lastupdate=getdate() "& VBCRLF
	strSql = strSql & " WHERE refidx = "&irefidx
	dbget.execute strSql
	
	 	strSql = "DELETE FROM db_board.dbo.tbl_partnerA_reference_attachfile	where refidx = "&irefidx
	dbget.execute strSql
	
	Response.Write "<script>alert('삭제되었습니다');location.href='/admin/board/partnerReferList.asp?menupos="&menupos&"&"&strparm&"';</script>"
	dbget.close()
	session.codePage = 949  
	Response.End 	 
CASE ELSE
	Response.Write "<script>alert('데이터 처리에 문제가 발생하였습니다.');history.back();</script>"
	session.codePage = 949
	response.end
END SELECT
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	session.codePage = 949
%>
%>