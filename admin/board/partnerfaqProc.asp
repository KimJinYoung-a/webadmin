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
' Description : faq  등록
' History : 2016.08.19 정윤정  생성
'###########################################################

dim sMode
dim sTitle, ifaqType, tContents , ifaqidx
dim strparm
dim strSql
dim stType, selSearch,strSearch,iCurrpage
 
sMode		= requestCheckvar(Request("hidM"),2)

ifaqidx= requestCheckVar(Request("fidx"),10 )
ifaqType	= requestCheckVar(Request("selFaqT"),4) 
sTitle	= requestCheckVar(Request("sT"),60)  
tContents	= ReplaceRequestSpecialChar(Request("tC")) 

 '--리스트 검색 파라미터================================================================
 iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호 
 stType 		= requestCheckVar(Request("stfaqT"),4)
 selSearch = requestCheckVar(Request("selSearch"),10)
 strSearch = requestCheckVar(Request("strSearch"),200)
  
  strParm = "iC="&iCurrpage&"&selFaqT="&stType&"&selSearch="&selSearch&"&strSearch="&strSearch 
 '--================================================================
 
SELECT CASE sMode
Case "I" 
	strSql = "INSERT INTO [db_board].[dbo].[tbl_partnerA_faq] (faqtype,title,contents,regid, lastupdate) "& VBCRLF
	strSql =	strSql	&" VALUES ("&ifaqType&",'"&sTitle&"','"&tContents&"','"&session("ssBctId")&"',getdate())"
	dbget.execute strSql
	
	Response.Write "<script>alert('등록되었습니다');location.href='/admin/board/partnerfaqList.asp?menupos="&menupos&"&"&strparm&"';</script>"
	dbget.close()
	session.codePage = 949  
	Response.End 
Case "U" 
	strSql = "UPDATE [db_board].[dbo].[tbl_partnerA_faq] "& VBCRLF
	strSql = strSql & " SET faqtype="&ifaqType&", title ='"&sTitle&"',contents='"&tContents&"',regid='"&session("ssBctId")&"',lastupdate=getdate() "& VBCRLF
	strSql = strSql & " WHERE faqidx = "&ifaqidx 
	dbget.execute strSql
	
	Response.Write "<script>alert('수정되었습니다');location.href='/admin/board/partnerfaqReg.asp?menupos="&menupos&"&fidx="&ifaqidx&"';</script>"
	dbget.close()
	session.codePage = 949  
	Response.End 
Case "D" 
	strSql = "UPDATE [db_board].[dbo].[tbl_partnerA_faq] "& VBCRLF
	strSql = strSql & " SET isusing=0,regid='"&session("ssBctId")&"',lastupdate=getdate() "& VBCRLF
	strSql = strSql & " WHERE faqidx = "&ifaqidx
	dbget.execute strSql
	
	Response.Write "<script>alert('삭제되었습니다');location.href='/admin/board/partnerfaqList.asp?menupos="&menupos&"&"&strparm&"';</script>"
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