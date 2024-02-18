
<% Option Explicit %>
<% response.Charset="euc-kr" %>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	Dim i, strsql
	For i=10 To 200
		strsql = " insert into db_sitemaster.dbo.tbl_piece "
		strsql = strsql & " (fidx, gubun, noticeYN, listimg, listtext, shorttext, listtitle, adminid, usertype, etclink, snsbtncnt, itemid, pieceidx, isusing, startdate, enddate, regdate, lastupdate, Deleteyn) "
		strsql = strsql & " values "
		strsql = strsql & " ('"&i&"', '1', 'N', 'http://fiximage.10x10.co.kr/m/2017/temp/@img_1000x1266_"&i&".jpg', '조각 "&i&" 내용', '조각"&i&"', '조각"&i&" 입니다.' "
		strsql = strsql & " 	, 'thensi7', '1', '', 0, '1239298, 1239287', '', 'Y', '1900-01-01', '2099-12-31', getdate(), getdate(), 'N') "
		dbget.execute strsql
	Next
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->