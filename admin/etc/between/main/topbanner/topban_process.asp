<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/projectcls.asp"-->
<%
Dim mode, sqlStr, idx
Dim gender, pjt_kind, imgURL, linkURL, BanBGColor, partnerNmColor, BanTxtColor, bantext1, bantext2, isusing
mode			= request("mode")
idx				= request("idx")
gender			= request("gender")
pjt_kind		= request("pjt_kind")
imgURL			= request("ban")
linkURL			= request("linkURL")
BanBGColor		= request("BanBGColor")
partnerNmColor	= request("partnerNmColor")
BanTxtColor		= request("BanTxtColor")
bantext1		= request("bantext1")
bantext2		= request("bantext2")
isusing			= request("isusing")

Select Case mode
	Case "I"
		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_between_main_topbanner (gender, pjt_kind, imgURL, linkURL, BanBGColor, partnerNmColor, BanTxtColor, bantext1, bantext2, regdate, adminid, isusing) VALUES "
		sqlStr = sqlStr & " ('"&gender&"', '"&pjt_kind&"', '"&imgURL&"', '"&linkURL&"', '"&BanBGColor&"', '"&partnerNmColor&"', '"&BanTxtColor&"', '"&bantext1&"', '"&bantext2&"', getdate(), '"&session("ssBctId")&"', 'Y') "
		dbCTget.execute sqlStr
	Case "U"
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_between_main_topbanner SET "
		sqlStr = sqlStr & " gender = '"&gender&"', "
		sqlStr = sqlStr & " pjt_kind = '"&pjt_kind&"', "
		sqlStr = sqlStr & " imgURL = '"&imgURL&"', "
		sqlStr = sqlStr & " linkURL = '"&linkURL&"', "
		sqlStr = sqlStr & " BanBGColor = '"&BanBGColor&"', "
		sqlStr = sqlStr & " partnerNmColor = '"&partnerNmColor&"', "
		sqlStr = sqlStr & " BanTxtColor = '"&BanTxtColor&"', "
		sqlStr = sqlStr & " bantext1 = '"&bantext1&"', "
		sqlStr = sqlStr & " bantext2 = '"&bantext2&"', "
		sqlStr = sqlStr & " lastupdate = getdate(), "
		sqlStr = sqlStr & " lastadminid = '"&session("ssBctId")&"', "
		sqlStr = sqlStr & " isusing = '"&isusing&"' "
		sqlStr = sqlStr & " WHERE idx = '"&idx&"' "
		dbCTget.execute sqlStr
End Select
Response.Write "<script language='javascript'>alert('저장 되었습니다.');location.href='/admin/etc/between/main/topbanner/index.asp?menupos="&menupos&"';</script>"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->