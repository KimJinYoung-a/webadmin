<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/play/play2016Cls.asp" -->
<%
	'### 기본정보 ###
	Dim i, l, tmp, vQuery, vAction, vMIdx, vDidx, vCate, vOpenDate, vState, vPartWDID, vPartMKID, vPartPBID
	vAction 	= requestCheckVar(Request("action"),10)
	vMIdx 		= requestCheckVar(Request("midx"),10)
	vDidx 		= requestCheckVar(Request("didx"),10)
	vCate		= requestCheckVar(Request("cate"),10)
	vOpenDate 	= requestCheckVar(Request("opendate"),10)
	vState 	= requestCheckVar(Request("state"),2)
	vPartMKID 	= requestCheckVar(Request("partmkid"),32)
	vPartWDID 	= requestCheckVar(Request("partwdid"),32)
	vPartPBID 	= requestCheckVar(Request("partpbid"),32)
	
	
	'### 컨텐츠 등록 - 공통부분 ###
	Dim vTitle, vSubCopy, vWorkText, vJikListImgURL, vPCIsExec, vPCExecFile, vMoIsExec, vMoExecFile, vMoBgColor, vPCBgColor, vLinkURL, vImageCopy, vSortNo
	Dim vPCContents, vMoContents, vJungListImgURL, vTitleStyle, vKeyword, vSearchListImg
	Dim vIsTagView, vTagSDate, vTagEDate, vTagAnnounceDate
	vTitle 		= html2db(requestCheckVar(Request("title"),150))
	vTitleStyle	= html2db(requestCheckVar(Request("titlestyle"),200))
	vSubCopy 		= html2db(Request("subcopy"))
	vPCContents	= html2db(Request("pc_contents"))
	vMoContents	= html2db(Request("mo_contents"))
	vPCIsExec 		= NullFillWith(requestCheckVar(Request("pc_isexec"),1),0)
	vPCExecFile	= requestCheckVar(Request("pc_execfile"),100)
	vMoIsExec 		= NullFillWith(requestCheckVar(Request("mo_isexec"),1),0)
	vMoExecFile	= requestCheckVar(Request("mo_execfile"),100)
	vMoBgColor		= requestCheckVar(Request("mo_bgcolor"),6)
	vPCBgColor		= requestCheckVar(Request("pc_bgcolor"),6)
	vLinkURL		= requestCheckVar(Request("linkurl"),6)
	vImageCopy		= html2db(requestCheckVar(Request("imagecopy"),6))
	vSortNo		= NullFillWith(requestCheckVar(Request("sortno"),6),0)
	vWorkText 		= html2db(Request("worktext"))
	vJikListImgURL = requestCheckVar(Request("jiklistimg"),100)
	vJungListImgURL = requestCheckVar(Request("junglistimg"),100)
	vSearchListImg = requestCheckVar(Request("searchlistimg"),100)
	vIsTagView		= NullFillWith(requestCheckVar(Request("istagview"),1),0)
	vTagSDate		= requestCheckVar(Request("tagsdate"),10)
	vTagEDate		= requestCheckVar(Request("tagedate"),10)
	vTagAnnounceDate	= requestCheckVar(Request("tagannouncedate"),10)
	vKeyword		= html2db(requestCheckVar(Replace(Request.Form("keyword")," ",""),300))
	If Right(vKeyword,1) = "," Then
		vKeyword = Left(vKeyword,(Len(vKeyword)-1))
	End If

	If vIsTagView = 0 Then
		vTagSDate = ""
		vTagEDate = ""
		vTagAnnounceDate = ""
	End If

	
	'### azit : cate = 3
	Dim vCate3Icon
	vCate3Icon		= requestCheckVar(Request("cate3icon"),100)

	
	If vAction = "insert" Then
		vQuery = "INSERT INTO [db_giftplus].[dbo].[tbl_play_detail](midx, cate, title, subcopy, startdate, pc_isExec, pc_execfile, "
		vQuery = vQuery & "mo_isExec, mo_execfile, mo_bgcolor, pc_bgcolor, state, pc_contents, mo_contents, "
		vQuery = vQuery & "partwdid, partmkid, partpbid, worktext, lastupdate, lastupdateID, iconimg, titlestyle, "
		vQuery = vQuery & "isTagView, tag_sdate, tag_edate, tag_announcedate, keyword "
		vQuery = vQuery & ") VALUES "
		vQuery = vQuery & "('" & vMIdx & "', '" & vCate & "', '" & vTitle & "', '" & vSubCopy & "', '" & vOpenDate & "', '" & vPCIsExec & "', '" & vPCExecFile & "', "
		vQuery = vQuery & "'" & vMoIsExec & "', '" & vMoExecFile & "', '" & vMoBGColor & "', '" & vPCBgColor & "', '" & vState & "', '" & vPCContents & "', '" & vMoContents & "', "
		vQuery = vQuery & "'" & vPartWDID & "', '" & vPartMKID & "', '" & vPartPBID & "', '" & vWorkText & "', getdate(), '" & session("ssBctId") & "', "
		vQuery = vQuery & "'" & vCate3Icon & "', '" & vTitleStyle & "', '" & vIsTagView & "', '" & vTagSDate & "', '" & vTagEDate & "', '" & vTagAnnounceDate & "', '" & vKeyword & "')"
		dbget.Execute vQuery
		
		vQuery = "select IDENT_CURRENT('db_giftplus.dbo.tbl_play_detail') as didx"
		rsget.Open vQuery, dbget, 1
		If Not Rsget.Eof then
			vDidx = rsget("didx")
		end if
		rsget.close

	ElseIf vAction = "update" Then
		vQuery = "UPDATE [db_giftplus].[dbo].[tbl_play_detail] SET "
		vQuery = vQuery & "title = '" & vTitle & "' "
		vQuery = vQuery & ", titlestyle = '" & vTitleStyle & "' "
		vQuery = vQuery & ", subcopy = '" & vSubCopy & "' "
		vQuery = vQuery & ", pc_contents = '" & vPCContents & "' "
		vQuery = vQuery & ", mo_contents = '" & vMoContents & "' "
		vQuery = vQuery & ", startdate = '" & vOpenDate & "' "
		vQuery = vQuery & ", pc_isExec = '" & vPCIsExec & "' "
		vQuery = vQuery & ", pc_execfile = '" & vPCExecFile & "' "
		vQuery = vQuery & ", mo_isExec = '" & vMoIsExec & "' "
		vQuery = vQuery & ", mo_execfile = '" & vMoExecFile & "' "
		vQuery = vQuery & ", pc_bgcolor = '" & vPCBgColor & "' "
		vQuery = vQuery & ", mo_bgcolor = '" & vMoBGColor & "' "
		vQuery = vQuery & ", state = '" & vState & "' "
		vQuery = vQuery & ", iconimg = '" & vCate3Icon & "' "
		vQuery = vQuery & ", partwdid = '" & vPartWDID & "' "
		vQuery = vQuery & ", partmkid = '" & vPartMKID & "' "
		vQuery = vQuery & ", partpbid = '" & vPartPBID & "' "
		vQuery = vQuery & ", worktext = '" & vWorkText & "' "
		vQuery = vQuery & ", isTagView = '" & vIsTagView & "' "
		vQuery = vQuery & ", tag_sdate = '" & vTagSDate & "' "
		vQuery = vQuery & ", tag_edate = '" & vTagEDate & "' "
		vQuery = vQuery & ", tag_announcedate = '" & vTagAnnounceDate & "' "
		vQuery = vQuery & ", keyword = '" & vKeyword & "' "
		vQuery = vQuery & ", lastupdate = getdate() "
		vQuery = vQuery & ", lastupdateID = '" & session("ssBctId") & "' "
		vQuery = vQuery & "where didx = '" & vDidx & "' "
		dbget.Execute vQuery

	End If
	
	'### 모든 이미지 저장 & 코너별(cate) 내용 저장
%>
	<!-- #include virtual="/admin/sitemaster/play2016/cornerproc_for_cate.asp" -->
<%
	
	Response.Write "<script>alert('처리되었습니다.');opener.location.reload();window.close();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->