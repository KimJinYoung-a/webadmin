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
<!-- #include virtual="/lib/classes/search/search_manageCls.asp"-->
<%
	Dim i, vAction, vQuery, vUserID, vIdx, vBgGubun, vBgColor, vBgImg, c, vMaskingImg, vViewGubun, vSDate, vEDate, vUseYN
	Dim vTextInfoUse, vTextInfo1, vTextInfo1url, vTextInfo2, vTextInfo2url, vMemo, vRegdate, vLastUserName, vLastdate, vRegUserName
	Dim vShhmmss, vEhhmmss
	
	'### 검색화면 기본 정보
	vIdx 		= requestCheckVar(Request("idx"),15)
	vAction	= requestCheckVar(Request("action"),10)
	vUserID	= session("ssBctId")
	
	vBgGubun	= requestCheckVar(Request("bggubun"),1)
	If vBgGubun = "c" Then	'### 퀵링크배경 단색
		vBgImg = ""
		vBgColor	= requestCheckVar(Request("bgcolor"),6)
	ElseIf vBgGubun = "i" Then	'### 퀵링크배경 이미지사용
		vBgImg	= requestCheckVar(Request("mbgimgurlm"),100)
		vBgColor	= ""
	End If
	
	vMaskingImg = requestCheckVar(Request("maskingimgurlm"),100)
	vViewGubun	= requestCheckVar(Request("viewgubun"),6)
	If vViewGubun = "always" Then	'### 상시노출
		vSDate = ""
		vEDate = ""
	Else	'### 기간설정
		vSDate = requestCheckVar(Request("sdate"),10)
		vEDate = requestCheckVar(Request("edate"),10)
		vShhmmss = requestCheckVar(Request("shhmmss"),8)
		vEhhmmss = requestCheckVar(Request("ehhmmss"),8)
		
		vSDate = vSDate & " " & vShhmmss
		vEDate = vEDate & " " & vEhhmmss
	End If
	vUseYN		= requestCheckVar(Request("useyn"),1)
	vMemo 		= html2db(Request("memo"))


	'### 검색화면 기본 정보
	vTextInfoUse	= requestCheckVar(Request("textinfouse"),1)
	If vTextInfoUse <> "0" Then
		If vTextInfoUse = "1" Then
			vTextInfo1		= requestCheckVar(Request("textinfo1"),20)
			vTextInfo1url	= requestCheckVar(Request("textinfo1url"),200)
		ElseIf vTextInfoUse = "2" Then
			vTextInfo1		= requestCheckVar(Request("textinfo1"),20)
			vTextInfo1url	= requestCheckVar(Request("textinfo1url"),200)
			vTextInfo2		= requestCheckVar(Request("textinfo2"),20)
			vTextInfo2url	= requestCheckVar(Request("textinfo2url"),200)
		End If
	End If


	If vAction = "" Then
		
		If vIdx = "" Then

			vQuery = "INSERT INTO [db_sitemaster].[dbo].[tbl_search_mainmanage]"
			vQuery = vQuery & "(bggubun, bgcolor, bgimg, maskingimg, viewgubun, "
			vQuery = vQuery & "sdate, edate, useyn, textinfouse, "
			vQuery = vQuery & "textinfo1, textinfo1url, textinfo2, textinfo2url, "
			vQuery = vQuery & "memo, reguserid, lastupdateid) "
			vQuery = vQuery & " VALUES "
			vQuery = vQuery & "('" & vBgGubun & "', '" & vBgColor & "', '" & vBgImg & "', '" & vMaskingImg & "', '" & vViewGubun & "', "
			vQuery = vQuery & "'" & vSDate & "', '" & vEDate & "', '" & vUseYN & "', '" & vTextInfoUse & "', "
			vQuery = vQuery & "'" & vTextInfo1 & "', '" & vTextInfo1url & "', '" & vTextInfo2 & "', '" & vTextInfo2url & "', "
			vQuery = vQuery & "'" & vMemo & "', '" & vUserID & "', '" & vUserID & "')"
			dbget.Execute vQuery

		Else

			vQuery = "UPDATE [db_sitemaster].[dbo].[tbl_search_mainmanage] SET "
			vQuery = vQuery & "bggubun = '" & vBgGubun & "' "
			vQuery = vQuery & ", bgcolor = '" & vBgColor & "' "
			vQuery = vQuery & ", bgimg = '" & vBgImg & "' "
			vQuery = vQuery & ", maskingimg = '" & vMaskingImg & "' "
			vQuery = vQuery & ", viewgubun = '" & vViewGubun & "' "
			vQuery = vQuery & ", sdate = '" & vSDate & "' "
			vQuery = vQuery & ", edate = '" & vEDate & "' "
			vQuery = vQuery & ", useyn = '" & vUseYN & "' "
			vQuery = vQuery & ", textinfouse = '" & vTextInfoUse & "' "
			vQuery = vQuery & ", textinfo1 = '" & vTextInfo1 & "' "
			vQuery = vQuery & ", textinfo1url = '" & vTextInfo1url & "' "
			vQuery = vQuery & ", textinfo2 = '" & vTextInfo2 & "' "
			vQuery = vQuery & ", textinfo2url = '" & vTextInfo2url & "' "
			vQuery = vQuery & ", memo = '" & vMemo & "' "
			vQuery = vQuery & ", lastupdateid = '" & vUserID & "' "
			vQuery = vQuery & ", lastupdatedate = getdate() "
			vQuery = vQuery & "where idx = '" & vIdx & "' "
			dbget.Execute vQuery

		End If

		Response.Write "<script>alert('처리되었습니다.');opener.location.reload();window.close();</script>"
		
	End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->