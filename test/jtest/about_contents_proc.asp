<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Response.CharSet="UTF-8"
'###########################################################
' Description : 히치하이커 컨텐츠
' Hieditor : 2014.07.17 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/about_hitchhiker_contentsCls.asp"-->
<%
dim contentslinkarr, deviceidxarr
dim deviceidx, contentsidx, mode, sortnum, isusing, contentslink
dim gubun, con_viewthumbimg, con_title, con_sdate, con_edate, con_movieurl, con_regdate, con_detail

Dim sData, rst, objJson, iBody, istrParam, lst

	deviceidx = Request("deviceidx")
	contentslink = request("contentslink")
	mode = requestCheckvar(Request("mode"),16)
	isusing = requestCheckvar(Request("isusing"),1)
	gubun = requestCheckvar(Request("hicprogbn"),1)
	sortnum = requestCheckvar(Request("sortnum"),10)
	contentsidx = requestCheckvar(Request("contentsidx"),10)

	con_title = requestCheckvar(Request("con_title"),60)
	con_sdate = requestCheckvar(Request("con_sdate"),10)
	con_edate = requestCheckvar(Request("con_edate"),10)
	con_detail = requestCheckvar(Request("con_detail"),150)
	con_regdate = requestCheckvar(Request("con_regdate"),10)
	con_movieurl = requestCheckvar(Request("con_movieurl"),500)
	con_viewthumbimg = requestCheckvar(Request("con_viewthumbimg"),150)

	dim sqlstr, getdate, i
	if mode = "EDIT" then
		istrParam = ""
		istrParam = istrParam & "{"
		istrParam = istrParam & "  ""contentIdx"": "& contentsidx &","
		istrParam = istrParam & "  ""conTitle"": """& html2db(con_title) &""","
		istrParam = istrParam & "  ""conSDate"": """& con_sdate &""","
		istrParam = istrParam & "  ""conEDate"": """& con_edate &""","
		istrParam = istrParam & "  ""isUsing"": """& isusing &""","
		istrParam = istrParam & "  ""conViewThumbImg"": """& con_viewthumbimg &""","
		istrParam = istrParam & "  ""conMovieUrl"": """& html2db(con_movieurl) &""","
		istrParam = istrParam & "  ""conDetail"": """& html2db(con_detail) &""""
		istrParam = istrParam & "}"
		SET objJson = CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objJson.OPEN "PUT", "http://localhost:58658/api/Hitchhiker/Update", true
			objJSON.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
			objJson.Send(istrParam)

			If objJson.ReadyState <> 4 Then
				objJson.waitForResponse 150
			End If

			If objJson.Status <> "200" Then
				response.write "저장실패"
				response.end
			End If
		SET objJson = nothing

	elseif mode = "NEW" then
		istrParam = ""
		istrParam = istrParam & "{"
		istrParam = istrParam & "  ""gubun"": "& gubun &","
		istrParam = istrParam & "  ""conTitle"": """& html2db(con_title) &""","
		istrParam = istrParam & "  ""conSDate"": """& con_sdate &""","
		istrParam = istrParam & "  ""conEDate"": """& con_edate &""","
		istrParam = istrParam & "  ""isUsing"": """& isusing &""","
		istrParam = istrParam & "  ""conViewThumbImg"": """& con_viewthumbimg &""","
		istrParam = istrParam & "  ""conMovieUrl"": """& html2db(con_movieurl) &""","
		istrParam = istrParam & "  ""conDetail"": """& html2db(con_detail) &""""
		istrParam = istrParam & "}"
		SET objJson = CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objJson.OPEN "POST", "http://localhost:58658/api/Hitchhiker/Create", true
			objJSON.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
			objJson.Send(istrParam)

			If objJson.ReadyState <> 4 Then
				objJson.waitForResponse 150
			End If

			If objJson.Status <> "200" Then
				response.write "저장실패"
				response.end
			End If
		SET objJson = nothing
	end if
%>

<script language = "javascript">
	alert("저장되었습니다."); //저장되었습니다 라는 메시지띄움
	opener.location.reload(); //이창을 띄운 부모창을 리로드함
	self.close();			  //이창을 닫음
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->