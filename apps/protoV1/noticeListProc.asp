<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/util/base64New.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/apps/common/appFunction.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/apps/protoV1/protoFunction.asp"-->
<!-- #include virtual="/apps/academy/notice/lecturerNoticecls.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'###############################################
' PageName : /apps/protoV1/noticeListProc.asp
' Discription : 공지사항 리스트
' Request : json > type, pushid, OS, versioncode, versionname, verserion
' Response : response > 결과
' History : 2016.10.18 허진원 : 신규 생성
'###############################################

'//헤더 출력
Response.ContentType = "text/html"

'---------------------------
'@@ 서버 점검시 아래 주석을 풀어주세요
Set oJson = jsObject()
oJson("response") = getErrMsg("9999",sFDesc)
oJson("faildesc") = "핑거스 서비스가 텐바이텐으로 이전되고 종료 되었습니다. 텐바이텐으로 문의 바랍니다. 감사합니다."
oJson.flush
Set oJson = Nothing
Response.End
'---------------------------

Dim sFDesc
Dim sType
Dim sUserID, sMovePage, PageSize
Dim sData : sData = Request("json")
Dim oJson

'// 전송결과 파징
On Error Resume Next

Dim oResult
Set oResult = JSON.parse(sData)
	sType = oResult.type
	sMovePage = oResult.movepage
	sUserID = request.Cookies("partner")("userid")
Set oResult = Nothing

'// json객체 선언
Set oJson = jsObject()

Dim URIFIX 
URIFIX = "https://webadmin.10x10.co.kr"
If application("Svr_Info")="Dev" Then
	 URIFIX = "http://testwebadmin.10x10.co.kr"   
End If

If sType="notilist" Then
	sType="G010"
	PageSize=12
Else
	sType="G020"
	PageSize=10
End If

If (Err) Then
    oJson("response") = getErrMsg("9999",sFDesc)
    oJson("faildesc") = "처리중 오류가 발생했습니다."
ElseIf (LCASE(sType)="") Then
    oJson("response") = getErrMsg("9999",sFDesc)
    oJson("faildesc") = "잘못된 접근입니다."
ElseIf (sUserID="") Then
    oJson("response") = getErrMsg("9999",sFDesc)
    oJson("faildesc") = "로그인 하셔야합니다."
Else
	If sMovePage = "" Then sMovePage=1
	Dim olect, ix
	Set olect = New ClecturerList
	olect.FPageSize = PageSize
	olect.FCurrPage = sMovePage
	olect.FrectDoc_Type = sType
	olect.FRectMakerID = sUserID
	olect.fnGetlecturerList()

	oJson("response") = getErrMsg("1000",sFDesc)
	oJson("totalcount") = olect.FTotalCount
	Set oJson("notilist") = jsArray()

	For ix=0 To olect.fresultcount-1
		 Set oJson("notilist")(Null) = jsObject()
		 oJson("notilist")(null)("subject") = CStr(stripHTML(olect.FItemList(ix).fdoc_subject))
		 oJson("notilist")(null)("date") = CStr(FormatDate(olect.FItemList(ix).fdoc_regdate,"0000.00.00"))
    	 oJson("notilist")(null)("url") = CStr(URIFIX & "/apps/academy/notice/noticeView.asp?idx=" & CStr(olect.FItemList(ix).fdoc_idx))
		 If olect.FItemList(ix).fdoc_ans_ox="o" Then
		 oJson("notilist")(null)("replyn") = "Y"
		 Else
		 oJson("notilist")(null)("replyn") = "N"
		 End If
	Next
End If

If Err Then Call OnErrNoti()
On Error Goto 0
'Json 출력(JSON)
oJson.flush
Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->