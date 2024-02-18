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
' PageName : /apps/protoV1/dashboardProc.asp
' Discription : 대시보드 정보
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
Dim sUserID
Dim sData : sData = Request("json")
Dim oJson, sOS, sVerCheck

'// 전송결과 파징
On Error Resume Next

Dim oResult
Set oResult = JSON.parse(sData)
	sType = oResult.type
	sUserID = request.Cookies("partner")("userid")
	sOS = requestCheckVar(oResult.OS,10)
	sVerCheck = requestCheckVar(oResult.version,6)  ''' 2016/12/19 추가 api 버전 확인
Set oResult = Nothing

'// json객체 선언
Set oJson = jsObject()

Dim URIFIX 
URIFIX = "https://webadmin.10x10.co.kr"
If application("Svr_Info")="Dev" Then
	 URIFIX = "http://testwebadmin.10x10.co.kr"   
End If

If (Err) Then
    oJson("response") = getErrMsg("9999",sFDesc)
    oJson("faildesc") = "처리중 오류가 발생했습니다."
ElseIf (LCASE(sType)<>"dashboard") Then
    oJson("response") = getErrMsg("9999",sFDesc)
    oJson("faildesc") = "잘못된 접근입니다."
ElseIf (sUserID="") Then
    oJson("response") = getErrMsg("9999",sFDesc)
    oJson("faildesc") = "로그인 하셔야합니다."
Else
	Dim sqlStr
	
	If (sVerCheck=2) Then
		'//뱃지 카운트 확인
		oJson("response") = getErrMsg("1000",sFDesc)
		sqlStr = "select * from [db_academy].[dbo].[tbl_academy_app_iconbadge_count] where makerid='" + Cstr(sUserID) + "'"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		if not rsACADEMYget.EOF Then
			oJson("ordercount")=rsACADEMYget("mibaljucnt")
			oJson("delivercount")=rsACADEMYget("ordercnt")
			oJson("cscount")=rsACADEMYget("cscnt")
			oJson("qnacount")=rsACADEMYget("qnacnt")
		Else
			oJson("ordercount")=0
			oJson("delivercount")=0
			oJson("cscount")=0
			oJson("qnacount")=0
		end if
		rsACADEMYget.Close
		Set oJson("submenu") = getFingerBoardMenuJSon
	Else
		sqlStr = "select count(itemid) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item" + vbcrlf
		sqlStr = sqlStr & " where itemid<>0" + vbcrlf
		sqlStr = sqlStr & " and sellyn='Y'" + vbcrlf
		sqlStr = sqlStr & " and isusing='Y'" + vbcrlf
		sqlStr = sqlStr & " and makerid='" + sUserID + "'"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		oJson("response") = getErrMsg("1000",sFDesc)
		oJson("sellproduct") = rsACADEMYget("cnt")
		rsACADEMYget.close

		sqlStr = "select count(itemid) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_wait_item" + vbcrlf
		sqlStr = sqlStr & " where itemid<>0" + vbcrlf
		sqlStr = sqlStr & " and currstate='1'" + vbcrlf
		sqlStr = sqlStr & " and makerid='" + sUserID + "'"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		oJson("waitproduct") = rsACADEMYget("cnt")
		rsACADEMYget.close
	End If

	Dim olect, ix, objRst
	Set olect = New ClecturerList
	olect.FPageSize = 5
	olect.FCurrPage = 1
	olect.FrectDoc_Type = "G010"
	olect.FRectMakerID = request.cookies("partner")("userid")
	olect.fnGetlecturerList()

	Set oJson("topnotice") = jsArray()

	For ix=0 To olect.fresultcount-1
		 Set oJson("topnotice")(Null) = jsObject()
		 oJson("topnotice")(null)("subject") = CStr(stripHTML(olect.FItemList(ix).fdoc_subject))
		 oJson("topnotice")(null)("date") = CStr(FormatDate(olect.FItemList(ix).fdoc_regdate,"0000.00.00"))
    	 oJson("topnotice")(null)("url") = CStr(URIFIX & "/apps/academy/notice/noticeView.asp?idx=" & CStr(olect.FItemList(ix).fdoc_idx))
	Next

	Dim ofreeboard
	Set ofreeboard = New ClecturerList
	ofreeboard.FPageSize = 3
	ofreeboard.FCurrPage = 1
	ofreeboard.FrectDoc_Type = "G020"
	ofreeboard.FRectMakerID = request.cookies("partner")("userid")
	ofreeboard.fnGetlecturerList()

	Set oJson("topfreeboard") = jsArray()

	For ix=0 To ofreeboard.fresultcount-1
		 Set oJson("topfreeboard")(Null) = jsObject()
		 oJson("topfreeboard")(null)("subject") = "[" & ofreeboard.FItemList(ix).GetTypeName & "] " & CStr(stripHTML(ofreeboard.FItemList(ix).fdoc_subject))
		 oJson("topfreeboard")(null)("date") = CStr(FormatDate(ofreeboard.FItemList(ix).fdoc_regdate,"0000.00.00"))
    	 oJson("topfreeboard")(null)("url") = CStr(URIFIX & "/apps/academy/notice/noticeView.asp?idx=" & CStr(ofreeboard.FItemList(ix).fdoc_idx))
		 If ofreeboard.FItemList(ix).fdoc_ans_ox="o" Then
		 oJson("topfreeboard")(null)("replyn") = "Y"
		 Else
		 oJson("topfreeboard")(null)("replyn") = "N"
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