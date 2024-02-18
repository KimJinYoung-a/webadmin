<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/util/base64New.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/lib/classes/seminar/seminarCls.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'###############################################
' PageName : /datagridsample/adminmanage/seminar/datatojson.asp
' Discription : datagridsample - seminarlist-view
' Response : response > 결과
' History : 2019.10.15 
'###############################################

'//헤더 출력
Response.ContentType = "application/json"
response.charset = "utf-8"

Dim sData : sData = Request("json")
Dim oJson
dim vSeminar , i
dim page :  page = 1
'// 전송결과 파징
on Error Resume Next

'// json객체 선언
SET oJson = jsArray()
Dim contents_json , contents_object

IF (Err) THEN
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "처리중 오류가 발생했습니다.1"
ELSE

	function roomColor(v)
        Select Case v
            Case "1" roomColor = "#f0787e"
            Case "2" roomColor = "#f5a841"
            Case "3" roomColor = "#5ac5bc"
            Case "4" roomColor = "#ff8817"
            Case "5" roomColor = "#ac4bff"
			Case "6" roomColor = "#9fc7e4"
        End Select
    end function 

    dim sqlStr
	sqlStr = "SELECT idx, roomname , ROW_NUMBER() OVER(ORDER BY orderNo asc) as rownumber  FROM db_partner.dbo.tbl_seminarRoom WITH(NOLOCK) WHERE isusing='Y' Order by orderNo ASC"
	rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
	
	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF            
			Set oJson(null) = jsObject()
                oJson(null)("text")          = ""& rsget("roomname") &""
                oJson(null)("id")	         = rsget("idx")
                oJson(null)("color")	     = ""& roomColor(rsget("rownumber")) &""
			rsget.MoveNext
		loop
	end if
	rsget.Close	

	'// 결과 출력
	IF (Err) then
		oJson("response") = getErrMsg("9999",sFDesc)
		oJson("faildesc") = "처리중 오류가 발생했습니다.2"
	end if
end if

'Json 출력(JSON)
oJson.flush
Set oJson = Nothing

if ERR then Call OnErrNoti()
On Error Goto 0
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->