<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'// UTF-8 변환
session.codePage = 65001
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  SCM 우편번호 찾기
' History : 2016.07.01 한용민 프론트 이전 생성
' 호출시 UTF-8이 아니면 에러남
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%

Function SendReqGet(call_url, sedata)
    dim igetURL
    igetURL = call_url
    if sedata<>"" then igetURL = igetURL&"?"&sedata
    
    dim xmlHttp
    Set xmlHttp = CreateObject("Msxml2.ServerXMLHTTP.3.0")
    
    xmlHttp.open "GET",igetURL, False
    xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"  ''UTF-8 charset 필요.
    xmlHttp.setTimeouts 5000,15000,15000,15000 ''2013/03/14 추가
    xmlHttp.Send
    
    SendReqGet = BinaryToText(xmlHttp.responseBody, "UTF-8")
    set xmlHttp=Nothing
end function

'//바이너리 데이터 TEXT형태로 변환
Function  BinaryToText(BinaryData, CharSet)
	 Const adTypeText = 2
	 Const adTypeBinary = 1

	 Dim BinaryStream
	 Set BinaryStream = CreateObject("ADODB.Stream")

	'원본 데이터 타입
	 BinaryStream.Type = adTypeBinary

	 BinaryStream.Open
	 BinaryStream.Write BinaryData
	 ' binary -> text
	 BinaryStream.Position = 0
	 BinaryStream.Type = adTypeText

	' 변환할 데이터 캐릭터셋
	 BinaryStream.CharSet = CharSet

	'변환한 데이터 반환
	 BinaryToText = BinaryStream.ReadText

	 Set BinaryStream = Nothing
End Function


dim currentPage : currentPage = request("currentPage")
dim countPerPage : countPerPage = request("countPerPage")
dim confmKey : confmKey = "U01TX0FVVEgyMDE2MDcwNDIwMjE0NDEzNTk5"
dim keyword : keyword = request("keyword")
dim callback : callback = request("callback")

'// 유효한 검색키워드만 추출 (Check SQL Inject)
keyword = ReplaceRequestSpecialChar(keyword)
keyword = RepWord(keyword,"[^가-힣a-zA-Z0-9.&%\-\_\s]","")

'keyword = "동숭동"

'response.write keyword
'response.end

if currentPage="" then currentPage=1
if countPerPage="" then countPerPage=10


''URLEncoder.encode(keyword,"UTF-8")
'dim addParam : addParam = "currentPage="&currentPage&"&countPerPage="&countPerPage&"&keyword="&server.UrlEncode(keyword)&"&confmKey="&confmKey
dim addParam : addParam = "currentPage="&currentPage&"&countPerPage="&countPerPage&"&keyword="&(keyword)&"&confmKey="&confmKey

dim iURI : iURI="http://www.juso.go.kr/addrlink/addrLinkApiJsonp.do"
dim retText : retText= SendReqGet(iURI,addParam)


        
response.write callback&TRIM(retText)
session.codePage = 949 : Response.End
%>