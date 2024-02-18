<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이미지 캐시 - Purging
' History : 2014.12.30 허진원 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
function fn_AddIISLOG(iAddLogs)
    ''addLog 추가 로그 //2017/07/19
    if (request.ServerVariables("QUERY_STRING")<>"") then iAddLogs="&"&iAddLogs
    response.AppendToLog iAddLogs
end function

	dim vImgUrl, vRstCd, vrtMsg, vBUF
	dim iURL, postdata
	vImgUrl = requestCheckvar(request("imgurl"),256)

	if vImgUrl<>"" then
		'XML 객체 생성
		dim objXML, xmlDOM
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		
		objXML.Open "GET", "https://api.xtrmcdn.co.kr:28091/api/v1/purge/TID_16641/?target=" & (vImgUrl), false
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setRequestHeader "X-ITX-Security-Secret", "88719016447d0a173b8e422f443509fc9c30bbb91d1ce22a4d6f278edb4d12b8"
		objXML.setRequestHeader "X-ITX-Security-Principal", "10x10"
		
		objXML.Send
		If objXML.Status = "200" Then
			vBUF =BinaryToText(objXML.ResponseBody, "euc-kr")

			'// 전송결과 파징
			dim oResult, strResult
			on Error Resume Next
			set oResult = JSON.parse(vBUF)
				set strResult = oResult.meta
			set oResult = Nothing
			On Error Goto 0

			vRstCd = strResult.statusCode
			''vrtMsg = strResult.message
		end if
		'XML 객체 해제
		Set objXML = Nothing
		
		call fn_AddIISLOG(vImgUrl)
	end if

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
%>
<script type="text/javascript">
function checkform(frm) {
	if(frm.imgurl.value.length<10) {
		alert("이미지 URL을 Full Path로 입력해주세요.");
		frm.imgurl.focus();
		return false;
	}

	if(frm.imgurl.value.indexOf('http')<0) {
		alert("HTTP를 포함한 전체 URL로 입력해주세요.");
		frm.imgurl.focus();
		return false;
	}

	return true;
}
</script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<h3>텐바이텐 이미지 캐시 Update</h3>
<!-- 표 상단바 시작-->
<form name="frm" method="POST" action="" onSubmit="return checkform(this);">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#FFFFFF" >
    <td width="80" bgcolor="#EEEEEE">이미지 URL</td>
    <td align="left">
		<input type="text" name="imgurl" value="<%=vImgUrl%>" maxlength="256" style="width:100%;" class="input_text" />
		<p style="color:#888; padding-top:3px;">※ 갱신을 원하시는 이미지 URL을 넣어주세요.</p>
	</td>
	<td width="60" bgcolor="#EEEEEE"><input type="image" src="/images/icon_confirm.gif" width="45" height="20" border="0" /></td>
</tr>
</table>
</form>
<!-- 표 상단바 끝-->

<% if vRstCd="200" then %>
<!-- 결과 표시 -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td>
		<ul>
			<li>이미지 업데이트 요청이 정상적으로 완료되었습니다.</li>
			<li>모든 서버에 적용될 때까지 시간이 조금 걸릴 수 있으니 기다려주세요.</li>
		</ul>
		<img src="<%=vImgUrl%>" width="100%" />
	</td>
</tr>
</table>
<% else %>
	<% if vRstCd<>"" then %><script type="text/javascript">alert('URL이 잘못되었거나, 캐시된 이미지가 없습니다.');</script><% end if %>
	<ul>
		<li>이미지 캐시 서비스에 캐시되어있는 이미지를 새로 갱신할 수 있습니다.</li>
		<li>요청이 완료되어 <span style="color:#FF6633;">실제 서비스에 적용</span>될 때까지 시간이 조금 걸릴 수 있으니 기다려주세요.</li>
		<li>같은 이미지를 <b style="color:#FF6633;">여러번 재요청</b>하지 마세요.</li>
	</ul>
<% end if %>
<!-- #include virtual="/admin/lib/poptail.asp"-->