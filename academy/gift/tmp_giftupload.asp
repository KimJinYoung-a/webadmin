<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  사은품종류 파일 등록 '도메인 문제로 우회 시킴
' History : 2010.09.27 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim strImgUrl , sName ,sSpan
	strImgUrl = request("strImgUrl")
	sName = RequestCheckvar(request("sName"),32)
	sSpan = request("sSpan")
  	if strImgUrl <> "" then
		if checkNotValidHTML(strImgUrl) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end If
  	if sSpan <> "" then
		if checkNotValidHTML(sSpan) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end if
%>

<script language="javascript">
	document.domain = "10x10.co.kr";	
	alert("이미지가 등록되었습니다.\n\n이미지 등록후 저장버튼을 눌러야 처리완료됩니다.");
	opener.fnAddImage2('<%=strImgUrl%>','<%=sName%>','<%=sSpan%>');
	self.close();
</script>

