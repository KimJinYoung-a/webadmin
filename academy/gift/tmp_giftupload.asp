<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����ǰ���� ���� ��� '������ ������ ��ȸ ��Ŵ
' History : 2010.09.27 �ѿ�� ����
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
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end If
  	if sSpan <> "" then
		if checkNotValidHTML(sSpan) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end if
%>

<script language="javascript">
	document.domain = "10x10.co.kr";	
	alert("�̹����� ��ϵǾ����ϴ�.\n\n�̹��� ����� �����ư�� ������ ó���Ϸ�˴ϴ�.");
	opener.fnAddImage2('<%=strImgUrl%>','<%=sName%>','<%=sSpan%>');
	self.close();
</script>

