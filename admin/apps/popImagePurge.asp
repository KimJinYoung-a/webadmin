<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �̹��� ĳ�� - Purging
' History : 2014.12.30 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
function fn_AddIISLOG(iAddLogs)
    ''addLog �߰� �α� //2017/07/19
    if (request.ServerVariables("QUERY_STRING")<>"") then iAddLogs="&"&iAddLogs
    response.AppendToLog iAddLogs
end function

	dim vImgUrl, vRstCd, vrtMsg, vBUF
	dim iURL, postdata
	vImgUrl = requestCheckvar(request("imgurl"),256)

	if vImgUrl<>"" then
		'XML ��ü ����
		dim objXML, xmlDOM
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		
		objXML.Open "GET", "https://api.xtrmcdn.co.kr:28091/api/v1/purge/TID_16641/?target=" & (vImgUrl), false
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setRequestHeader "X-ITX-Security-Secret", "88719016447d0a173b8e422f443509fc9c30bbb91d1ce22a4d6f278edb4d12b8"
		objXML.setRequestHeader "X-ITX-Security-Principal", "10x10"
		
		objXML.Send
		If objXML.Status = "200" Then
			vBUF =BinaryToText(objXML.ResponseBody, "euc-kr")

			'// ���۰�� ��¡
			dim oResult, strResult
			on Error Resume Next
			set oResult = JSON.parse(vBUF)
				set strResult = oResult.meta
			set oResult = Nothing
			On Error Goto 0

			vRstCd = strResult.statusCode
			''vrtMsg = strResult.message
		end if
		'XML ��ü ����
		Set objXML = Nothing
		
		call fn_AddIISLOG(vImgUrl)
	end if

	'//���̳ʸ� ������ TEXT���·� ��ȯ
	Function  BinaryToText(BinaryData, CharSet)
		 Const adTypeText = 2
		 Const adTypeBinary = 1
	
		 Dim BinaryStream
		 Set BinaryStream = CreateObject("ADODB.Stream")
	
		'���� ������ Ÿ��
		 BinaryStream.Type = adTypeBinary
	
		 BinaryStream.Open
		 BinaryStream.Write BinaryData
		 ' binary -> text
		 BinaryStream.Position = 0
		 BinaryStream.Type = adTypeText
	
		' ��ȯ�� ������ ĳ���ͼ�
		 BinaryStream.CharSet = CharSet 
	
		'��ȯ�� ������ ��ȯ
		 BinaryToText = BinaryStream.ReadText
	
		 Set BinaryStream = Nothing
	End Function 
%>
<script type="text/javascript">
function checkform(frm) {
	if(frm.imgurl.value.length<10) {
		alert("�̹��� URL�� Full Path�� �Է����ּ���.");
		frm.imgurl.focus();
		return false;
	}

	if(frm.imgurl.value.indexOf('http')<0) {
		alert("HTTP�� ������ ��ü URL�� �Է����ּ���.");
		frm.imgurl.focus();
		return false;
	}

	return true;
}
</script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<h3>�ٹ����� �̹��� ĳ�� Update</h3>
<!-- ǥ ��ܹ� ����-->
<form name="frm" method="POST" action="" onSubmit="return checkform(this);">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#FFFFFF" >
    <td width="80" bgcolor="#EEEEEE">�̹��� URL</td>
    <td align="left">
		<input type="text" name="imgurl" value="<%=vImgUrl%>" maxlength="256" style="width:100%;" class="input_text" />
		<p style="color:#888; padding-top:3px;">�� ������ ���Ͻô� �̹��� URL�� �־��ּ���.</p>
	</td>
	<td width="60" bgcolor="#EEEEEE"><input type="image" src="/images/icon_confirm.gif" width="45" height="20" border="0" /></td>
</tr>
</table>
</form>
<!-- ǥ ��ܹ� ��-->

<% if vRstCd="200" then %>
<!-- ��� ǥ�� -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td>
		<ul>
			<li>�̹��� ������Ʈ ��û�� ���������� �Ϸ�Ǿ����ϴ�.</li>
			<li>��� ������ ����� ������ �ð��� ���� �ɸ� �� ������ ��ٷ��ּ���.</li>
		</ul>
		<img src="<%=vImgUrl%>" width="100%" />
	</td>
</tr>
</table>
<% else %>
	<% if vRstCd<>"" then %><script type="text/javascript">alert('URL�� �߸��Ǿ��ų�, ĳ�õ� �̹����� �����ϴ�.');</script><% end if %>
	<ul>
		<li>�̹��� ĳ�� ���񽺿� ĳ�õǾ��ִ� �̹����� ���� ������ �� �ֽ��ϴ�.</li>
		<li>��û�� �Ϸ�Ǿ� <span style="color:#FF6633;">���� ���񽺿� ����</span>�� ������ �ð��� ���� �ɸ� �� ������ ��ٷ��ּ���.</li>
		<li>���� �̹����� <b style="color:#FF6633;">������ ���û</b>���� ������.</li>
	</ul>
<% end if %>
<!-- #include virtual="/admin/lib/poptail.asp"-->