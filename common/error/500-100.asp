<%@ language="VBScript" %>
<%
  Option Explicit


  Const lngMaxFormBytes = 200

  Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
  Dim strMethod, lngPos, datNow, strQueryString, strURL

  If Response.Buffer Then
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/html"
    Response.Expires = 0
  End If

  Set objASPError = Server.GetLastError
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<HTML><HEAD><TITLE>10x10 �������� ������ �߻��߽��ϴ�.</TITLE>
<META HTTP-EQUIV="Content-Type" Content="text/html; charset=euc-kr">
<STYLE type="text/css">
  BODY { font: 9pt/12pt ���� }
  H1 { font: 13pt/15pt ����; font-weight:bold; }
  H2 { font: 9pt/12pt ���� }
  A:link { color: red }
  A:visited { color: maroon }
</STYLE>
</HEAD>
<BODY>
<TABLE width="620" border="0" cellspacing="5">
<TR>
	<td width="117" valign="top"><img src="/fiximage/web2012/common/footer_logo.png" /></td>
	<TD>
		<h1>�� �������� ǥ���� �� �����ϴ�.</h1>
		�����Ϸ��� �������� ������ �߻��Ǿ� ǥ���� �� �����ϴ�.
	</td>
</tr>
<tr>
	<td colspan="2">
		<hr>
		<p>������ ���� ������ �ٹ����� <strong>�ý�����</strong>���� �������ּ���.</p>
		<ul>
			<li>������ �߻��� ���� �������� <strong>URL �ּ�</strong>�� �˷��ּ���.</li>
			<li>�Ʒ� ��� ������ �����ؼ� ������ �ּ���.</li>
			<li>���� ���� �Ǵ� �߻� ��Ȳ�� �˷��ֽø� �ذῡ ���� ������ �˴ϴ�.</li>
		</ul>
		<h2>HTTP 500.100 - ���� ���� ����: ASP �����Դϴ�.<br>IIS(���ͳ� ���� ����)</h2>
		<hr>
		<p><strong>��� ����</strong> (�ý��������� �������ּ���.)</p>
		<ul>
			<li>���� ����:<br> <%
			  Dim bakCodepage
			  on error resume next
				bakCodepage = Session.Codepage
				Session.Codepage = 1252
			  on error goto 0
			  Response.Write Server.HTMLEncode(objASPError.Category)
			  If objASPError.ASPCode > "" Then Response.Write Server.HTMLEncode(", " & objASPError.ASPCode)
			    Response.Write Server.HTMLEncode(" (0x" & Hex(objASPError.Number) & ")" ) & "<br>"
			  If objASPError.ASPDescription > "" Then
				Response.Write Server.HTMLEncode(objASPError.ASPDescription) & "<br>"
			  elseIf (objASPError.Description > "") Then
				Response.Write Server.HTMLEncode(objASPError.Description) & "<br>"
			  end if
			  blnErrorWritten = False
			  ' Only show the Source if it is available and the request is from the same machine as IIS
			  If objASPError.Source > "" Then
			    strServername = LCase(Request.ServerVariables("SERVER_NAME"))
			    strServerIP = Request.ServerVariables("LOCAL_ADDR")
			    strRemoteIP =  Request.ServerVariables("REMOTE_ADDR")
			    If (strServerIP = strRemoteIP) And objASPError.File <> "?" Then
			      Response.Write Server.HTMLEncode(objASPError.File)
			      If objASPError.Line > 0 Then Response.Write ", line " & objASPError.Line
			      If objASPError.Column > 0 Then Response.Write ", column " & objASPError.Column
			      Response.Write "<br>"
			      Response.Write "<font style=""COLOR:000000; FONT: 8pt/11pt courier new""><b>"
			      Response.Write Server.HTMLEncode(objASPError.Source) & "<br>"
			      If objASPError.Column > 0 Then Response.Write String((objASPError.Column - 1), "-") & "^<br>"
			      Response.Write "</b></font>"
			      blnErrorWritten = True
			    End If
			  End If
			  If Not blnErrorWritten And objASPError.File <> "?" Then
			    Response.Write "<b>" & Server.HTMLEncode(  objASPError.File)
			    If objASPError.Line > 0 Then Response.Write Server.HTMLEncode(", line " & objASPError.Line)
			    If objASPError.Column > 0 Then Response.Write ", column " & objASPError.Column
			    Response.Write "</b><br>"
			  End If
			%>
			</li>
			<li>������ ����:<br> <%= Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT")) %> <br><br></li>
			<li>������:<br> <%
			  strMethod = Request.ServerVariables("REQUEST_METHOD")
			  Response.Write strMethod & " "
			  If strMethod = "POST" Then
			    Response.Write Request.TotalBytes & " bytes to "
			  End If
			  Response.Write Request.ServerVariables("SCRIPT_NAME")
			  Response.Write "</li>"
			  If strMethod = "POST" Then
			    Response.Write "<p><li>POST Data:<br>"
			    ' On Error in case Request.BinaryRead was executed in the page that triggered the error.
			    On Error Resume Next
			    If Request.TotalBytes > lngMaxFormBytes Then
			      Response.Write Server.HTMLEncode(Left(Request.Form, lngMaxFormBytes)) & " . . ."
			    Else
			      Response.Write Server.HTMLEncode(Request.Form)
			    End If
			    On Error Goto 0
			    Response.Write "</li>"
			  End If
			%> <br><br></li>
			<li>�ð�:<br> <%
			  datNow = Now()
			  Response.Write Server.HTMLEncode(FormatDateTime(datNow, 1) & ", " & FormatDateTime(datNow, 3))
			  on error resume next
				Session.Codepage = bakCodepage
			  on error goto 0
			%> </li>
		</ul>
	</TD>
</TR>
</TABLE>
</BODY>
</HTML>
