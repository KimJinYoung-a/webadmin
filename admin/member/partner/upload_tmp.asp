<%
'####### ���ε��� ������ �����ʿ��� �ٷ� innerHTML ���ְ� ����κ��� �����Ͱ� document.domain ���������� �ۼ��� ������ �νĸ��ؼ� ��¿�� ���� ��ȸ�� ��.

Dim vTemp_URL, vTemp_RealURL, vTemp_Name, i
vTemp_URL 		= Split(Request("fileurl_tmp"),",")
vTemp_RealURL	= Split(Request("realfileurl_tmp"),",")
vTemp_Name 		= Split(Request("filename_tmp"),",")

For i = 0 To UBOUND(vTemp_URL)
%>
<script language="javascript">
	var f = opener.eval("document.all.fileup");

	var rowLen = f.rows.length;
	var r  = f.insertRow(rowLen++);
	var c0 = r.insertCell(0);
	
	var Html;

	c0.innerHTML = "&nbsp;";
	var inHtml = "<input type='hidden' name='info_file' value='<%=vTemp_URL(i)%>'><input type='hidden' name='info_realfile' value='<%=vTemp_RealURL(i)%>'><img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'> <a href=javascript:alert('���������ؾ��մϴ�.');><%=vTemp_Name(i)%></a>";
	c0.innerHTML = inHtml;
</script>
<%
Next
%>

<script language="javascript">
	alert("÷�������� ��ϵǾ����ϴ�.\n\n���� ����� Ȯ�ι�ư�� ������ ����˴ϴ�.");
	window.close();
</script>