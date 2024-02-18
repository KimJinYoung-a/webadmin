<%
'####### 업로드한 파일을 오프너에서 바로 innerHTML 해주고 내용부분의 에디터가 document.domain 문제때문에 작성한 내용을 인식못해서 어쩔수 없이 우회를 함.

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
	var inHtml = "<input type='hidden' name='info_file' value='<%=vTemp_URL(i)%>'><input type='hidden' name='info_realfile' value='<%=vTemp_RealURL(i)%>'><img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'> <a href=javascript:alert('먼저저장해야합니다.');><%=vTemp_Name(i)%></a>";
	c0.innerHTML = inHtml;
</script>
<%
Next
%>

<script language="javascript">
	alert("첨부파일이 등록되었습니다.\n\n파일 등록후 확인버튼을 눌러야 저장됩니다.");
	window.close();
</script>