<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim vIdx, vDisp1, vQuery, vTheme, vSortNo, vIsUsing

If Request("action") = "proc" Then
	Call Proc()
End If

vIdx = Request("idx")
vDisp1 = Request("disp1")
If vIdx <> "" Then
	vQuery = "select * from db_sitemaster.dbo.tbl_category_MDChoice_theme where idx = '" & vIdx & "'"
	rsget.Open vQuery,dbget,1
	vTheme = html2db(rsget("subject"))
	vSortNo = rsget("sortno")
	vIsUsing = rsget("isusing")
Else
	vSortNo = "99"
	vIsUsing = "Y"
End If
%>
<html>
<head></head>
<body LEFTMARGIN="0"  MARGINWIDTH="0" MARGINHEIGHT="0">
<script language='javascript'>
function trim(str) {
	str = str.replace(/^\s+/, '');
	for (var i = str.length - 1; i > 0; i--) {
		if (/\S/.test(str.charAt(i))) {
			str = str.substring(0, i + 1);
			break;
		}
	}
	return str;
}
function chk_input(f){
	if(trim(f.theme.value) == ""){
		alert('�׸��� �Է��ϼ���');
		f.theme.value="";
		f.theme.focus();
		return;
	}
	if(trim(f.sortno.value) == ""){
		alert('���Ĺ�ȣ�� �Է��ϼ���');
		f.sortno.value="";
		f.sortno.focus();
		return;
	}
	f.submit();
}
</script>
<form name="frm1" method="post" action="category_md_choice_newtheme.asp" style="margin:0px;" onsubmit="return false;">
<input type="hidden" name="action" value="proc">
<input type="hidden" name="idx" value="<%=vIdx%>">
<input type="hidden" name="disp1" value="<%=vDisp1%>">
<% If vIdx <> "" Then Response.Write "No. " & vIdx & "&nbsp;" End If %>
�׸� : <input type="text" name="theme" value="<%=vTheme%>" size="40">
���Ĺ�ȣ : <input type="text" name="sortno" value="<%=vSortNo%>" size="10">
������� : <select name="isusing">
<option value="Y" <% If vIsUsing = "Y" Then %>selected<% End If %>>���</option>
<option value="N" <% If vIsUsing = "N" Then %>selected<% End If %>>������</option>
</select>
<input type="button" value="�� ��" onClick="chk_input(this.form);">
&nbsp;&nbsp;&nbsp;
<input type="button" value="�� ��" onClick="location.href='category_md_choice_newtheme.asp';">
</form>
<script>
frm1.theme.focus();
</script>
</body>
</html>
<%
Function Proc()
	Dim vIdx, vDisp1, vQuery, vTheme, vSortNo, vIsUsing
	vIdx = Request("idx")
	vDisp1 = Request("disp1")
	vTheme = html2db(Request("theme"))
	vTheme = Replace(vTheme,"'","")
	vTheme = Replace(vTheme,chr(34),"")
	vIsUsing = Request("isusing")
	vSortNo = Request("sortno")
	
	If vIdx <> "" Then
		vQuery = "UPDATE db_sitemaster.dbo.tbl_category_MDChoice_theme SET subject = '" & vTheme & "', sortno = '" & vSortNo & "', isusing = '" & vIsUsing & "' WHERE idx = '" & vIdx & "'"
		dbget.Execute vQuery
	Else
		vQuery = "INSERT INTO db_sitemaster.dbo.tbl_category_MDChoice_theme(disp1, subject, sortno, isusing) VALUES('" & vDisp1 & "','" & vTheme & "','" & vSortNo & "','" & vIsUsing & "')"
		dbget.Execute vQuery
	End If
	
	Response.Write "<Script>alert('����Ǿ����ϴ�.');parent.location.reload();</script>"
	dbget.close()
	Response.End
End Function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->