
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#DDDDFF" height="35">
    <td>
    	<% if request.ServerVariables("SCRIPT_NAME")="/admin/hitchhiker/mainbanner/index.asp" then %>
    		<a href="/admin/hitchhiker/mainbanner/index.asp?menupos=<%=menupos%>"><b>���ι�ʰ���</b></a>
    	<% else %>
    		<a href="/admin/hitchhiker/mainbanner/index.asp?menupos=<%=menupos%>">���ι�ʰ���</a>
    	<% end if %>
    </td>
    <td>
		<% if request.ServerVariables("SCRIPT_NAME")="/admin/hitchhiker/issuearea/about_list.asp" then %>
    		<a href="/admin/hitchhiker/issuearea/about_list.asp?menupos=<%=menupos%>"><b>�̽�����</b></a>
    	<% else %>
			<a href="/admin/hitchhiker/issuearea/about_list.asp?menupos=<%=menupos%>">�̽�����</a>
		<% end if %>
    </td>
    <td>
		<% if request.ServerVariables("SCRIPT_NAME")="/admin/hitchhiker/preview/index.asp" then %>
			<a href="/admin/hitchhiker/preview/index.asp?menupos=<%=menupos%>"><b>������</b></a>
    	<% else %>
    		<a href="/admin/hitchhiker/preview/index.asp?menupos=<%=menupos%>">������</a>
    	<% end if %>
    </td>
</tr>
</table>
<br>

