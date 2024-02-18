
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#DDDDFF" height="35">
    <td>
    	<% if request.ServerVariables("SCRIPT_NAME")="/admin/hitchhiker/mainbanner/index.asp" then %>
    		<a href="/admin/hitchhiker/mainbanner/index.asp?menupos=<%=menupos%>"><b>메인배너관리</b></a>
    	<% else %>
    		<a href="/admin/hitchhiker/mainbanner/index.asp?menupos=<%=menupos%>">메인배너관리</a>
    	<% end if %>
    </td>
    <td>
		<% if request.ServerVariables("SCRIPT_NAME")="/admin/hitchhiker/issuearea/about_list.asp" then %>
    		<a href="/admin/hitchhiker/issuearea/about_list.asp?menupos=<%=menupos%>"><b>이슈영역</b></a>
    	<% else %>
			<a href="/admin/hitchhiker/issuearea/about_list.asp?menupos=<%=menupos%>">이슈영역</a>
		<% end if %>
    </td>
    <td>
		<% if request.ServerVariables("SCRIPT_NAME")="/admin/hitchhiker/preview/index.asp" then %>
			<a href="/admin/hitchhiker/preview/index.asp?menupos=<%=menupos%>"><b>프리뷰</b></a>
    	<% else %>
    		<a href="/admin/hitchhiker/preview/index.asp?menupos=<%=menupos%>">프리뷰</a>
    	<% end if %>
    </td>
</tr>
</table>
<br>

