<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_apply.asp"-->
<%
	Dim cApply, page, userid, username, i
	page = request("page")
	userid = request("userid")
	username = request("username")
	
	if page = "" then page = 1
		
set cApply = new CCultureApply
	cApply.FPageSize = 20
	cApply.FCurrPage = page
	cApply.Fuserid = userid
	cApply.Fusername = username
	cApply.getApplyList()
%>

<script language="javascript">
function trview(tri)
{
	document.getElementById(""+tri+"").style.display = "block";
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			&nbsp;아이디: <input type="text" name="userid" value="<%= userid%>" size="10">
			&nbsp;&nbsp;&nbsp;이름: <input type="text" name="username" value="<%= username%>" size="10"> 	
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">		
		</td>
	</tr>
	</form>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if cApply.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= cApply.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= cApply.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center">idx</td>
		<td align="center">아이디</td>
		<td align="center">이름</td>	
		<td align="center">전화번호</td>
		<td align="center">이메일</td>
		<td align="center">등록일</td>
    </tr>
	<% for i=0 to cApply.FresultCount-1 %>
    <tr bgcolor="#FFFFFF" style="cursor:pointer;" onClick="trview('tr<%=i%>')">
		<td align="center"><%= cApply.FItemList(i).Fidx %></td>
		<td align="center"><font color="<%= getUserLevelColorByDate(cApply.FItemList(i).Fuserlevel, left(cApply.FItemList(i).Fregdate,10)) %>"><%= cApply.FItemList(i).Fuserid %></font></td>
		<td align="center"><%= cApply.FItemList(i).Fusername %></td>
		<td align="center"><%= cApply.FItemList(i).Fusercell %></td>
		<td align="center"><%= cApply.FItemList(i).Fusermail %></td>
		<td align="center"><%= cApply.FItemList(i).Fregdate %></td>
    </tr>
    <tr id="tr<%=i%>" bgcolor="#FFFFFF" style="display:none;">
    	<td colspan="6">
    		URL : <a href="<%= CHKIIF(Left(cApply.FItemList(i).Flinkurl,7)="http://","","http://") %><%= cApply.FItemList(i).Flinkurl %>" target="_blank"><%= cApply.FItemList(i).Flinkurl %></a><br>
    		신청이유 :<br>
    		<%= Replace(cApply.FItemList(i).Fwhyapply,vbCrLf,"<br>") %>
    	</td>
    </tr>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if cApply.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= cApply.StartScrollPage-1 %>&userid=<%=userid%>&username=<%=username%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + cApply.StartScrollPage to cApply.StartScrollPage + cApply.FScrollCount - 1 %>
				<% if (i > cApply.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(cApply.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&userid=<%=userid%>&username=<%=username%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if cApply.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&userid=<%=userid%>&username=<%=username%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%	set cApply = nothing	%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->