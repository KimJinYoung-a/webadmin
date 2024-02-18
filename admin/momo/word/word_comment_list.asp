<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 감성사전 코맨트 리스트
' Hieditor : 2009.10.28 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<% 
dim keyid , page, vIdx, vGubun, vParam
	keyid = request("keyid")
	page = request("page")
	vGubun = request("gubun")
	if page = "" then page = 1
		
	vParam = "keyid="&keyid&"&page="&page&"&gubun="&vGubun&""

dim ocomment , i
set ocomment = new cword_list
	ocomment.FPageSize = 20
	ocomment.FCurrPage = page
	ocomment.frectkeyid = keyid
	ocomment.frectgubun = vGubun
	ocomment.fwordcomment_list()
%>

<script language="javascript">

	function comment_delete(idx){
		var comment_delete = window.open('/admin/momo/word/word_comment_process.asp?idx='+idx,'comment_delete','width=800,height=600,scrollbars=yes,resizable=yes');
		comment_delete.focus();
	}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<a href="/admin/momo/word/word_comment_list.asp?keyid=<%=keyid%>&page=1&gubun="><% IF vGubun = "" Then %><b><% End If %>[전체보기]<% IF vGubun = "" Then %></b><% End If %></a>&nbsp;&nbsp;&nbsp;
		<a href="/admin/momo/word/word_comment_list.asp?keyid=<%=keyid%>&page=1&gubun=p"><% IF vGubun = "p" Then %><b><% End If %>[포토만보기]<% IF vGubun = "p" Then %></b><% End If %></a>&nbsp;&nbsp;&nbsp;
		<a href="/admin/momo/word/word_comment_list.asp?keyid=<%=keyid%>&page=1&gubun=n"><% IF vGubun = "n" Then %><b><% End If %>[일반글만보기]<% IF vGubun = "n" Then %></b><% End If %></a>
	</td>
	<td align="right"><input type="button" value="Best적용" onClick="frm.submit();"></td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="word_comment_list_process.asp" method="post">
<% if ocomment.fresultcount > 0 then %>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center">
			idx
		</td>	
		<td align="center">
			감성사전id
		</td>
		<td align="center">
			이미지
		</td>			
		<td align="center">
			고객ID
		</td>	
		<td align="center">
			코맨트
		</td>
		<td align="center">
			베스트
		</td>
		<td align="center">
			등록일
		</td>	
		<td align="center">
			사용여부
		</td>
			
		<td align="center">
			비고
		</td>	
	</tr>
	<% for i = 0 to ocomment.fresultcount -1
			vIdx = vIdx & ocomment.fitemlist(i).fidx & ","
	%>
	<tr bgcolor="FFFFFF">
		<td align="center">
			<%= ocomment.fitemlist(i).fidx %>
		</td>	
		<td align="center">
			<%= ocomment.fitemlist(i).fkeyid %>
		</td>
		<td align="center">
			<img src="<%=webImgUrl%>/momo/photo/user/<%= ocomment.FItemList(i).fmainimage %>" width=50 height=50>
		</td>			
		<td align="center">
			<%= ocomment.fitemlist(i).fuserid %>
		</td>	
		<td align="center">
			<%= nl2br(ocomment.fitemlist(i).fcomment) %>
		</td>	
		<td align="center">
			<input type="checkbox" name="isbest" value="<%= ocomment.fitemlist(i).fidx %>" <% If ocomment.fitemlist(i).fisbest = "o" Then Response.Write " checked" End If %>>
		</td>
		<td align="center">
			<%= left(ocomment.fitemlist(i).fregdate,10) %>
		</td>	
		<td align="center">
			<%= ocomment.fitemlist(i).fisusing %>
		</td>			
		<td align="center">
			<input type="button" class="button" value="노출안함" onclick="comment_delete(<%= ocomment.fitemlist(i).fidx %>);">
		</td>	
	</tr>	
	<% next %>
<% else %>
<tr bgcolor="FFFFFF">
	<td align="center">검색 결과가 없습니다.
	</td>	
</tr>
<% end if %>
<input type="hidden" name="totidx" value="<%=vIdx%>">
<input type="hidden" name="keyid" value="<%=keyid%>">
<input type="hidden" name="nowpage" value="<%=page%>">
<input type="hidden" name="gubun" value="<%=vGubun%>">
</form>
    <tr height="25" bgcolor="FFFFFF">
    	<td></td>
		<td colspan="7" align="center">
	       	<% if ocomment.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= ocomment.StartScrollPage-1 %>&keyid=<%=keyid%>&gubun=<%=vGubun%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + ocomment.StartScrollPage to ocomment.StartScrollPage + ocomment.FScrollCount - 1 %>
				<% if (i > ocomment.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ocomment.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&keyid=<%=keyid%>&gubun=<%=vGubun%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if ocomment.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&keyid=<%=keyid%>&gubun=<%=vGubun%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
		<td align="center"><input type="button" value="Best적용" onClick="frm.submit();"></td>
	</tr>
</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
set	ocomment = nothing
%>