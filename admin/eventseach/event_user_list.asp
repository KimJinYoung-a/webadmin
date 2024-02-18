<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  이벤트 응모자 리스트
' History : 2007.09.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventuserclass.asp"-->

<% 
dim seachbox,eventbox ,page , i
	seachbox = request("seachbox")
	eventbox = request("eventbox")
	page = Request("page")

If page="" Then page = 1

dim oeventuserlist
	set oeventuserlist = new Ceventuserlist
	
	if eventbox = "3" then
		oeventuserlist.FPagesize = 5000
		oeventuserlist.FCurrPage = page
		oeventuserlist.frectgubun="ONEVT"
		oeventuserlist.frectseachbox = seachbox
		oeventuserlist.Feventuserlist3()
	end if 
	if eventbox = "5" then
		oeventuserlist.FPagesize = 5000
		oeventuserlist.FCurrPage = page
		oeventuserlist.frectgubun="ONEVT"
		oeventuserlist.frectseachbox = seachbox
		oeventuserlist.Feventuserlist5()
	end if 
	if eventbox = "7" then
		oeventuserlist.FPagesize = 5000
		oeventuserlist.FCurrPage = page
		oeventuserlist.frectgubun="ONEVT"
		oeventuserlist.frectseachbox = seachbox
		oeventuserlist.Feventuserlist7()
	end if 
	if eventbox = "9" then
		oeventuserlist.FPagesize = 50
		oeventuserlist.FCurrPage = page
		oeventuserlist.frectgubun="ONEVT"		
		oeventuserlist.frectseachbox = seachbox
		oeventuserlist.Feventuserlist99()
	end if 
%>
<script type="text/javascript">

function excel(seachbox,eventbox){
	var popup = window.open('/admin/eventseach/event_user_list_excel.asp?seachbox='+seachbox+'&eventbox='+eventbox,'excel','width=1024,height=768,scrollbars=yes,resizable=yes');
	popup.focus();
}

function goPage(pg){
	document.frm.page.value=pg;
	document.frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		&nbsp;EventType: <% DraweventGubun "eventbox", eventbox %>
	</td>
	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		&nbsp;EventCode: &nbsp; <input type="text" name="seachbox" value="<%= seachbox %>" size="10">
	</td>
</tr>
</form>	
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" name="excelbox" value="엑셀파일로저장" class="button" onclick="excel('<%=seachbox%>','<%=eventbox%>');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oeventuserlist.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">ID</td>
	<!--<td align="center">주민번호</td>-->
	<td align="center">Sex</td>
	<td align="center">Name</td>
	<td align="center">Address</td>
	<td align="center">Level</td>
	<td align="center">Comment</td>
	<td align="center">Mail</td>
	<td align="center">Tel</td>
	<td align="center">Hp</td>		
	
	<% If eventbox = "9" Then %>
		<td align="center">당첨횟수</td>		
		<td align="center">최근당첨일</td>
	<% End If %>
	
	<td align="center">비고</td>
</tr>

<% if oeventuserlist.FResultCount > 0 then %>
<% for i=0 to oeventuserlist.FResultCount-1 %>
	<% if oeventuserlist.flist(i).finvaliduserid<>"" then %>
		<tr align="center" bgcolor="#e1e1e1">
	<% else %>
		<tr align="center" bgcolor="#FFFFFF">
	<% end if %>

		<td><%= printUserId(oeventuserlist.flist(i).fuserid, 2, "*") %></td>
		<!--<td><%'= left(oeventuserlist.flist(i).fjuminno,6) %>-->
		</td>
		<td>
			<% if mid(oeventuserlist.flist(i).fjuminno,8,1) = "1" then
				response.write "남성"
				else
				response.write "여성"
			end if %>
		</td>
		<td><%= oeventuserlist.flist(i).fusername %></td>
		<td align="left">
			[<%= printUserId(oeventuserlist.flist(i).fzipcode, 2, "*") %>]
			&nbsp;<%= printUserId(oeventuserlist.flist(i).faddress1, 2, "*") %>
			&nbsp;<%= printUserId(oeventuserlist.flist(i).fuseraddr, 2, "*") %>
		</td>
		<td><%= getUserLevelStr(oeventuserlist.flist(i).fLevel) %></td>
		<td><%= oeventuserlist.flist(i).fevtcom_txt %></td>
		<td><%= printUserId(oeventuserlist.flist(i).fusermail, 2, "*") %></td>
		<td><%= printUserId(oeventuserlist.flist(i).fuserphone, 2, "*") %></td>
		<td><%= printUserId(oeventuserlist.flist(i).fusercell, 2, "*") %></td> 

		<% If eventbox = "9" Then %>
			<td><%= oeventuserlist.flist(i).fWcnt %></td> 
			<td><%= oeventuserlist.flist(i).fWdate %></td>
		<% End If %>

		<td>
			<% if oeventuserlist.flist(i).finvaliduserid<>"" then %>
				불량고객
			<% end if %>
		</td>
	</tr>   
<% next %>

<%
'/컬쳐스테이션
If eventbox = "9" Then
%>
	<tr valign="bottom" bgcolor="FFFFFF">
		<td colspan="15" align="center">
		<%
			if oeventuserlist.HasPreScroll then
				Response.Write "<a href='javascript:goPage(" & oeventuserlist.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
			else
				Response.Write "[pre] &nbsp;"
			end if
	
			for i=0 + oeventuserlist.StartScrollPage to oeventuserlist.FScrollCount + oeventuserlist.StartScrollPage - 1
	
				if i>oeventuserlist.FTotalpage then Exit for
	
				if CStr(page)=CStr(i) then
					Response.Write " <font color='red'>[" & i & "]</font> "
				else
					Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
				end if
	
			next
	
			if oeventuserlist.HasNextScroll then
				Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
			else
				Response.Write "&nbsp; [next]"
			end if
		%>
		</td>
	</tr>
<% End If %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set oeventuserlist=nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->