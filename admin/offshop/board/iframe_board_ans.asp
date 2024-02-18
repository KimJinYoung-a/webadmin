<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인 통합 게시판
' History : 2010.06.18 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/board/board_cls.asp"-->

<link rel="stylesheet" href="/css/scm.css" type="text/css">
<%
Dim sRegistId, iDoc_Idx, iAns_Idx, sAns_Content , i	,page ,olect ,vParam ,olectview
	iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")
	iAns_Idx		= NullFillWith(requestCheckVar(Request("aidx"),10),"")
	sRegistId	 	= NullFillWith(requestCheckVar(Request("registid"),50),"")
	page = requestCheckVar(request("page"),10)

if page = "" then page = 1

vParam = "didx="&iDoc_Idx&"&aidx="&iAns_Idx&"&registid="&sRegistId

set olect = new clecturer_list
	olect.FPageSize = 20
	olect.FCurrPage = page
	olect.FrectDoc_Idx = iDoc_Idx
	olect.fnGetoboardList()

set olectview = new clecturer_list

If iAns_Idx <> "" Then	
	olectview.FrectAns_Idx = iAns_Idx
	olectview.fnGetoboardView()
	
	sAns_Content = olectview.foneitem.FAns_Content
	
	If sAns_Content = "" Then
		Response.Write "<script>alert('잘못된 접근입니다.');location.href='iframe_lecturer_ans.asp?didx="&iDoc_Idx&"&page="&page&"';</script>"
		dbget.close()
		Response.End
	End IF
End If
%>

<script language="javascript">

	function jsGoPage(iP){
		document.frmpage.iC.value = iP;
		document.frmpage.submit();
	}
	
	function ans_edit(aidx)
	{
		location.href = "iframe_board_ans.asp?didx=<%=iDoc_Idx%>&page=<%=page%>&aidx="+aidx+"&registid=<%=sRegistId%>";
	}
	
	function ans_del(aidx)
	{
		if(confirm("선택하신 글을 삭제하시겠습니까?") == true) {
			location.href = "board_ans_proc.asp?didx=<%=iDoc_Idx%>&page=<%=page%>&aidx="+aidx+"&del=o&registid=<%=sRegistId%>";
		} else {
			return false;
		}
	}
	
	function checkform(frm)
	{
		if (frm.ans_content.value == "")
		{
			alert("답변을 입력하세요!");
			frm.ans_content.focus();
			return false;
		}
	}
	
</script>

<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="board_ans_proc.asp" method="post" onSubmit="return checkform(this);">
<input type="hidden" name="didx" value="<%=iDoc_Idx%>">
<input type="hidden" name="aidx" value="<%=iAns_Idx%>">
<input type="hidden" name="registid" value="<%=sRegistId%>">
<input type="hidden" name="page" value="<%=page%>">
<tr bgcolor="#FFFFFF">
	<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">답변내용</td>
	<td align="left"><textarea class="textarea" name="ans_content" cols="112" rows="5"><%=sAns_Content%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="right">
		<input type="submit" value="답변저장" class="button">
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center" width=220>작성자</td>
	<td align="center">내&nbsp;&nbsp;&nbsp;용</td>
</tr>
<%
IF olect.fresultcount > 0 THEN
	
For i =0 To olect.fresultcount -1
%>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="center" valign="top" style="padding:3 0 0 3">
		<%
		if olect.FItemList(i).fusername <> "" then
			Response.Write olect.FItemList(i).fusername
		else
			Response.Write olect.FItemList(i).fcompany_name
		end if
		
		Response.Write "(" & olect.FItemList(i).fid & ")"
		Response.Write "<br>" & olect.FItemList(i).fans_regdate
		If olect.FItemList(i).fid = session("ssBctId") Then
			Response.Write "<br><img src='http://fiximage.10x10.co.kr/web2009/common/cmt_modify.gif' style='cursor:pointer' onClick='ans_edit(" & olect.FItemList(i).fans_idx & ")'>"
			Response.Write "&nbsp;<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' style='cursor:pointer' onClick='ans_del(" & olect.FItemList(i).fans_idx & ")'>"
		elseif olect.FItemList(i).fid = session("ssBctId") or (C_ADMIN_AUTH) Then			
			Response.Write "<br><img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' style='cursor:pointer' onClick='ans_del(" & olect.FItemList(i).fans_idx & ")'>"
		End If		
		%>
	</td>
	<td align="left" style="padding:3 3 3 3"><%=replace(olect.FItemList(i).fans_content,vbCrLf,"<br>")%></td>
</tr>
<% Next %>
<tr>
	<td colspan="2">
		<!-- 페이징처리 -->
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="15" align="center">
		       	<% if olect.HasPreScroll then %>
					<span class="list_link"><a href="?page=<%= olect.StartScrollPage-1 %>&vParam=<%=vParam%>">[pre]</a></span>
				<% else %>
				[pre]
				<% end if %>
				<% for i = 0 + olect.StartScrollPage to olect.StartScrollPage + olect.FScrollCount - 1 %>
					<% if (i > olect.FTotalpage) then Exit for %>
					<% if CStr(i) = CStr(olect.FCurrPage) then %>
					<span class="page_link"><font color="red"><b><%= i %></b></font></span>
					<% else %>
					<a href="?page=<%= olect.StartScrollPage-1 %>&vParam=<%=vParam%>" class="list_link"><font color="#000000"><%= i %></font></a>
					<% end if %>
				<% next %>
				<% if olect.HasNextScroll then %>
					<span class="list_link"><a href="?page=<%= olect.StartScrollPage-1 %>&vParam=<%=vParam%>">[next]</a></span>
				<% else %>
				[next]
				<% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<% Else %>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="2" align="center" class="page_link">[답변이 없습니다.]</td>
		</tr>
<%
	End If
%>
</table>

<%
set olect = nothing
set olectview = nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->