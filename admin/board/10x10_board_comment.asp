<!-- #include virtual="/lib/classes/10x10_board_commentcls.asp" -->
<%
 Dim ocomm,cpage,iy,ix

  cpage = request("cpage")
  if cpage = "" then cpage = 1

  set ocomm = new CComment
  ocomm.FPageSize = 500
  ocomm.FCurrPage = cpage
  ocomm.comment_list idx

%>
<script language="JavaScript">
<!--
   function sendcomm(){
     if (document.commform.userid.value == ""){
	   alert("로그인 먼저 해주세요!");
       document.commform.comment.focus();
	 } 
     else if (document.commform.comment.value == ""){
	   alert("코멘트를 써주세요!");
       document.commform.comment.focus();
	 }
	 else if(confirm("코멘트를 저장하시겠습니까?")){
	   document.commform.submit();
	 }
   
   }

//-->
</script>
<% if ocomm.FResultCount<1 then %>
<% else %>
<table width="650" border="0" cellspacing="0" cellpadding="0" class="a">
  <tr> 
	<td colspan="2" style="border-bottom:1px solid #CDCDCD;border-top:1px solid #CDCDCD; padding:0 0 0 0">*****코멘트 리스트*****</td>
  </tr>
<% for iy =0 to ocomm.FResultCount -1 %>
  <tr> 
	<td valign="top" class="p12" style="padding:5 0 5 0"><font color="#393939"><% = nl2br(ocomm.FItemList(iy).Fcomment) %></font>
	</td>
	<td align="right" valign="top" class="verdana-xsmall" style="padding:6 6 6 6" width="100">
	  <% =ocomm.FItemList(iy).Fusername %>(<% =ocomm.FItemList(iy).Fuserid %>)<br><span class="verdana-small"><% = FormatDateTime(ocomm.FItemList(iy).Fregdate,2) %></span>
	 <% if ocomm.FItemList(iy).Fuserid = session("ssBctId") then %>
		<a href="/admin/board/lib/comment_delete_ok.asp?idx=<% =ocomm.FItemList(iy).Fidx %>"><font color="#808080">삭제</font></a>
	 <% end if %>
	 </td>
  </tr>
  <tr> 
	<td colspan="2" align="center" class="verdana-small" style="padding:0 0 0 0; border-top:1px solid #E6E6E6"><img src="/images/spacer.gif" width="100%" height="1"></td>
  </tr>
<% next %>
</table>
<% end if %>
<table width="650" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td align="center">
		  <% if ocomm.HasPreScroll then %>
			  <a href="?cpage=<%= CStr(ocomm.StarScrollPage-1) %>&idx=<% = idx %>&page=<% =page %>">
				◁</a>
		  <% else %> 
		  <% end if %>		
		  <% for ix = ocomm.StarScrollPage to (ocomm.StarScrollPage + ocomm.FScrollCount - 1) %>
			  <% if (ix > ocomm.FTotalPage) then Exit For %>
			  <% if CStr(ix) = CStr(ocomm.FCurrPage) then %>
				[<%= CStr(ix) %>]
			  <% else %>
				<a href="?cpage=<%= CStr(ix) %>&idx=<% = idx %>&page=<% =page %>" class="page_link">
				<%= CStr(ix) %>
				</a>
			  <% end if %>
		  <% next %>
		  <% if ocomm.HasNextScroll then %>
			<a href="?cpage=<%= CStr(ocomm.StarScrollPage + ocomm.FScrollCount) %>&idx=<% = idx %>&page=<% =page %>">▷</a>
		  <% else %>
		  <% end if %>	
	</td>
</tr>
</table>
<table width="650" border="0" cellspacing="0" cellpadding="0" class="a">
<form name="commform" method="post" action="/admin/board/lib/comment_input_ok.asp">
<input type="hidden" name="NextPage" value="<% = Request.ServerVariables("Path_Info") & "?" &  Request.ServerVariables("Query_String") %>">
<input type="hidden" name="userid" value="<%= session("ssBctId") %>">
<input type="hidden" name="username" value="<%= session("ssBctCname") %>">
<input type="hidden" name="idx" value="<% = idx %>">
 <tr> 
	<td colspan="2" class="p11" style="padding:6 6 0 6; border-top:1px solid #666666"><font color="#FF6600">간단한 코멘트를 달아주세요!</font></td>
  </tr>
  <tr> 
	<td valign="top"> 
	  <textarea name="comment" cols="80" rows="6"></textarea>
	</td>
	<td valign="top">
	<a href="javascript:sendcomm();" onfocus="this.blur();">코멘트 쓰기</a></td>
  </tr>
 </form>
</table>
<br><br>