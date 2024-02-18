<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->
<%
dim CateCode , yearUse , isusing, mdpick, i, page
	page = requestCheckVar(request("page"),5)

if page = "" then page = 1

dim oDiary
set oDiary = new DiaryCls
	oDiary.FPageSize = 20
	oDiary.FCurrPage = page
	oDiary.getDiaryOneplusOne_List
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<script type="text/javascript">
	function popRegNew(mode,idx){
		var popRegNew = window.open('/admin/diary2009/pop_OneplusOne_reg.asp?mode='+mode+'&idx='+idx,'popRegNew','width=1024,height=768,status=yes,scrollbars=yes')
		popRegNew.focus();
	}
	document.domain ='10x10.co.kr';
</script>
</head>
<body>
<div class="contSectFix scrl">
	<div class="pad20">
		<table class="tbType1 listTb">
		<tr>
			<td style="text-align:left;">
				<input type="button" name="ins" value="신규등록" class="button_s" onclick="popRegNew('add','');">
				<input type="button" name="ins" value="닫 기" class="button_s" onclick="window.close();">
			</td>
		</tr>
		</table>
		<div class="tPad15">
			<table class="tbType1 listTb">
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td nowrap> 번호</td>
					<td nowrap> 상품번호 </td>
					<td nowrap>1+1 이미지(좌)-품절 전</td>
					<td nowrap>1+1 이미지-품절 후</td>
					<td nowrap> 시작일 </td>
					<td nowrap> 사용여부 </td>
				</tr>
				<% For i =0 To  oDiary.FResultCount -1 %>
				<tr align="center" <% if oDiary.FItemList(i).FIsusing = "Y" or oDiary.FItemList(i).FStartdate >= date() then %>bgcolor="#FFFFFF"<% else %>bgcolor="#DDDDDD"<% end if %> onclick="popRegNew('edit','<%=oDiary.FItemList(i).FIdx%>');" style="cursor:hand;">
					<td nowrap><%= oDiary.FItemList(i).FIdx %></td>
					<td nowrap><%= oDiary.FItemList(i).FItemid %></td>
					<td nowrap><img src="<%=uploadUrl%>/diary/oneplusone/<%= oDiary.FItemList(i).FImage1 %>" border="0" width="50" height="50">	</td>
					<td nowrap><img src="<%=uploadUrl%>/diary/oneplusone/<%= oDiary.FItemList(i).FImageEnd %>" border="0" width="50" height="50">	</td>
					<td nowrap><%= oDiary.FItemList(i).FStartdate %><%=chkiif(oDiary.FItemList(i).FStartdate < date(),"</br><span style='color:red;'>(종료)</span>","")%></td>
					<td nowrap><%= oDiary.FItemList(i).FIsusing %></td>
				</tr>
				<%Next%>
				<tr bgcolor="#FFFFFF">
					<td colspan="12" align="center">
						<a href="?page=1&isusingbox=<%=isusing%>&cate=<%=catecode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/pprev_btn.gif" width="10" height="10" border="0"></a>
						<% if oDiary.HasPreScroll then %>
							<span class="list_link"><a href="?page=<%= oDiary.StartScrollPage-1 %>&isusingbox=<%=isusing%>&cate=<%=catecode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
						<% else %>
						&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;
						<% end if %>
						<% for i = 0 + oDiary.StartScrollPage to oDiary.StartScrollPage + oDiary.FScrollCount - 1 %>
							<% if (i > oDiary.FTotalpage) then Exit for %>
							<% if CStr(i) = CStr(oDiary.FCurrPage) then %>
							<span class="page_link"><font color="red"><b><%= i %>&nbsp;&nbsp;</b></font></span>
							<% else %>
							<a href="?page=<%= i %>&isusingbox=<%=isusing%>&cate=<%=catecode%>" class="list_link"><font color="#000000"><%= i %>&nbsp;&nbsp;</font></a>
							<% end if %>
						<% next %>
						<% if oDiary.HasNextScroll then %>
							<span class="list_link"><a href="?page=<%= i %>&isusingbox=<%=isusing%>&cate=<%=catecode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
						<% else %>
						&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;
						<% end if %>
						<a href="?page=<%= oDiary.FTotalpage %>&isusingbox=<%=isusing%>&cate=<%=catecode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/nnext_btn.gif" width="10" height="10" border="0"></a>
					</td>
				</tr>
			</table>
		</div>
	</div>
</div>
<% set oDiary = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->