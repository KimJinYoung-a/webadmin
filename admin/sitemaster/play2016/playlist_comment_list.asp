<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/play/play2016Cls.asp" -->
<%
	Dim i, cPl, vDidx, vPage, vTitle, vQuery, vCommTitle, vComment1, vComment2, vComment3
	vDidx = requestCheckVar(request("didx"),10)
	vTitle = requestCheckVar(request("title"),150)
	vPage = NullFillWith(requestCheckVar(request("page"),10),1)

	vQuery = "select comm_title, comment1, comment2, comment3 from [db_giftplus].[dbo].[tbl_play_playlist] where didx = '" & vDidx & "'"
	rsget.CursorLocation = adUseClient
	rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
	If not rsget.eof Then
		vCommTitle = rsget("comm_title")
		vComment1 = rsget("comment1")
		vComment2 = rsget("comment2")
		vComment3 = rsget("comment3")
	End If
	rsget.close

	SET cPl = New CPlay
	cPl.FRectDIdx = vDidx
	cPl.FCurrPage = vPage
	cPl.FPageSize = 20
	cPl.fnGetPlayPlaylistUser
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery_common.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<script>
function searchFrm(p){
	frm1.page.value = p;
	frm1.submit();
}
function jsDelete(i){
	if(confirm("" + i + " 번 글을 삭제하시겠습니까?\n\n※ 삭제하면 복구가 불가능 합니다.") == true) {
		frmDel.idx.value = i;
		frmDel.submit();
		return true;
	}else{
		return false;
	}
}
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
</head>
<body>
<div class="contSectFix scrl">
	<form name="frmDel" method="get" action="comment_delete.asp" style="margin:0px;">
	<input type="hidden" name="cate" value="1">
	<input type="hidden" name="didx" value="<%=vDidx%>">
	<input type="hidden" name="idx" value="">
	</form>
	<form name="frm1" method="get" action="" style="margin:0px;">
	<input type="hidden" name="didx" value="<%=vDidx%>">
	<input type="hidden" name="title" value="<%=vTitle%>">
	<input type="hidden" name="page" value="<%=vPage%>">
	<div class="pad20">
		<div class="overHidden">
			<div class="ftLt">
				<p class="cBk1 ftLt">* 총 <%=cPl.FTotalCount%> 개</p>
			</div>
			<div class="ftRt">
				<p class="cBk1 ftLt">didx : <%=vDidx%>, 주제 : <%=vCommTitle%></p>
			</div>
		</div>
		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div>idx</div></th>
					<th><div>작성글</div></th>
					<th><div>작성인</div></th>
					<th><div>작성일</div></th>
					<th><div></div></th>
				</tr>
				</thead>
				<tbody>
				<%
					If cPl.FResultCount > 0 Then
						For i=0 To cPl.FResultCount-1
				%>
						<tr onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
							<td><%=cPl.FItemList(i).Fidx%><br /><br />[<%=cPl.FItemList(i).Fdevice%>]</td>
							<td style="text-align:left;">
								#<%=cPl.FItemList(i).FCate1Comment1%> <%=vComment1%><br />
								#<%=cPl.FItemList(i).FCate1Comment2%> <%=vComment2%>
								<% If vComment3 <> "" Then %>
									<br />#<%=cPl.FItemList(i).FCate1Comment3%> <%=vComment3%>
								<% End If %>
							</td>
							<td><%=cPl.FItemList(i).Fuser%></td>
							<td><%=cPl.FItemList(i).Fregdate%></td>
							<td>[<a href="" onClick="jsDelete('<%=cPl.FItemList(i).Fidx%>');return false;">삭제</a>]</td>
						</tr>
				<%
						Next
					End If
				%>
				</tbody>
			</table>
			<br />
			<div class="ct tPad20 cBk1">
				<% if cPl.HasPreScroll then %>
				<a href="javascript:searchFrm('<%= cPl.StartScrollPage-1 %>')">[pre]</a>
				<% else %>
	    			[pre]
	    		<% end if %>
	    		
	    		<% for i=0 + cPl.StartScrollPage to cPl.FScrollCount + cPl.StartScrollPage - 1 %>
	    			<% if i>cPl.FTotalpage then Exit for %>
	    			<% if CStr(vPage)=CStr(i) then %>
	    			<span class="cRd1">[<%= i %>]</span>
	    			<% else %>
	    			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
	    			<% end if %>
	    		<% next %>
				
				<% if cPl.HasNextScroll then %>
	    			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
	    		<% else %>
	    			[next]
	    		<% end if %>
			</div>
		</div>
	</div>
	</form>
</div>

<% SET cPl = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->