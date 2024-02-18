<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 오프라인 뉴스
' History : 최초생성자모름
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/classes/board/offshop_newscls.asp" -->
<%
dim i, j, page, shopid, isusing, research
	page        = requestCheckVar(request("page"),10)
	shopid      = requestCheckVar(request("shopid"),32)
	isusing     = requestCheckVar(request("isusing"),1)
	research    = requestCheckVar(request("research"),2)

if page="" then page=1
if (research="") and (isusing="") then isusing="Y"

dim offnews
set offnews = New COffshopNewsEvent
offnews.FRectIsusing = isusing
offnews.FRectShopid = shopid
offnews.FPageSize = 20
offnews.FCurrPage = page
offnews.FScrollCount = 10
offnews.GetOffshopNewsList

%>
<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 9pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>
<script language='javascript'>
function  TnSearch(frm){
	if (frm.rectuserid.length<1){
		alert('검색어를 입력하세요.');
		return;
	}
	frm.method="get";
	frm.submit();
}
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
			샾 : <% drawSelectBoxOffShopAll "shopid",shopid %> &nbsp;&nbsp;
			사용구분 :
			<select name="isusing" class="select" >
			    <option value="">ALL
			    <option value="Y" <%= chkIIF(isusing="Y","selected","") %> >Y
			    <option value="N" <%= chkIIF(isusing="N","selected","") %> >N
			</select>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="1" bordercolordark="White" bordercolorlight="black" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="50" align="center">번호</td>
    <td width="100" align="center">샵</td>
    <td width="100" align="center">구분</td>
    <td align="center">제목</td>
    <td width="100" align="center">작성자</td>
    <td width="100" align="center">작성일</td>
    <td width="80" align="center">사용유무</td>
    <td width="100" align="center">유효기간</td>
  </tr>
<% for i = 0 to (offnews.FResultCount - 1) %>
  <tr height="20">
    <td align="center">&nbsp;<%= offnews.FItemList(i).Fidx %></td>
    <td align="center"><%= offnews.FItemList(i).Fshopname %></td>
	<td align="center"><%= fnGetCommonCode("noticegubun",offnews.FItemList(i).Fgubun)%></td>
    <td>&nbsp;<a href="offshop_news_event_edit.asp?idx=<%= offnews.FItemList(i).Fidx %>&menupos=<%= menupos %>"><%= offnews.FItemList(i).Ftitle %></a></td>
    <td align="center"><%= offnews.FItemList(i).Fuserid %></td>
    <td align="center"><%= FormatDate(offnews.FItemList(i).Fregdate, "0000.00.00") %></td>
    <td align="center">
	<% if offnews.FItemList(i).Fisusing ="Y" then %>
	Y
	<% else %>
	<font color="red">N</font>
	<% end if %>
    </td>
    <td align="center"><%= offnews.FItemList(i).Fenddate %></td>
  </tr>
<% next %>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td align="center" height="30">
		<% if offnews.HasPreScroll then %>
			<a href="javascript:NextPage('<%= offnews.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + offnews.StartScrollPage to offnews.FScrollCount + offnews.StartScrollPage - 1 %>
			<% if i>offnews.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if offnews.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<tr>
	<td align="right"><a href="offshop_news_event_write.asp"><font color="red">News & 이벤트 등록</font></a>&nbsp;&nbsp;&nbsp;</td>
</tr>
</table>
<br><br>
<% set offnews = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->