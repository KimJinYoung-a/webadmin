<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이미지링크관리
' History : 2019.08.06 원승현 : 신규작성
'			2022.07.07 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- # include virtual="/admin/pcmain/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/imageLinkCls.asp" -->
<%
	Dim isusing , dispcate
	dim page 
	Dim i
	dim imageLinkList
	Dim sDt , modiTime

	page = requestCheckVar(getNumeric(request("page")),10)
	dispcate = request("disp")
	isusing = requestCheckVar(request("isusing"),1)
	sDt = request("prevDate")

	if page="" then page=1
	If isusing = "" Then isusing ="Y"

	set imageLinkList = new CimageLink
	imageLinkList.FPageSize		= 20
	imageLinkList.FCurrPage		= page
	imageLinkList.Fisusing			= isusing
	imageLinkList.GetContentsList()

%>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script type='text/javascript'>
<!--
//수정
function jsmodify(v){
	location.href = "linkimageinsert.asp?menupos=<%=menupos%>&idx="+v+"&paramisusing=<%=isusing%>";
}
$(function() {
  	$("input[type=submit]").button();

  	// 라디오버튼
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");
	
});
-->
</script>
<!-- 검색 시작 -->
<form name="frm" method="post" style="margin:0px;" action="/admin/sitemaster/ImageLinkMap/index.asp">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			* 사용여부 :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			<!--지정일자 <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>-->
			</div>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="submit" class="button_s" value="검 색">
		</td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
<div style="float:right;clear:both;"><a href="" onclick="window.open('popimagelinkedit.asp','imagelinkedit','width=1200,height=700,scrollbars=yes,resizable=yes');return false;"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a></div>
<br><br>
<!--  리스트 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
		총 등록수 : <b><%=imageLinkList.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=imageLinkList.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
	<td width="10%">타이틀</td>
	<td width="12%">이미지</td>
    <td width="5%">등록일</td>
    <td width="10%">등록자</td>
    <td width="10%">최종수정자</td>
	<td width="10%">등록자이미지</td>
	<td width="10%">프론트노출이름</td>
    <td width="10%">사용여부</td>
	<td width="10%">기타</td>
</tr>
<% 
	for i=0 to imageLinkList.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(imageLinkList.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=imageLinkList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=imageLinkList.FItemList(i).Fidx%></td>
    <td onclick="jsmodify('<%=imageLinkList.FItemList(i).Fidx%>');" style="cursor:pointer;">
		<%= ReplaceBracket(imageLinkList.FItemList(i).Ftitle) %>
	</td>
	<td onclick="jsmodify('<%=imageLinkList.FItemList(i).Fidx%>');" style="cursor:pointer;"><img src="<%=imageLinkList.FItemList(i).FImage%>" width="70%"></td>
	<td><%=left(imageLinkList.FItemList(i).Fregdate,10)%></td>
	<td><%=imageLinkList.FItemList(i).Fusername%></td>
	<td>
		<%
			modiTime = imageLinkList.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write imageLinkList.FItemList(i).Fusername2 & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
	<td><img src="<%=imageLinkList.FItemList(i).FRegUserImage%>" width="70%"></td>
	<td><%= ReplaceBracket(imageLinkList.FItemList(i).FRegUserFrontName) %></td>
    <td><%=chkiif(imageLinkList.FItemList(i).Fisusing="N","사용안함","사용함")%></td>
	<td><input type="button" value="수정" onclick="window.open('popimagelinkedit.asp?menupos=<%=menupos%>&idx=<%=imageLinkList.FItemList(i).Fidx%>','imagelinkedit','width=1200,height=700,scrollbars=yes,resizable=yes');return false;" class="button" />&nbsp;&nbsp;<input type="button" value="상품등록" onclick="jsmodify('<%=imageLinkList.FItemList(i).Fidx%>');" class="button" /></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if imageLinkList.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= imageLinkList.StartScrollPage-1 %>&isusing=<%=isusing%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + imageLinkList.StartScrollPage to imageLinkList.StartScrollPage + imageLinkList.FScrollCount - 1 %>
				<% if (i > imageLinkList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(imageLinkList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if imageLinkList.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set imageLinkList = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->