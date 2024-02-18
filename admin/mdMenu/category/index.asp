<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mdMenu/catemanageCls.asp" -->
<%
Dim olist
Dim userid, dispCate, maxDepth, mduserid
Dim page, i, Depth1Code, Depth1Name
Dim searchUserid, MDUserList
maxDepth = 2

userid		= requestCheckvar(request("userid"),34)
dispCate	= requestCheckVar(Request("disp"),16)
page 		= requestCheckVar(Request("page"),2)
If page = "" Then page = 1

SET olist = new CMDCategory
	olist.FPageSize		= 500
	olist.FCurrPage		= 1
	olist.FRectUserid	= userid
	olist.FRectCatecode	= dispCate
	olist.getMDCategoryRegedList

	searchUserid	= DrawUserIdCombo("userid", userid)
	MDUserList		= DrawUserIdOption
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function form_check(f){
	f.submit();
}
function saveCateUserid(ctCode, mdid){
	if(mdid == ''){
		alert('담당자를 선택하세요');
		return;
	}
	if(ctCode.length == '3'){
		if(confirm("선택하신 1Depth이하 카테고리가 전부 담당자가 됩니다.\n\n저장하시겠습니까?")) {
			document.sfrm.catecode.value = ctCode;
			document.sfrm.mduserid.value = mdid;
			document.sfrm.mode.value = "I";
			document.sfrm.submit();
		}
	}else{
		if(confirm("저장 하시겠습니까?")) {
			document.sfrm.catecode.value = ctCode;
			document.sfrm.mduserid.value = mdid;
			document.sfrm.mode.value = "I";
			document.sfrm.submit();
		}
	}
}
function delCateUserid(ctCode){
	if(ctCode.length == '3'){
		if(confirm("선택하신 1Depth이하 카테고리가 전부 삭제 됩니다.\n\n삭제하시겠습니까?")) {
			document.sfrm.catecode.value = ctCode;
			document.sfrm.mode.value = "D";
			document.sfrm.submit();
		}
	}else{
		if(confirm("삭제 하시겠습니까?")) {
			document.sfrm.catecode.value = ctCode;
			document.sfrm.mode.value = "D";
			document.sfrm.submit();
		}
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page">
<tr align="center" bgcolor="#FFFFFF" height="50" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td bgcolor="<%= adminColor("gray") %>" align="left">
		담당자 : <%= searchUserid %>
		&nbsp;&nbsp;
		전시카테고리 : <!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색" onclick="form_check(this.form)">
	</td>
</tr>
</form>
</table>
<br><br>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="8" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="sfrm" method="POST" action="/admin/mdMenu/category/cateManage_process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="catecode">
<input type="hidden" name="mduserid">
<input type="hidden" name="mode">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%=olist.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25" >
	<td>카테고리코드</td>
    <td>카테고리1Depth</td>
    <td>카테고리2Depth</td>
    <td>담당자</td>
    <td>관리</td>
</tr>
<%
If olist.FResultCount > 0 Then
	For i = 0 to olist.FResultCount -1
		If olist.FItemList(i).FDepth = 1 Then 
			Depth1Code = olist.FItemList(i).FCatecode
			Depth1Name = olist.FItemList(i).FCatename
		End If
		
		If Depth1Code = "" Then
			Depth1Code = LEFT(olist.FItemList(i).FCatecode,3)
			Depth1Name = fnGet1DepthCode(left(olist.FItemList(i).FCatecode,3))
		End If
%>
<tr align="center" <%= Chkiif(olist.FItemList(i).FDepth="1","bgcolor=SKYBLUE","bgcolor=FFFFFF") %>  height="25">
	<td align="left"><%= olist.FItemList(i).FCatecode %></td>
    <td align="left"><%= Depth1Name %></td>
    <td align="left">
	<%
		If CStr(Depth1Code) = CStr(Left(olist.FItemList(i).FCatecode,3)) Then
			If olist.FItemList(i).FDepth <> 1 Then
				response.write olist.FItemList(i).FCatename
			End If
		End If
	%>
    </td>
    <td><%= olist.FItemList(i).FUsername %></td>
    <td>
    	<select name="mduserid<%=i%>" class="select">
    		<option value="">선택</option>
			<%= MDUserList %>
    	</select>
    	<input type="button" value="저장" class="button_s" onclick="saveCateUserid('<%= olist.FItemList(i).FCatecode %>', document.sfrm.mduserid<%=i%>.value)" >
		<% If olist.FItemList(i).FUsername <> "" Then %>
    	<input type="button" value="삭제" class="button_s" onclick="delCateUserid('<%= olist.FItemList(i).FCatecode %>')">
    	<% End If %>
    </td>
</tr>
<%
	Next 
Else
%>
<tr align="center" height="50" bgcolor="FFFFFF">
	<td colspan="5">데이터가 없습니다.</td>
</tr>
<% End If %>
</form>
</table>
<% SET olist = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->