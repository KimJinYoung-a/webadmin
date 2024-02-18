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
Dim userid, dispCate, maxDepth
Dim page, i
maxDepth = 2

userid		= requestCheckvar(request("userid"),34)
dispCate	= requestCheckVar(Request("disp"),16)
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function form_check(f){
	if (f.userid.value == ""){
		alert('담당자를 선택하세요');
		f.userid.focus();
		return;
	}
	f.submit();
}
function chgMD(uid){
	location.replace('/admin/mdMenu/category/popCatemanager.asp?userid='+uid);
}
</script>
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="itemreg" method="post" action="/admin/mdMenu/category/cateManage_process.asp" onsubmit="return false;" style="margin:0;">
<tr align="center" bgcolor="FFFFFF"  height="25">
    <td>담당자</td>
    <td colspan="3" align="left"><% DrawUserIdCombo "userid", userid, "chgMD(this.value);" %></td>
</tr>
<tr align="left">
	<td align="center" height="30" width="15%" bgcolor="FFFFFF" title="프론트에 진열될 카테고리" style="cursor:help;">카테고리</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
		<input type="button" value="저장" onclick="form_check(this.form)" class="button_s">
	</td>
</tr>
</form>
</table>
<br><br>
<%
Dim olist, regMd, regCode
SET olist = new CMDCategory
	olist.FPageSize = 500
	olist.FCurrPage = 1
	olist.getMDCategoryRegedList
%>
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="FFFFFF"  height="25">
    <td>카테고리</td>
    <td>Depth</td>
    <td>카테고리 명</td>
    <td>담당자</td>
</tr>
<% For i = 0 to olist.FResultCount -1 %>
<tr align="center" height="25" <%= Chkiif(olist.FItemList(i).FDepth="1","bgcolor=SKYBLUE","bgcolor=FFFFFF") %> >
	<td><%=olist.FItemList(i).FCatecode %></td>
	<td><%=olist.FItemList(i).FDepth %></td>
	<td><%= olist.FItemList(i).FCatename %></td>
	<td>
	<%
		If olist.FItemList(i).FDepth = 1 AND olist.FItemList(i).FUserid <> "" Then 
			regCode = olist.FItemList(i).FCatecode
			regMd	= olist.FItemList(i).FUsername
		End If

		If CStr(LEFT(olist.FItemList(i).FCatecode,3)) = CStr(regCode) Then
			response.write regMd
		Else
			response.write olist.FItemList(i).FUsername
		End IF
	%>
	</td>
</tr>
<% Next %>
</table>
<%
SET olist = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->