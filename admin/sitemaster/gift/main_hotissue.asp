<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : GIFT ���� HOT ISSUE ����
' Hieditor : ������ ����
'			 2022.07.08 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftmain_cls.asp" -->
<%
dim page, i
	page = requestCheckVar(getNumeric(request("page")),10)
	if page = "" then page = 1 end if
	
dim cGift
	set cGift = new Cgift_list
	cGift.FPageSize = 15
	cGift.FCurrPage = page
	cGift.FRectIsusing = "Y"
	cGift.FRectIsOpen = "Y"

	cGift.sbHotIssueList
%>
<script type='text/javascript'>

function NextPage(p){
	frm1.page.value = p;
	frm1.submit();
}

function talkhotissue(i){
	var talkhotissuepop = window.open('main_hotissue_write.asp?idx='+i+'','talkhotissuepop','width=1200,height=768,scrollbars=yes,resizable=yes');
	talkhotissuepop.focus();
}
</script>

<form name="frm1" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=request("menupos")%>">
<input type="hidden" name="page" value="">
</form>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		���Ĺ�ȣ���� : 0�� ���� ��, �׸���ȣ�� �ֱ��ϼ��� ��
	</td>
	<td align="right">	
		<input type="button" value="���۾���" onClick="talkhotissue('')" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%=cGift.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=cGift.FtotalPage%></b>
	</td>
</tr>
<tr height="30" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>���Ĺ�ȣ</td>
    <td>�׸�idx</td>
    <td>�׸�����</td>
    <td>���� ~ ����</td>
    <td>��������</td>
    <td></td>
</tr>
<%
	for i=0 to cGift.FResultCount - 1
%>
<tr bgcolor="#FFFFFF" height="30">
    <td align="center"><%=cGift.FItemList(i).Fsortno%></td>
    <td align="center"><%=cGift.FItemList(i).FthemeIdx%></td>
    <td align="center"><%= ReplaceBracket(cGift.FItemList(i).Fsubject) %></td>
    <td align="center"><%=Left(cGift.FItemList(i).Fstartdate,10)%> ~ <%=Left(cGift.FItemList(i).Fenddate,10)%></td>
    <td align="center"><%=CHKIIF(cGift.FItemList(i).Fisusing="Y","�����","����ó����")%></td>
	<td align="center">
		[<a href="<%=wwwUrl%>/gift/shop/themeView.asp?themeIdx=<%=cGift.FItemList(i).FthemeIdx%>" target="_blank">�� ��</a>]&nbsp;&nbsp;&nbsp;
		[<a href="javascript:talkhotissue('<%=cGift.FItemList(i).Fidx%>');">�� ��</a>]
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" height="30">
    <td colspan="15" align="center">
    <% if cGift.HasPreScroll then %>
		<a href="javascript:NextPage('<%= cGift.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + cGift.StartScrollPage to cGift.FScrollCount + cGift.StartScrollPage - 1 %>
		<% if i>cGift.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if cGift.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>

<% Set cGift = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->