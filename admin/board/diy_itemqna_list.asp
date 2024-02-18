<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/upcheitemqna/diy_item_qnacls.asp"-->
<%
dim notupbea, imageon, mifinish, makerid, research ,page
mifinish = request("mifinish")
'makerid = session("ssBctID")
research = request("research")
page=request("page")
if page="" then page=1

if research="" and mifinish="" then mifinish="on"

dim itemqna
set itemqna = new CItemQna
itemqna.FPageSize = 20
itemqna.FCurrPage=page
itemqna.FReckMiFinish = mifinish
itemqna.FRectMakerid = makerid
itemqna.ItemQnaList

dim i
%>

<script>
function NextPage(pg){
	document.location.href="/designer/board/diy_itemqna_list.asp?mifinish=<%=mifinish %>&research=<%=research %>&page="+ pg;
}
</script>


<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	<input type=checkbox name=mifinish <% if mifinish="on" then response.write "checked" %> >��ó�����˻�
	        </td>
	        <td align="right">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->



<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr bgcolor="<%= adminColor("tabletop") %>">
    <td width="150" align="center">����(���̵�)</td>
    <td align="center">����..</td>
    <td width="60" align="center">��ǰID</td>
    <td width="80" align="center">�귣��</td>
    <td width="60" align="center">��۱���</td>
    <td width="80" align="center">�ۼ���</td>
    <td width="80" align="center">�亯��</td>
    <td width="80" align="center">�亯��</td>
  </tr>
<% for i = 0 to (itemqna.FResultCount - 1) %>
  <tr bgcolor="#FFFFFF" >
    <td>&nbsp;<%= itemqna.FItemList(i).Fusername %>(<%= itemqna.FItemList(i).Fuserid %>)</td>
    <td>&nbsp;<a href="diy_itemqna_view.asp?id=<%= itemqna.FItemList(i).Fid %>&menupos=<%= menupos %>"><%= db2html(itemqna.FItemList(i).Ftitle) %></a></td>
    <td align="center"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= itemqna.FItemList(i).FItemID %>" target=_blank><%= itemqna.FItemList(i).FItemID %></a></td>
    <td align="center"><%= itemqna.FItemList(i).Fmakerid %></td>
    <td align="center"><%= itemqna.FItemList(i).GetDeliveryTypeName %></td>
    <td align="center"><%= FormatDate(itemqna.FItemList(i).Fregdate, "0000-00-00") %></td>
    <td align="center"><%= itemqna.FItemList(i).Freplyuser %></td>
    <td align="center">
    <% if Not IsNULL(itemqna.FItemList(i).FReplydate) then %>
    <%= FormatDate(itemqna.FItemList(i).FReplydate, "0000-00-00") %>
    <% end if %>
    </td>
  </tr>
<% next %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if itemqna.HasPreScroll then %>
				<a href="javascript:NextPage('<%= CStr(itemqna.StartScrollPage - 1) %>')">[prev]</a>
			<% else %>
				[prev]
			<% end if %>
			<% for i = itemqna.StartScrollPage to (itemqna.StartScrollPage + itemqna.FScrollCount - 1) %>
			  <% if (i > itemqna.FTotalPage) then Exit For %>
			  <% if CStr(i) = CStr(itemqna.FCurrPage) then %>
				 [<%= i %>]
			  <% else %>
				 <a href="javascript:NextPage('<%= i %>')" class="id_link">[<%= i %>]</a>
			  <% end if %>
			<% next %>
			<% if itemqna.HasNextScroll then %>
				<a href="javascript:NextPage('<%= CStr(itemqna.StartScrollPage + itemqna.FScrollCount) %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->


<%
set itemqna = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->