<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� �Ǹ� ��� ����
' Hieditor : 2011.04.22 �̻� ����
'			 2012.08.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/search/recomKeywordItemCls.asp" -->
<%
Dim i
Dim group_no : group_no = requestCheckvar(Trim(request("group_no")),10)
Dim keyword : keyword = requestCheckvar(Trim(request("keyword")),50)
Dim page : page = requestCheckvar(Trim(request("page")),10)

if (page="") then page=1

Dim oRecomKeywordItem

set oRecomKeywordItem = new CRecomKeywordItem
oRecomKeywordItem.FPageSize = 50
oRecomKeywordItem.FCurrPage = page
oRecomKeywordItem.FRectGroup_no = group_no

if (group_no<>"") then
oRecomKeywordItem.getRecomKeywordItemList
end if
%>
<script language="javascript">
function NextPage(i){
    document.frm.page.value=i;
    document.frm.submit();
}

function AddRecomKeywordItem() {
	var frm = document.frmadditem;

	if (frm.itemid.value.length<4) {
		alert('��ǰ�ڵ带 �Է��ϼ���.');
        frm.itemid.focus();
		return;
	}

	if (frm.group_no.value.length<1) {
		alert('�ش� �׷��ȣ�� �������� �ʾҽ��ϴ�.');
		return;
	}

	if (confirm('��ǰ�� �߰� �Ͻðڽ��ϱ�?') == true) {
		frm.submit();
	}
}

function delItem(group_no,itemid){
    var frm = document.frmDel;

    if (confirm("�ش� ��ǰ�� �����Ͻðڽ��ϱ�?")){
        frm.group_no.value=group_no;
        frm.itemid.value=itemid;
        frm.submit();
    }
    
}
</script>
<!-- �׼� ���� -->
<p>
<form name="frm" method="get">
<input type="hidden" name="group_no" value="<%=group_no%>">
<input type="hidden" name="keyword" value="<%=keyword%>">
<input type="hidden" name="page" value="<%=page%>">
</form>
<p>
<form name="frmadditem" method="post" action="keywordRecom_Process.asp">
<input type="hidden" name="mode" value="additem">
<input type="hidden" name="group_no" value="<%=group_no%>">
<input type="hidden" name="keyword" value="<%=keyword%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			Ű���� : <strong><%=keyword %></strong>
		</td>
		<td align="right">
		    ��ǰ�ڵ�:<input type="text" name="itemid" value="" size="10" maxlength="10">
		    <input type="button" class="button" value="��ǰ �߰�" onClick="AddRecomKeywordItem()">
			&nbsp;
		</td>
	</tr>
</table>
</form>
<!-- �׼� �� -->
<p>

<table width="100%" align="center" cellpadding="4" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
	    <td colspan="20">
    		�˻���� : <b><%= oRecomKeywordItem.FTotalcount %></b>
    		&nbsp;
    		������ : <b><%= page %> / <%= oRecomKeywordItem.FTotalPage %></b>
    	</td>
    </tr>
    
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="80" height="22" >��ǰ�ڵ�</td>
    	<td width="50">�̹���</td>
    	<td width="100">�귣��ID</td>
		<td width="200">��ǰ��</td>
		
        <td width="100">�ǸŰ�</td>
        <td width="100">���԰�</td>
        <td width="100">���Ա���</td>
        <td width="70">�Ǹſ���</td>
        <td width="70">��뿩��</td>
        <td width="90">��������</td>
		<td width="80">����</td>
	</tr>
	<%
	for i = 0 To oRecomKeywordItem.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
	    <td height="22" ><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oRecomKeywordItem.FItemList(i).Fitemid %>" target="_blank" title="�̸�����"><%= oRecomKeywordItem.FItemList(i).Fitemid %></a></td>
		<td><img src="<%= oRecomKeywordItem.FItemList(i).Fsmallimage %>"></td>
        <td align="left"><%= oRecomKeywordItem.FItemList(i).Fmakerid %></td>
        <td align="left"><%= oRecomKeywordItem.FItemList(i).Fitemname %></td>
        
		<td align="right">
        <%
            Response.Write FormatNumber(oRecomKeywordItem.FItemList(i).Forgprice,0) 
			'���ΰ�
			if oRecomKeywordItem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>("&CLng((oRecomKeywordItem.FItemList(i).Forgprice-oRecomKeywordItem.FItemList(i).Fsailprice)/oRecomKeywordItem.FItemList(i).Forgprice*100) & "%��)" & FormatNumber(oRecomKeywordItem.FItemList(i).Fsailprice,0) & "</font>"
			end if
			'������
			if oRecomKeywordItem.FItemList(i).FitemCouponYn="Y" then
				Select Case oRecomKeywordItem.FItemList(i).FitemCouponType
					Case "1"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oRecomKeywordItem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oRecomKeywordItem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
				end Select
			end if
        %>
        </td>
        <td align="right">
        <%
            '���ΰ�
			if oRecomKeywordItem.FItemList(i).Fsailyn="Y" then
			    if (oRecomKeywordItem.FItemList(i).Fsailsuplycash>oRecomKeywordItem.FItemList(i).Forgsuplycash) then
			        Response.Write "<strong>"&FormatNumber(oRecomKeywordItem.FItemList(i).Forgsuplycash,0)&"</strong>"
			        Response.Write "<br><strong><font color=#F08050>" & FormatNumber(oRecomKeywordItem.FItemList(i).Fsailsuplycash,0) & "</font></strong>"
			    else
			        Response.Write FormatNumber(oRecomKeywordItem.FItemList(i).Forgsuplycash,0)
    				Response.Write "<br><font color=#F08050>" & FormatNumber(oRecomKeywordItem.FItemList(i).Fsailsuplycash,0) & "</font>"
    			end if
    		else
    		    Response.Write FormatNumber(oRecomKeywordItem.FItemList(i).Forgsuplycash,0)
			end if
			'������
			if oRecomKeywordItem.FItemList(i).FitemCouponYn="Y" then
				if oRecomKeywordItem.FItemList(i).FitemCouponType="1" or oRecomKeywordItem.FItemList(i).FitemCouponType="2" then
					if oRecomKeywordItem.FItemList(i).Fcouponbuyprice=0 or isNull(oRecomKeywordItem.FItemList(i).Fcouponbuyprice) then
						Response.Write "<br><font color=#5080F0>" & FormatNumber(oRecomKeywordItem.FItemList(i).Forgsuplycash,0) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & FormatNumber(oRecomKeywordItem.FItemList(i).Fcouponbuyprice,0) & "</font>"
					end if
				end if
			end if
        %>
        </td>
			
		<td><%= fnColor(oRecomKeywordItem.FItemList(i).Fmwdiv,"mw") %></td>
        <td ><%= fnColor(oRecomKeywordItem.FItemList(i).Fsellyn,"yn") %></td>
		<td><%= fnColor(oRecomKeywordItem.FItemList(i).Fisusing,"yn") %></td>

        <td>
            <% if oRecomKeywordItem.FItemList(i).Flimityn="Y" then %>
                ����(<%=oRecomKeywordItem.FItemList(i).GetLimitEa%>)
            <% end if %>
        </td>
		<td>
		    <input type="button" value="����" class="button" onClick="delItem('<%= group_no %>','<%=oRecomKeywordItem.FItemList(i).Fitemid%>')">    
		</td>
        
	</tr>
	<%
	next
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="11">
	<% if (oRecomKeywordItem.FTotalCount <1) then %>
			�˻������ �����ϴ�.
    <% else %>
        <% if oRecomKeywordItem.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oRecomKeywordItem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oRecomKeywordItem.StartScrollPage to oRecomKeywordItem.FScrollCount + oRecomKeywordItem.StartScrollPage - 1 %>
			<% if i>oRecomKeywordItem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oRecomKeywordItem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	<% end if %>
	    </td>
	</tr>
</table>

<form name="frmDel" method="post" action="keywordRecom_Process.asp">
<input type="hidden" name="mode" value="delitem">
<input type="hidden" name="group_no" value="">
<input type="hidden" name="itemid" value="">
</form>
<% 
SET oRecomKeywordItem = NOTHING
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
