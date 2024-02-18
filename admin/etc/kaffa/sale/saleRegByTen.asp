<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/kaffa/itemsalecls.asp"-->
<%
Dim i
Dim regyn : regyn= requestCheckvar(request("regyn"),10)
Dim page  : page= requestCheckvar(request("page"),10)
Dim clsSale

if (page="") then page=1

Set clsSale = new CSale
clsSale.FCurrPage	= page
clsSale.FPageSize = 20
clsSale.FRectTenCodePreReg = regyn
clsSale.getTenSaleListWithKaffa
%>

<script language='javascript'>
function regByTenSale(tensalecode){
    if (confirm('TEN �����ڵ�'+tensalecode+'�� ����Ͻðڽ��ϱ�?\n\n���ϱⰣ�� ��ϵ� ��ǰ�� ���ܵǸ�, �߱�����Ʈ ������ǰ�� ��ϵ˴ϴ�.')){
        document.frmSubmit.tensalecode.value=tensalecode;
        document.frmSubmit.mode.value="T";
        document.frmSubmit.submit();
    }
}
function goPage(p){
    document.frmSearch.page.value = p;
    document.frmSearch.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frmSearch" method="get"  >
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
		<td align="left">
		��Ͽ��� :
		<select name="regyn">
            <option value="">��ü
            <option value="Y" <%=CHKIIF(regyn="Y","selected","") %> >����
            <option value="N" <%=CHKIIF(regyn="N","selected","") %> >�̵��
		</select>
		</td>
		<td  width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frmSearch.submit();">
		</td>
	</tr>
	</form>
</table>
<!---- /�˻� ---->
<p>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="12">�˻���� : <b><%= FormatNumber(clsSale.FTotalCount,0) %></b>&nbsp;&nbsp;������ : <b><%= FormatNumber(page,0) %> / <%= FormatNumber(clsSale.FTotalPage,0) %></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>TEN�����ڵ�</td>
    	<td>TEN���θ�</td>
    	<td>TEN������</td>
    	<td>TEN���԰�����</td>
    	<td>TEN������</td>
    	<td>TEN������</td>
    	<td>TEN����</td>
    	<td>TEN�����</td>
    	<td>���ɼ���</td>
    	<td>ó��</td>
    </tr>
    <% For i = 0 To clsSale.FResultCount - 1 %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%=clsSale.FItemList(i).FTENsale_code %></td>
    	<td><%=clsSale.FItemList(i).FTENsale_name %></td>
    	<td><%=clsSale.FItemList(i).FTENsale_rate %>%</td>
    	<td><%=clsSale.FItemList(i).getTenSaleMarginGubun %></td>
    	<td><%=clsSale.FItemList(i).FTENsale_startdate %></td>
    	<td><%=clsSale.FItemList(i).FTENsale_enddate %></td>
    	<td><%=clsSale.FItemList(i).getTenSaleStateName %></td>
    	<td><%=clsSale.FItemList(i).FTENregdate %></td>
    	<td><%=clsSale.FItemList(i).FvalidCnt %></td>
    	<td>
    	<% if isNULL(clsSale.FItemList(i).FDiscountKey)  then %>
    	    <input type="button" value="���" onClick="regByTenSale('<%=clsSale.FItemList(i).FTENsale_code %>')">
    	<% else %>
    	    <%=clsSale.FItemList(i).FDiscountKey%>
    	<% end if %>
    	</td>
    </tr>
    <% next %>
    <tr height="20">
		<td colspan="11" align="center" bgcolor="#FFFFFF">
		<% If clsSale.HasPreScroll Then %>
			<a href="javascript:goPage('<%= clsSale.StartScrollPage-1 %>');">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i=0 + clsSale.StartScrollPage To clsSale.FScrollCount + clsSale.StartScrollPage - 1 %>
			<% If i>clsSale.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If clsSale.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>');">[next]</a>
		<% Else %>
		[next]
		<% End If %>
		</td>
	</tr>
</table>

<%
Set clsSale = Nothing
%>
<form name="frmSubmit" method="post" action="saleitemProc.asp">
<input type="hidden" name="mode" value="T">
<input type="hidden" name="tensalecode" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->