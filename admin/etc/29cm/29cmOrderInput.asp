<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbagirlopen.asp" -->
<!-- #include virtual="/lib/db/dbagirlHelper.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/aGirlOrderCls.asp"-->

<%
Dim page : page = requestCheckVar(page,10)
Dim OrderStatus : OrderStatus = requestCheckVar(request("OrderStatus"),10)
Dim research    : research = requestCheckVar(request("research"),10)

if (page="") then page=1
if research="" and OrderStatus="" then OrderStatus="N"
 
Dim oGirlOrder
set oGirlOrder = new aGirlOrder
if (OrderStatus="N") then
    oGirlOrder.getAgirlNotRegOrderList(0)
elseif (OrderStatus="Y") then
    oGirlOrder.getAgirlNotRegOrderList(3)
elseif (OrderStatus="7") then
    oGirlOrder.getAgirlNotRegOrderList(7)
else
    oGirlOrder.getAgirlNotRegOrderList("")
end if

dim i
dim pOrderSerial
%>
<script language='javascript'>
function fnCheckValidAll(bool, comp){
    var frm = comp.form;

    if (!comp.length){
        if (comp.disabled==false){
            comp.checked = bool;
            AnCheckClick(comp);
        }
    }else{
        for (var i=0;i<comp.length;i++){
            if (comp[i].disabled==false){
                comp[i].checked = bool;
                AnCheckClick(comp[i]);
            }
        }
    }
}

function OrderInput(frm){
    var checkedExists = false;
    if (!frm.cksel.length){
        if (frm.cksel.checked){
            checkedExists = true;
        }
    }else{
        
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                checkedExists = true;
                break;
            }
        }
    }
    
    if (!checkedExists){
        alert('���� �ֹ��� �����ϴ�.');
        return;
    }
    
    if (confirm('�ֹ��� �Է� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr >
		<td class="a">
    		�ֹ��� : 
    		
    		�ֹ����� :
    		<select name="OrderStatus">
    		
    		<option value="">��ü
    		<option value="N" <%= CHkIIF(OrderStatus="N","selected","") %> >��Ȯ��
    		<option value="Y" <%= CHkIIF(OrderStatus="Y","selected","") %> >Ȯ��
    		<option value="7" <%= CHkIIF(OrderStatus="7","selected","") %> >���Ϸ�
    		</select>
    		
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="left"><input type="button" value="���ó��� �ֹ��Է�" onClick="OrderInput(frmSvArr);"></td>
	<td colspan="12" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oGirlOrder.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(oGirlOrder.FTotalCount,0) %></td>
</tr>
<form name="frmSvArr" method="post" action="29cmOrderInput_Process.asp">
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckValidAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">���ֹ���ȣ</td>
	<td width="50">����ǰ�ڵ�</td>
	<td width="50">���ɼ��ڵ�</td>
	<td width="150">����ǰ�� <font color="blue">[�ɼ�]</font></td>
	<td width="50">������</td>
	<td width="50">������</td>
	<td width="50">����</td>
	<td width="50">�ǸŰ�</td>
	<td width="50">�ǰ�����<br>(Total)</td>
	<td width="50">TenItemID</td>
	<td width="50">TenItemOption</td>
	<td width="50">Status</td>
</tr>
<% for i=0 to oGirlOrder.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="20">	
    <td align="center">
        <% if  pOrderSerial=oGirlOrder.FItemList(i).FOrderserial then %>
        =
        <% else %>
            <% if IsNULL(oGirlOrder.FItemList(i).FpartnerItemID) or IsNULL(oGirlOrder.FItemList(i).FpartnerOption) then %>
            <input type="checkbox" name="cksel" Disabled onClick="AnCheckClick(this);"  value="<%= oGirlOrder.FItemList(i).FOrderserial %>">
            <% else %>
            <input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oGirlOrder.FItemList(i).FOrderserial %>">
            <% end if %>
        <% end if %>
    <% pOrderSerial= oGirlOrder.FItemList(i).FOrderserial %>
    </td>
    <td align="center"><%= oGirlOrder.FItemList(i).FOrderserial %></td>
    <td align="center"><%= oGirlOrder.FItemList(i).FItemSeq %></td>
    <td align="center"><%= oGirlOrder.FItemList(i).FOptionCode %></td>
    <td align="center"><%= oGirlOrder.FItemList(i).FItemName %>
    <% if oGirlOrder.FItemList(i).FOptionValue<>"" then %>
    <font color="blue">[<%= oGirlOrder.FItemList(i).FOptionValue %>]</font>
    <% end if %>
    </td>
    <td align="center"><%= oGirlOrder.FItemList(i).FOrderName %></td>
    <td align="center"><%= oGirlOrder.FItemList(i).FReceiveName %></td>
    <td align="center"><%= oGirlOrder.FItemList(i).FOrderCount %></td>
    <td align="center"><%= FormatNumber(oGirlOrder.FItemList(i).FRealSellPrice,0) %></td>
    <td align="center"><%= FormatNumber(oGirlOrder.FItemList(i).FPayRealPrice,0) %></td>
    <td align="center"><%= oGirlOrder.FItemList(i).FpartnerItemID %></td>
    <td align="center"><%= oGirlOrder.FItemList(i).FpartnerOption  %></td>
    
    <td align="center"><%= oGirlOrder.FItemList(i).FOrderItemStatus  %></td>
    
    
</tr>
<% next %>
</form>
<tr height="20">
    <td colspan="14" align="center" bgcolor="#FFFFFF">
    <!--
        <% if oGirlOrder.HasPreScroll then %>
		<a href="javascript:goPage('<%= oGirlOrder.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>
    
    	<% for i=0 + oGirlOrder.StartScrollPage to oGirlOrder.FScrollCount + oGirlOrder.StartScrollPage - 1 %>
    		<% if i>oGirlOrder.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>
    
    	<% if oGirlOrder.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    -->
    </td>
</tr>
</table>

<%
set oGirlOrder = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbagirlclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->