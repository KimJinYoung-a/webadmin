<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� ������
' Hieditor : 2012.03.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/shopcscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/offshop/order/shopcscenter_order_cls.asp"-->
<!-- #include virtual="/admin/offshop/shopcscenter/cscenter_Function_off.asp"-->
<%
dim masteridx ,ojumun , oaslist, totalascount ,ix , orderno ,shopid , oaslistmaejang , oaslistfinal
dim maejangascount , finalascount
	masteridx = RequestCheckVar(request("masteridx"),16)

totalascount = 0

set ojumun = new COrder
	if masteridx <> "" then
	    ojumun.FRectmasteridx = masteridx
	    ojumun.fQuickSearchOrderMaster
	end if

if ojumun.ftotalcount > 0 then
	orderno = ojumun.FOneItem.Forderno
	shopid = ojumun.FOneItem.Fshopid
end if

'/����a/s �Ǽ�
set oaslist = new COrder
	if masteridx <> "" then
	    oaslist.FRectmasteridx = masteridx
	    oaslist.fGetCSASTotalCount
		
	    totalascount = oaslist.FResultCount
	end if

'/����ó�� ���Ǽ�
set oaslistmaejang = new COrder
	if masteridx <> "" then
	    oaslistmaejang.FRectmasteridx = masteridx
	    oaslistmaejang.frectcurrstate = "'B001','B004'"
	    oaslistmaejang.frectdeleteyn = "N"
	    oaslistmaejang.fGetCSASTotalCount
		
	    maejangascount = oaslistmaejang.FResultCount
	end if

'/�����Ϸ�ó�� ���Ǽ�
set oaslistfinal = new COrder
	if masteridx <> "" then
	    oaslistfinal.FRectmasteridx = masteridx
	    oaslistfinal.frectcurrstate = "'B006','B008'"
	    oaslistfinal.frectdeleteyn = "N"
	    oaslistfinal.fGetCSASTotalCount
		
	    finalascount = oaslistfinal.FResultCount
	end if	 
	 

%>

<script language="javascript">
</script>

<% if (masteridx<>"") then %>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="FFFFFF">
<tr height="25">
	<td align="left">
		<input type="button" class="button" value="��üA/S��û" class="csbutton" onclick="javascript:PopOpenServiceItemas('<%= masteridx %>');">
		<input type="button" class="button" value="����Ϸ�ó��[<%=maejangascount%>]" class="csbutton" onclick="javascript:PopmaejangAction('<%=orderno%>','<%= shopid %>','','notfinish');">
		<input type="button" class="button" value="�����Ϸ�ó��/����[<%= finalascount %>]" class="csbutton" onclick="javascript:Cscenter_Action_List_off('<%= masteridx %>','<%=orderno%>','','notfinish','<%= shopid %>');">
    </td>
    <td align="right">
    	<input type="button" class="button" value="�����������" style="width:90px;" onclick="javascript:popOrderReceipt('<%= orderno %>');">
	</td>
</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr valign="top">
	<td>
		<!-- ���Ż�ǰ���� -->
		<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25" bgcolor="<%= adminColor("topbar") %>" style="padding:2 2 2 2">
		    <td colspan="10">
		    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		    		<tr>
		    			<td width="500">
		    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���Ż�ǰ����</b>
					    	[<b><%= orderno %> �ֹ� ���� ��A/S <%=totalascount%>��</b>]
    				    </td>
    				    <td align="right">
    				    </td>
    				</tr>
    			</table>
    		</td>
		</tr>
		<tr height="400" bgcolor="#FFFFFF">
		    <td valign="top">
		        <table height="25" width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#BABABA">
		            <tr align="center" bgcolor="<%= adminColor("topbar") %>" style="padding:2">
                    	<td width="30">����</td>
                    	<td width="80">CODE</td>                      	
                        <td width="120">�귣��ID</td>
                    	<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
                    	<td width="30">����</td>
                    	<td width="50">����<br>�Һ��ڰ�</td>
                    	<td width="50">�ǸŰ�</td>
                    </tr>
                    <tr>
                        <td height="1" colspan="15" bgcolor="#BABABA"></td>
                    </tr>
                 </table>
                 <table height="365" width="100%" border=0 cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
                    <tr height="100%">
                        <td colspan="12">
                	        <iframe name="orderdetail" src="/admin/offshop/shopcscenter/order/orderitemmaster.asp?masteridx=<%= masteridx %>" border=0 frameSpacing=0 frameborder="no" width="100%" height="100%" leftmargin="0"></iframe>
                        </td>
                    <tr>
                </table>
		    </td>
		</tr>
		</table>
		<!-- ���Ż�ǰ���� -->
	</td>
	<td width="5"></td>
	<td width="250" align="right">
		<!-- ���������� -->
		<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<form name="frmbuyerinfo" onsubmit="return false;">
		<tr height="25" bgcolor="<%= adminColor("topbar") %>">
		    <td colspan="2">
		    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		    		<tr>
		    			<td width="100">
		    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>������ ����</b>
    				    </td>
    				    <td align="right">    				    	
    				    </td>
    				</tr>
    			</table>
    		</td>
		</tr>
		<tr height="23">
		    <td bgcolor="<%= adminColor("topbar") %>">IDX</td>
		    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.fmasteridx %></td>
		</tr>
		<tr height="23">
		    <td bgcolor="<%= adminColor("topbar") %>">�ֹ���ȣ</td>
		    <td bgcolor="#FFFFFF"><%= orderno %></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">�����ڸ�</td>
		    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FBuyName %></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">��ȭ��ȣ</td>
		    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FBuyPhone %></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">�ڵ���</td>
		    <td bgcolor="#FFFFFF">
		        <%= ojumun.FOneItem.FBuyHp %>
		        <input type="button" name="buyhp" class="button" value="SMS" onclick="javascript:PopCSSMSSend_off('<%= ojumun.FOneItem.FBuyHp %>','<%= ojumun.FOneItem.Fmasteridx %>','<%= ojumun.FOneItem.forderno %>','');">
		    </td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">�̸���</td>
		    <td bgcolor="#FFFFFF">
		        <%= ojumun.FOneItem.FBuyEmail %>
		    </td>
		</tr>
		</form>
		</table>
		<!-- ���������� -->
		<Br>
	    <!-- �ֹ����� -->
		<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">�ֹ�����</td>
		    <td bgcolor="#FFFFFF">				        
		        <font color="<%= ojumun.FOneItem.CancelYnColor %>"><%= ojumun.FOneItem.CancelYnName %></font>
		    </td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">�ֹ��Ͻ�</td>
		    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FRegDate %></td>
		</tr>
		<!-- �ֹ����� -->		
		</table>	
	</td>
</tr>
</table>

<br>

<% else %>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr height="50">
    <td align="center"> [ �󼼳����� ���÷��� �ֹ���ȣ�� ���� �ϼ��� ]</td>
</tr>
</table>
<% end if %>

<%
set ojumun = Nothing
set oaslist = Nothing
set oaslistmaejang = Nothing
set oaslistfinal = Nothing
%>
<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->