<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ������
' Hieditor : 2011.03.08 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim masteridx ,ojumun , oaslist, totalascount ,ix
	masteridx = RequestCheckVar(request("masteridx"),16)

totalascount = 0

set ojumun = new COrder
	if masteridx <> "" then
	    ojumun.FRectmasteridx = masteridx
	    ojumun.fQuickSearchOrderMaster
	end if
	
set oaslist = new COrder
	if masteridx <> "" then
	    oaslist.FRectmasteridx = masteridx
	    oaslist.fGetCSASTotalCount
	
	    totalascount = oaslist.FResultCount
	end if
%>

<script language="javascript">
</script>

<% if (masteridx<>"") then %>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="FFFFFF">
<tr height="25">
	<td align="left">
		<input type="button" class="button" value="���" class="csbutton" style="width:60px;" onclick="javascript:PopOpenCancelItem('<%= masteridx %>');">
		<input type="button" class="button" value="�±�ȯ" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemChange('<%= masteridx %>');">
		<input type="button" class="button" value="������߼�" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemOmit('<%= masteridx %>');">	        
		<!--<input type="button" class="button" value="���񽺹߼�" class="csbutton" style="width:70px;" onclick="javascript:PopOpenServiceItemMore('<%'= masteridx %>');">-->
		<input type="button" class="button" value="�������ǻ���" class="csbutton" style="width:90px;" onclick="javascript:PopOpenReadMe('<%= masteridx %>','');">
    </td>
    <td align="right">
		<input type="button" class="button" value="�����������" style="width:90px;" onclick="javascript:popOrderReceipt('<%= masteridx %>');">
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
					    	&nbsp;
					    	[<b><%= masteridx %></b>]
					    	&nbsp;
					    	<input type="button" class="button" value="����CS <%= totalascount %>��" class="csbutton" style="width:90px;" onclick="javascript:Cscenter_Action_List_off('<%= masteridx %>','');">
    				    </td>
    				    <td align="right">
    				    	<input type="button" class="button" value="������ǰ����" class="csbutton" style="width:90px;" onclick="misendmaster('<%= masteridx %>');">
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
                    	<td width="50">�������</td>
                    	<td width="80">CODE</td>                      	
                        <td width="120">�귣��ID</td>
                    	<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
                    	<td width="30">����</td>
                    	<td width="50">����<br>�Һ��ڰ�</td>
                    	<td width="50">�ǸŰ�</td>
                    	<td width="70">Ȯ����</td>
                    	<td width="70">�����</td>
                    	<td width="125">�������</td>
                    </tr>
                    <tr>
                        <td height="1" colspan="15" bgcolor="#BABABA"></td>
                    </tr>
                 </table>
                 <table height="365" width="100%" border=0 cellspacing=0 cellpadding=0 class=a bgcolor="FFFFFF">
                    <tr height="100%">
                        <td colspan="12">
                	        <iframe name="orderdetail" src="/admin/offshop/cscenter/order/orderitemmaster.asp?masteridx=<%= masteridx %>" border=0 frameSpacing=0 frameborder="no" width="100%" height="100%" leftmargin="0"></iframe>
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
    				    	<input type="button" class="button" value="��������������" class="csbutton" onclick="javascript:PopBuyerInfo_off('<%= masteridx %>');">
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
		    <td bgcolor="#FFFFFF"><%= masteridx %></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">�����ڸ�</td>
		    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="buyname" value="<%= ojumun.FOneItem.FBuyName %>" size="8" readonly></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">��ȭ��ȣ</td>
		    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="buyphone" value="<%= ojumun.FOneItem.FBuyPhone %>" readonly></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">�ڵ���</td>
		    <td bgcolor="#FFFFFF">
		        <input type="text" class="text_ro" name="buyhp" value="<%= ojumun.FOneItem.FBuyHp %>" readonly>
		        <input type="button" name="buyhp" class="button" value="SMS" onclick="javascript:PopCSSMSSend_off('<%= ojumun.FOneItem.FBuyHp %>','<%= ojumun.FOneItem.Fmasteridx %>','');">
		    </td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">�̸���</td>
		    <td bgcolor="#FFFFFF">
		        <input type="text" class="text_ro" name="buyemail" value="<%= ojumun.FOneItem.FBuyEmail %>" size="20" readonly>
		        <input type="button" name="email" class="button" value="mail" onclick="javascript:PopCSMailSend_off('<%= ojumun.FOneItem.FBuyEmail %>','<%= ojumun.FOneItem.Fmasteridx %>');">
		    </td>
		</tr>
		</form>
		</table>
		<!-- ���������� -->
        <p>
		<!-- ������� -->
		<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<form name="frmreqinfo" onsubmit="return false;">
		<tr height="25" bgcolor="<%= adminColor("topbar") %>">
		    <td colspan="2">
		    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		    		<tr>
		    			<td width="100">
		    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��� ����</b>
    				    </td>
    				    <td align="right">
    				    	<input type="button" class="button" value="�������������" class="csbutton" onclick="javascript:PopReceiverInfo_off('<%= masteridx %>');">
    				    </td>
    				</tr>
    			</table>
    		</td>
		</tr>
		<tr>
		    <td width="100" bgcolor="<%= adminColor("topbar") %>">�����θ�</td>
		    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" value="<%= ojumun.FOneItem.FReqName %>" readonly></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">��ȭ��ȣ</td>
		    <td bgcolor="#FFFFFF"><input type="text" class="text_ro" value="<%= ojumun.FOneItem.FReqPhone %>" readonly></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">�ڵ���</td>
		    <td bgcolor="#FFFFFF">
		        <input type="text" class="text_ro" value="<%= ojumun.FOneItem.FReqHp %>" readonly>
		        <input type="button" name="reqhp" class="button" value="SMS" onclick="javascript:PopCSSMSSend_off('<%= ojumun.FOneItem.FReqHp %>','<%= ojumun.FOneItem.Fmasteridx %>','');">
		    </td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">����ּ�</td>
		    <td bgcolor="#FFFFFF">
		        <input type="text" class="text_ro" name="txzip1" value="<%= ojumun.FOneItem.FReqZipCode %>" size="7" readonly>
		        <input type="text" class="text_ro" value="<%= ojumun.FOneItem.FReqZipAddr %>" size="18" readonly><br>
		        <textarea class="textarea_ro" rows="2" cols="28" readonly><%= ojumun.FOneItem.FReqAddress %></textarea>
            </td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">��Ÿ����</td>
		    <td bgcolor="#FFFFFF">
		        <textarea class="textarea_ro" rows="2" cols="28" readonly><%= ojumun.FOneItem.FComment %></textarea>
		    </td>
		</tr>
		</form>		
		</table>
		<!-- ������� -->
		<Br>
	    <!-- �ֹ����� -->
		<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">�ֹ�����</td>
		    <td bgcolor="#FFFFFF">				        
		        [<font color="<%= ojumun.FOneItem.shopIpkumDivColor %>"><%= ojumun.FOneItem.shopIpkumDivName %></font>]
		        <% if ojumun.FOneItem.FCancelYn<>"N" then %>
		        <font color="<%= ojumun.FOneItem.CancelYnColor %>"><%= ojumun.FOneItem.CancelYnName %></font>
		        <% end if %>
		    </td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">�ֹ��Ͻ�</td>
		    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.FRegDate %></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("topbar") %>">�ֹ��뺸</td>
		    <td bgcolor="#FFFFFF"><%= ojumun.FOneItem.Fbaljudate %></td>
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
%>
<!-- #include virtual="/admin/offshop/cscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->