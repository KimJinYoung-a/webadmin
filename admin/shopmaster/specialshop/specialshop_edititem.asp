<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ȸ����
' Hieditor : 2009.12.28 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/specialshop/specialshop_cls.asp"-->

<%
dim id , i
	id = request("id")
	
if id = "" then
	response.write "<script>alert(id���� �����ϴ�'); self.close();</script>"
end if	
	
dim ospecialshop_item
set ospecialshop_item = new cspecialshop_list
	ospecialshop_item.frectid = id
	ospecialshop_item.fspecialshop_itemlist()	
%>

<script language="javascript">

	function SaveArr(){
		if (frm.itemidarr.value==''){
			alert('��ǰ�ڵ带 �Է� �ϼ���');
			frm.itemidarr.focus();
			return;
		}
		
		frm.mode.value='itemadd';	
		frm.action='/admin/shopmaster/specialshop/specialshop_process.asp';
		frm.submit();		
	}

	function dellitem(idx){	
		frm.mode.value='dellitem';	
		frm.idx.value=idx;	
		frm.action='/admin/shopmaster/specialshop/specialshop_process.asp';
		frm.submit();		
	}

</script>

<!-- �׼� ���� -->
�ػ�ǰ���
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="post" action="">
<input type="hidden" name="mode">
<input type="hidden" name="id" value="<%=id %>">
<input type="hidden" name="idx" >	
	<tr>
		<td>
			<!-- input type=text name="itemidarr" size="30" maxlength=64 -->
			<textarea name="itemidarr" rows="5"></textarea>
			<input type=button value="��ǰ�߰�" onClick="SaveArr(frm)" class="button">
			<br>(���콺�� �ܾ �����ؼ� �ٿ������� ���� ������ ������ Ư�����ڰ� ����� �� �ֽ��ϴ�. �׷��� �������ϴ�.)
		</td>
		<td class="a" align="right">
		</td>
	</tr>
</form>	
</table>
<!-- �׼� �� -->
<br><font color="#3366FF">+ ������: BLUE 15%, VIP silver 20%, VIP gold 25%, STAFF 25%, FAMILY 20%</font>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if ospecialshop_item.ftotalcount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= ospecialshop_item.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">   		
		<td align="center">��ǰ�ڵ�</td>
		<td align="center">�̹���</td>				
		<td align="center">��ǰ��</td>
		<td align="center">ǰ������</td>
		<td align="center">�ǸŰ�<br/><font color="#3366FF">�ִ����ΰ�(25%)</font></td>
		<td align="center">���ް�</td>
		<td align="center">����</td>
		<td align="center">���</td>	
    </tr>
	<% for i=0 to ospecialshop_item.ftotalcount-1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			
	
    <% if ospecialshop_item.FItemList(i).fstatus = "3" then %>    
    <tr align="center" bgcolor="#FFFFaa">
    <% else %>    
    <tr align="center" bgcolor="#FFFFFF">
	<% end if %>	
		<td align="center">
			<a href="<%=wwwUrl%>/shopping/category_prd.asp?itemid=<%=ospecialshop_item.FItemList(i).fitemid%>" onfocus="this.blur()" target="_blink"><%=ospecialshop_item.FItemList(i).fitemid%></a>
		</td>	
		<td align="center">
			<img src="<%=ospecialshop_item.FItemList(i).FImageSmall%>" width=50 height=50>
		</td>		
		<td align="center">
			<%=ospecialshop_item.FItemList(i).fitemname%>
		</td>		
		<td align="center">
			<%
			if ospecialshop_item.FItemList(i).fsellyn <> "Y" then 
				response.write "ǰ��"
			else
				response.write "�Ǹ���"
			end if				 
				
			%>
		</td>
		
		<td class="verdana-small">		 
			<% if ospecialshop_item.FItemList(i).IsSail then %><font color="#F08050"><% end if %><%= FormatNumber(ospecialshop_item.FItemList(i).FSellCash,0) %>��</font>
			 <br>  <font color="#3366FF"><%= FormatNumber(ospecialshop_item.FItemList(i).getRealPrice ,0) %>��</font>
			 
		</td>
		<td  class="verdana-small">	
			<%=FormatNumber(ospecialshop_item.FItemList(i).FBuyCash,0)%>��
		</td>
		<td class="verdana-small"><%IF ospecialshop_item.FItemList(i).getMargin < 0 then%><font color="red"><%end if%><%=FormatNumber(ospecialshop_item.FItemList(i).getMargin,1)%>%</td>
		<td align="center">
			<input type="button" class="button" value="����" onclick="dellitem(<%=ospecialshop_item.FItemList(i).fidx%>);">
		</td>
    </tr>   
	</form>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[��ϵǾ� �ִ� ��ǰ�� �����ϴ�.]</td>
		</tr>
	<% end if %>

</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->