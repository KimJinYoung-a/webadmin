<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ��󿬶���
' Hieditor : 2011.01.14 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/shopmaster/shopuser_cls.asp"-->

<%
dim omember , i, shopid

'/����
if (C_IS_SHOP) then
	
	'//�������϶�
	if C_IS_OWN_SHOP then
		
		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if		
	else
		shopid = C_STREETSHOPID
	end if
else
	'/��ü
	if (C_IS_Maker_Upche) then
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if

set omember = new cshopuser_list
	omember.getofflinemember_list()
	
%>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get action="">
<input type="hidden" name="editor_no">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
	
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<br>		
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		�� �̸��� ���콺�� �������� ������ ��Ÿ���ϴ�.
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= omember.FTotalCount %></b> �� �� 500�Ǳ��� �˻� �˴ϴ�.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��å</td>
	<td>�̸�</td>
	<td>�μ�</td>
	<td>�ڵ�����ȣ</td>	
	<td>ȸ����ȭ</td>
	<td>����</td>
	<td>�����ȣ(070)</td>
	<td>�̸���</td>
	<td>MSN�޽���</td>
</tr>
<% if omember.ftotalcount > 0 then %>

<% for i=0 to omember.FTotalCount - 1 %>
<tr height=30 align="center" bgcolor="<% if omember.FItemList(i).fstatediv="Y" then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
	<td><%= omember.FItemList(i).fjob_name %></td>
	<td>
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td id="photo<%=i%>" alt="<img src='<%= omember.FItemList(i).fuserimage %>' width='110'>"><%= omember.FItemList(i).fusername %>(<%= omember.FItemList(i).fposit_name %>)</td>
		</tr>
		</table>
		<div id="ddd0" style="background-color:white; border-width:1px; border-style:solid; width:110; position:absolute; left:10; top:10; z-index:1; display:none"></div>
	</td>
	<td>
		<%= omember.FItemList(i).fpart_name %><Br>
		<% if omember.FItemList(i).fpart_sn = "16" or omember.FItemList(i).fpart_sn = "18" or omember.FItemList(i).fpart_sn = "19" then %>
			<font color="grey">	
			<% if omember.FItemList(i).fpart_name <> "" then %>
				<%= omember.FItemList(i).fshopfirst %>/<%= omember.FItemList(i).fshopname %> (������: <%= omember.FItemList(i).fshopcount %>��)
			<% else %>
				��������
			<% end if %>
		<% end if %>
	</td>
	<td><%= omember.FItemList(i).fusercell %></td>	
	<td><%= omember.FItemList(i).finterphoneno %></td>
	<td><%= omember.FItemList(i).fextension %></td>
	<td><%= omember.FItemList(i).fdirect070 %></td>
	<td><%= omember.FItemList(i).fusermail %></td>
	<td><%= omember.FItemList(i).fmsnmail %></td>
</tr>
<% next %>

<% else %>

<tr>
	<td colspan="15" align="center" bgcolor="#FFFFFF">���(�˻�)�� ����ڰ� �����ϴ�.</td>
</tr>

<% end if %>
</table>

<script language="javascript">

document.onmousemove=function(){ 
	oElement = document.elementFromPoint(event.x, event.y);
	var ddd0 = document.getElementById("ddd0");
	if(oElement.id.indexOf('photo')!=-1)
	{
		ddd0.style.display='';
		ddd0.style.pixelLeft=event.x+10 + document.body.scrollLeft;
		ddd0.style.pixelTop=event.y-80 + document.body.scrollTop;
		ddd0.innerHTML=oElement.alt;
	} else { 
		ddd0.style.display='none';
	}
}

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->