<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �𺰱�������
' Hieditor : 2010.12.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone/zone_cls.asp"-->
<%
Dim ozone,idx , i , shopid,zonename,racktype,unit,orderno,regdate ,isusing , zonegroup
dim menupos
	idx = requestCheckVar(request("idx"),10)
	menupos = requestCheckVar(request("menupos"),10)

set ozone = new czone_list
	ozone.frectidx = idx
	
	'//�����ÿ��� ����
	if idx <> "" then		
		ozone.fzone_oneitem()
		
		if ozone.ftotalcount >0 then			
			shopid = ozone.FOneItem.fshopid
			zonegroup = ozone.FOneItem.fzonegroup
			racktype = ozone.FOneItem.fracktype			
			zonename = ozone.FOneItem.fzonename
			unit = ozone.FOneItem.funit			
			regdate = ozone.FOneItem.fregdate
			isusing = ozone.FOneItem.fisusing						
		end if
	end if
	
%>

<script language="javascript">
	
	function reg(){
		if (frm.shopid.value=='') {
			alert('���� ������ �ּ���');
			frm.zonename.focus();
			return;
		}

		if (frm.zonegroup.value=='') {
			alert('�׷� ������ �ּ���');
			frm.zonegroup.focus();			
			return;
		}

		if (frm.racktype.value=='') {
			alert('�Ŵ� Ÿ���� ������ �ּ���');
			frm.racktype.focus();			
			return;
		}
		
		if (frm.zonename.value=='') {
			alert('�������� �Է��� �ּ���');
			frm.zonename.focus();
			return;
		}
		
		if (frm.unit.value=='') {
			alert('�ش������� ����� �Է��� �ּ���');
			frm.unit.focus();			
			return;
		}
		
		if(frm.unit.value!=''){
			if (!IsDouble(frm.unit.value)){
				alert('�ش������� ����� ���ڸ� �����մϴ�.');
				frm.unit.focus();
				return;
			}
		}	

		if (frm.isusing.value=='') {
			alert('��뿩�θ� ������ �ּ���');
			frm.isusing.focus();			
			return;
		}
		
		frm.action='zone_process.asp';
		frm.mode.value = "zonereg";
		frm.submit();
	}
	
</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		�� <font color="red">[�߿�] </font>���峻 ������ ����ǰų� ��������,
		<br>���� ������ �������� ���ð�, ������ ��������, ���� ����ϼ���.
		<br>���� ������ ���� ����� �������� ������ ��� �Ͻǰ��,
		<br>���� �������� ��ϵǾ��� ��ǰ���� ��� ���� �������� ����Ǵ� ������ �߻��˴ϴ�
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr bgcolor="#FFFFFF">
	<td align="center">��ȣ<br></td>
	<td>
		<%=idx%><input type="hidden" name="idx" value="<%=idx%>">
	</td>
</tr>	
<tr bgcolor="#FFFFFF">
	<td align="center">SHOP</td>
	<td>
		<% drawSelectBoxOffShop "shopid",shopid %>
	</td>
</tr>
	
<tr bgcolor="#FFFFFF">
	<td align="center">�׷�</td>
	<td>
		<% drawSelectBoxOffShopzonegroup "zonegroup",zonegroup,"" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">�Ŵ�Ÿ��</td>
	<td>
		<% drawSelectBoxOffShopracktype "racktype",racktype,"" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">�󼼱���</td>
	<td>
		<input type="text" name="zonename" value="<%=zonename%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">UNIT</td>
	<td>
		<input type="text" name="unit" value="<%=unit%>" size=5 maxlength=5> ex)1
		<Br>�� �ش������� ����� ����Ͻðų�, ���������� ���ϽŴ�� �����ؼ� ����Ͻø� �˴ϴ�
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">��뿩��<br></td>
	<td>
		<select name="isusing">
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan=2>
		<input type="button" value="����" class="button" onclick="reg();">
	</td>
</tr>
</form>
</table>	

<% set ozone = nothing %>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
