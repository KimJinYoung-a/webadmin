<%@  language="VBScript" %>
<% option explicit %> 
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/tenmember/lib/header.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<% 
Dim oipchul,iscmlinkno   
iscmlinkno		=  requestCheckvar(Request("iSL"),10)
set oipchul = new CIpChulStorage
oipchul.FRectId = iscmlinkno
oipchul.GetIpChulMaster  

	  function sGetDivCodeName(Fdivcode)
		if Fdivcode="002" then
			sGetDivCodeName = "��Ź"
		elseif Fdivcode="001" then
			sGetDivCodeName = "����"
		elseif Fdivcode="003" then
			sGetDivCodeName = "����"
		elseif Fdivcode="004" then
			sGetDivCodeName = "�ܺ�"
		elseif Fdivcode="005" then
			sGetDivCodeName = "����"
		elseif Fdivcode="006" then
			sGetDivCodeName = "B2B"
		elseif Fdivcode="007" then
			sGetDivCodeName = "��Ÿ"
		elseif Fdivcode="101" then
			sGetDivCodeName = "��Ź���"
		elseif Fdivcode="801" then
			sGetDivCodeName = "Off����"
		elseif Fdivcode="802" then
			sGetDivCodeName = "Off��Ź"
		elseif Fdivcode="999" then
			sGetDivCodeName = "��Ÿ(�������)"
		end if
	end function
 %> 
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<!--���ڰ���-->
	<form name="frmEapp" method="post" action="/admin/approval/eapp/regeapp.asp">
	<input type="hidden" name="tC" value="">
	<input type="hidden" name="ieidx" value="58"> <!-- ������ȣ ����!!-->
	<input type="hidden" name="iSL" value="<%=iscmlinkno%>">
	<input type="hidden" name="mRP" value="<%=formatnumber(oipchul.FOneItem.Ftotalbuycash*-1,0)%>">
	</form>
	<div id="divEapp" style="display:none;">
	<p>&nbsp;������ ���� ��Ÿ��� �����ϰ��� �Ͽ��� ���� �� �簡 �ٶ��ϴ�. </p>
	<p>&nbsp;</p>
	<p align="center">-  ��  ��  - </p>
	<p>&nbsp;</p>
	<p>1. ��Ÿ������: </p>
	<p>&nbsp;</p>
	<p>2. ��Ÿ�����: </p>
	<p>&nbsp;</p>
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
			<td>����ڵ�</td>
			<td>���óID</td>
			<td>���ó��</td>
			<td>�����</td>
			<td>��û��</td>
			<td>�ǸŰ�</td>
			<td>���</td>
			<td>���԰�</td>
			<td>����</td>
			<td>������</td>
		</tr>
		<tr bgcolor="#FFFFFF"  align="center">
			<td><%=oipchul.FOneItem.Fcode %></td>
			<td><%=oipchul.FOneItem.Fsocid%></td>
			<td><%= oipchul.FOneItem.Fsocname%></td>
			<td><%= oipchul.FOneItem.Fchargeid %>&nbsp;(<%= oipchul.FOneItem.Fchargename %>)</td>
			<td><%= oipchul.FOneItem.Fscheduledt %></td>
			<td><%= FormatNumber(oipchul.FOneItem.Ftotalsellcash*-1,0) %></td>
			<td><%= FormatNumber(oipchul.FOneItem.Ftotalsuplycash*-1,0) %></td>
			<td><%= FormatNumber(oipchul.FOneItem.Ftotalbuycash*-1,0) %></td>
			<td><%=sGetDivCodeName(oipchul.FOneItem.Fdivcode) %></td>
			<td><% if oipchul.FOneItem.Ftotalsellcash<>0 then %>
				  <%= 100-CLng(oipchul.FOneItem.Ftotalsuplycash/oipchul.FOneItem.Ftotalsellcash*100*100)/100 %>%
				<% end if %>
			</td>
		</tr>
	</table>
	<%if oipchul.FOneItem.Fprizecnt > 0 then%>
	<p>&nbsp;</p>
	<p>3. ��÷������: </p>
	<p style="color:blue">- �̺�Ʈ�ҵ漼���� �繫�� ��������</p>
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
			<td>No</td>
			<td>��÷��ID</td>
			<td>��÷�ڸ�</td>
		</tr>
		<tr bgcolor="#FFFFFF"  align="center">
			<td>1</td>
			<td></td>
			<td></td>
		</tr>
		<tr bgcolor="#FFFFFF"  align="center">
			<td>2</td>
			<td></td>
			<td></td>
		</tr>
		<tr bgcolor="#FFFFFF"  align="center">
			<td>3</td>
			<td></td>
			<td></td>
		</tr>
	</table>
	<%end if%>
	<br /><br />
	<p>
		<b>* ���ó ID</b><br />
		<table border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
			<td>���óID</td>
			<td>����</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>etcout</td>
			<td>����̵�</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemgift</td>
			<td>��÷����ǰ</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemgift_all</td>
			<td>���Ż���ǰ</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemsample</td>
			<td>���û�� (ex.�Կ�����)</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemAD</td>
			<td>�������� (ex.����)</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemgift_Biz</td>
			<td>����� (ex.�ŷ�ó����)</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemstaff</td>
			<td>�����Ļ��� (ex.����ŰƮ)</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itempay</td>
			<td>�޿��ͼ�</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemdisuse</td>
			<td>���ս�</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>itemloss</td>
			<td>����ս�</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>shopitemsample</td>
			<td>���û��(����)</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>shopitemloss</td>
			<td>����ս�(����)</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>parcelloss</td>
			<td>�ù�н�</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>donation</td>
			<td>���</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>csservice</td>
			<td>CS���</td>
		</tr>
		</table>
	</p>
	</div>
	 
	 <%set oipchul = nothing
	 
%>
	<script type="text/javascript">  
		document.frmEapp.tC.value = document.all.divEapp.innerHTML.replace(/\r|\n/g,"");
	 	document.frmEapp.submit();
		</script>
	<!--/���ڰ���-->

