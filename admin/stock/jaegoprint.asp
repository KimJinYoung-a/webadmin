<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����ľǸ���Ʈ���
' History : 2007�� 7�� 13�� �ѿ�� ����
' 			2007�� 12�� 4�� �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/jaegostock.asp"-->

<%
dim idx, fidx, sql , i 				'��������
	idx = request("idx")				'�ε������� �޾ƿ´�.
	fidx = Left(idx,Len(idx)-1)			'�޾ƿ� �ε��� ���� �ϳ������� ����
	 
dim oip1 						'Ŭ��������
	set oip1 = new Cfitemlist		'������ ��Ż�� �ֱ�
	oip1.Frectidx = fidx
	oip1.fprintlist()				'Ŭ�������� 
%>


<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<input type="button" value="��Ʈ����ϱ�" onclick="javascript:window.print();">
        	<font color="red"><strong>�ɼ��� ���� ��ǰ�� ��� "0000"���� ǥ��˴ϴ�.</strong></font>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">�̹���</td>
		<td width="40">��ǰ<br>�ڵ�</td>
		<td width="100">�귣��ID</td>
		<td>��ǰ��[�ɼ�]</td>
		<td width="40">����<br>���</td>
		<td width="40">���<br>�ľ�</td>
		<td width="80">�۾�����<br>�۾������Ͻ�</td>
	</tr>
<% 
dim sql3
if oip1.FTotalCount > 0 then 		'���ڵ� ���� 0���� ũ�� 
%>	 
	<% for i=0 to oip1.FTotalCount - 1 %>
		<form name="frmBuyPrc<%=i%>" method="get">			<!--for�� �ȿ��� i ���� ������ ����-->
		<input type="hidden" name="mode">
		<tr align="center" bgcolor="#FFFFFF">
			<td><img src="<%= oip1.flist(i).fsmallimage %>" width=50 height=50><input type="hidden" name="smallimage" value="<%= oip1.flist(i).fsmallimage %>"></td>	<!--'�̹��� -->
			<td><%= oip1.flist(i).fitemid %><input type="hidden" name="fitemid" value="<%= oip1.flist(i).fitemid %>"></td>				 					<!--'��ǰ��ȣ	 -->
			<td><%= oip1.flist(i).fmakerid %><input type="hidden" name="fmakerid" value="<%= oip1.flist(i).fmakerid %>"></td>									 <!--'�귣��id -->
		<!--��ǰ����� -->
			<td align="left">
				<%= oip1.flist(i).fitemname %><input type="hidden" name="fitemname" value="<%= oip1.flist(i).fitemname %>">
				<br>
				<font color="blue">
				<%= oip1.flist(i).fitemoptionname %><input type="hidden" name="itemoptionname" value="<%= oip1.flist(i).fitemoptionname %>">
				</font>
			</td>				
		<!--��ǰ�� -->									
			<td><%= oip1.flist(i).frealstock %><input type="hidden" name="frealstock" value="<%= oip1.flist(i).frealstock %>"></td>									 <!--'����ľǿ���� -->
			<td></td>															
			<td>
				<!--����ľǶ�����-->
					<% if oip1.flist(i).fstatecd = 1 then %>
						 �۾�����
					<% elseif oip1.flist(i).fstatecd = 5 then %>
						 ����ľǿϷ�
					<% elseif oip1.flist(i).fstatecd = 7 then %>
						 �Ϸ�(�ݿ���)
					<% elseif oip1.flist(i).fstatecd = 8 then %>
						 �Ϸ�(�̹ݿ�)
					<% end if %>
				<!--����ľǶ���-->
			</td>						
		</tr>
		</form>	
		

	<% next %>
	
<% else %>
	<tr bgcolor="#FFFFFF">
	<td colspan=11 align=center>[ �˻������ �����ϴ�. ]</td>
	</tr>
<% end if %>
</table>

<%
set oip1 = nothing
%>	
<!-- #include virtual="/lib/db/dbclose.asp" -->
<script language="javascript">
opener.location.reload();
</script>
