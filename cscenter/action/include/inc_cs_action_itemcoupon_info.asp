<%
'###########################################################
' Description : ��ǰ�������� ����
' Hieditor : 2023.10.16 �ѿ�� ����
'###########################################################
%>
<%
' �ֹ���� or (��ǰ����(��ü���) or ȸ����û(�ٹ����ٹ��))
if (IsCSCancelProcess(divcd) or IsCSReturnProcess(divcd)) then
%>
	<% if oCsItemCoupon.FResultCount>0 then %>
		<tr bgcolor="FFFFFF" align="center">
			<td>������</td>
			<td width="60">���ΰ�</td>
			<td width="150">��ȿ�Ⱓ</td>
			<td width="80">����</td>
		</tr>
		<% for i = 0 to oCsItemCoupon.FResultCount-1 %>
			<tr bgcolor="FFFFFF" align="center">
				<td align="left" >
					<%= oCsItemCoupon.FItemList(i).fitemcouponname %>
					<br>�����ڵ�:<%= oCsItemCoupon.FItemList(i).fitemcouponidx %>
				</td>
				<td >
					<%= oCsItemCoupon.FItemList(i).GetDiscountStr %>
				</td>
				<td >
					<%= ChkIIF(Right(oCsItemCoupon.FItemList(i).Fitemcouponstartdate,8)="00:00:00",Left(oCsItemCoupon.FItemList(i).Fitemcouponstartdate,10),oCsItemCoupon.FItemList(i).Fitemcouponstartdate) %>
					~
					<%= ChkIIF(Right(oCsItemCoupon.FItemList(i).Fitemcouponexpiredate,8)="23:59:59",Left(oCsItemCoupon.FItemList(i).Fitemcouponexpiredate,10),oCsItemCoupon.FItemList(i).Fitemcouponexpiredate) %>
				</td>
				<td >
					<%= oCsItemCoupon.FItemList(i).GetOpenStateName %>

					<% if (oCsItemCoupon.FItemList(i).forderserial="" or isnull(oCsItemCoupon.FItemList(i).forderserial)) and oCsItemCoupon.FItemList(i).fusedyn<>"Y" then %>
						<br>�����̻��
					<% else %>
						<br>�������
					<% end if %>

					<% if not(oCsItemCoupon.FItemList(i).IsItemCouponCopyValid) then %>
						<br><font color="red">��߱޺Ұ�</font>
					<% else %>
						<br>��߱ް���
					<% end if %>
				</td>
			</tr>
		<% next %>
	<% end if %>
<% end if %>