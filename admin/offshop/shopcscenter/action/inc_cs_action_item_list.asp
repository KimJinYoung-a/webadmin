<%
'###########################################################
' Description : ���� ������
' Hieditor : 2012.03.20 �ѿ�� ����
'###########################################################

dim IsItemCanceled, IsPossibleModifyItem, IsItemDisabled, IsItemChecked
dim ocsOrderDetail
dim BaesongMethod, SumBeasongPayNotCanceled, SumItemCostSumNotCanceled
dim strhtmldisabled, strhtmlcancel, strhtmlmodify

'������
set ocsOrderDetail = new COrder

	'//���� ���¿����� ��ü �ֹ����
	if (IsStatusRegister) then

		'/�ֹ� ���̺� masteridx
		ocsOrderDetail.FRectmasteridx = masteridx
	    ocsOrderDetail.fGetOrderDetailByCsDetail

	'//����, �Ϸ���¿����� ������ ������ ������
	else

		'/cs���̺� masteridx
		ocsOrderDetail.FRectCsAsID = csmasteridx
	    ocsOrderDetail.fGetCsDetailList
	end if

%>
<% if (IsDisplayItemList and (ocsOrderDetail.FTotalCount > 0)) then %>
<tr>
    <td >
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" width="80">������ǰ</td>
            <td bgcolor="#FFFFFF">
                <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
				<tr height="20" align="center" bgcolor="#F4F4F4">
					<td>����</td>
					<td>����</td>
					<td>��ǰ�ڵ�</td>
					<td>�귣��ID</td>
					<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
					<td>
						<% if (ocsaslist.FOneItem.IsCancelProcess_off) then %>
							���/���ֹ�
						<% else %>
							����/���ֹ�
						<% end if %>
					</td>
					<td>�ǸŰ���</td>
					<td>��������</td>
				</tr>
				<% '��ũ��Ʈ�� �ܼ�ȭ�ϱ� ���� �Ʒ��� ���� ���̸� �� ����� �д�.(orderdetailidx �� �Ѱ��� ���� 2���̻��� ��츦 �и��ؼ� �ۼ����� �ʾƵ� �ȴ�.) %>
				<input type="hidden" name="Deliverdetailidx">
				<input type="hidden" name="DeliverMakerid">
				<input type="hidden" name="Deliveritemcost">

				<input type="hidden" name="Deliverdetailidx">
				<input type="hidden" name="DeliverMakerid">
				<input type="hidden" name="Deliveritemcost">

				<input type="hidden" name="dummystarter" value="">
				<input type="hidden" name="orderdetailidx">
				<input type="hidden" name="odlvtype">
				<input type="hidden" name="itemno">
				<input type="hidden" name="regitemno">
				<input type="hidden" name="makerid">
				<input type="hidden" name="dummystopper" value="">

				<input type="hidden" name="dummystarter" value="">
				<input type="hidden" name="orderdetailidx">
				<input type="hidden" name="odlvtype">
				<input type="hidden" name="itemno">
				<input type="hidden" name="regitemno">
				<input type="hidden" name="makerid">
				<input type="hidden" name="dummystopper" value="">
				<%
				SumBeasongPayNotCanceled = 0
				SumItemCostSumNotCanceled = 0

				for i = 0 to ocsOrderDetail.FResultCount - 1

				IsItemCanceled = (ocsOrderDetail.FItemList(i).FCancelyn = "Y")

				strhtmlcancel = ""

				if (IsItemCanceled) then
					'���
					strhtmlcancel = "bgcolor='#CCCCCC' class='gray'"
				else
					if (IsPossibleCheckItem_off(divcd, IsOrderCanceled, IsItemCanceled) = true) then
						'üũ���ɻ�ǰ
						strhtmlcancel = "bgcolor='#FFFFFF'"
					else
						'üũ�Ұ���ǰ
						strhtmlcancel = "bgcolor='#EEEEEE' class='gray'"
					end if
				end if

				strhtmldisabled = ""

				if (IsStatusRegister) then
					if (IsPossibleCheckItem_off(divcd, IsOrderCanceled, IsItemCanceled) = true) then
						'üũ���ɻ�ǰ
						strhtmldisabled = ""
					else
						'üũ�Ұ���ǰ
						strhtmldisabled = "disabled"
					end if
				else
					'�������Ŀ��� �׻� üũ�Ǿ� �ְ�, �����Ұ�
					strhtmldisabled = "checked disabled"
				end if

				strhtmlmodify = ""

				if (IsStatusRegister or (IsStatusEdit and IsCSReturnProcess_off(divcd))) then
					if (IsPossibleCheckItem_off(divcd, IsOrderCanceled, IsItemCanceled) = true) then
						'üũ���ɻ�ǰ
						strhtmlmodify = ""
					else
						'üũ�Ұ���ǰ
						strhtmlmodify = "style='text-align:center;background-color:#DDDDFF;' readonly"
					end if
				else
					'�������Ŀ��� �����Ұ�
					strhtmlmodify = "style='text-align:center;background-color:#DDDDFF;' readonly"
				end if

				%>

				<%
				'��ǰ ����Ʈ
				distinctid = ocsOrderDetail.FItemList(i).Forgdetailidx
				%>
				<tr align="center" <%= strhtmlcancel %>>
					<td height="25">
						<input type="hidden" name="dummystarter" value="">
						<input type="checkbox" name="orderdetailidx" value="<%= ocsOrderDetail.FItemList(i).Forgdetailidx %>" onClick="AnCheckClick(this); CheckSelect(this);" <%= strhtmldisabled %>>
					</td>
					<td><font color="<%= ocsOrderDetail.FItemList(i).CancelStateColor %>"><%= ocsOrderDetail.FItemList(i).CancelYnName %></font></td>
					<td>
						<%=ocsOrderDetail.FItemList(i).fitemgubun%>-<%=CHKIIF(ocsOrderDetail.FItemList(i).fitemid>=1000000,Format00(8,ocsOrderDetail.FItemList(i).fitemid),Format00(6,ocsOrderDetail.FItemList(i).fitemid))%>-<%=ocsOrderDetail.FItemList(i).fitemoption%>
					</td>
					<td width="90">
						<acronym title="<%= ocsOrderDetail.FItemList(i).Fmakerid %>">
							<%= Left(ocsOrderDetail.FItemList(i).Fmakerid,32) %>
						</acronym>
					</td>
					<td align="left">
						<acronym title="<%= ocsOrderDetail.FItemList(i).FItemName %>"><%= DDotFormat(ocsOrderDetail.FItemList(i).FItemName,64) %></acronym>
						<% if (ocsOrderDetail.FItemList(i).FItemoptionName <> "") then %>
							<br><font color="blue">[<%= ocsOrderDetail.FItemList(i).FItemoptionName %>]</font><br>
						<% end if %>
						<div id="causepop_<%= distinctid %>" style="position:absolute;"></div>
					</td>
					<td>
						<input type="text" name="regitemno" onKeyUp="CheckMaxItemNo(this, <%= ocsOrderDetail.FItemList(i).FItemNo %>);" <%= strhtmlmodify %> value="<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo_off(IsStatusRegister) %>" size="2" style="text-align:center">
						/
						<input type="text" name="itemno" value="<%= ocsOrderDetail.FItemList(i).FItemNo %>" size="2" style="text-align:center;background-color:#DDDDFF;" readonly>
					</td>
					<input type="hidden" name="sellprice" value="<%= ocsOrderDetail.FItemList(i).fsellprice %>">
					<% if (IsItemCanceled) then %>
						<td align="right"><font color="gray"><%= FormatNumber(ocsOrderDetail.FItemList(i).fsellprice,0) %></font></td>
					<% elseif (ocsOrderDetail.FItemList(i).FItemNo < 1) then %>
						<td align="right"><font color="red"><%= FormatNumber(ocsOrderDetail.FItemList(i).fsellprice,0) %></font></td>
					<% else %>
					<td align="right">
						<font color="blue"><%= FormatNumber(ocsOrderDetail.FItemList(i).fsellprice,0) %></font>
					</td>
					<% end if %>
					<td align="center">
						<% if (IsStatusFinished) and ((divcd="A010") or (divcd="A011")) then %>
							<br><input type="button" class="button" value="�ҷ����" onClick="popBadItemReg('10<%= Format00(6,ocsOrderDetail.FItemList(i).FItemid) %><%= ocsOrderDetail.FItemList(i).FItemOption %>','<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) %>');">
						<% end if %>
						<% if ocsOrderDetail.FItemList(i).fcancelorgdetailidx <> "" then %>
							��һ�idx : <%=ocsOrderDetail.FItemList(i).fcancelorgdetailidx%>
						<% end if %>
					</td>
					<input type="hidden" name="makerid" value="<%= ocsOrderDetail.FItemList(i).Fmakerid %>">
					<input type="hidden" name="odlvtype" value="<%= ocsOrderDetail.FItemList(i).Fodlvtype %>">
					<input type="hidden" name="dummystopper" value="">
				</tr>
				<% next %>
            	</table>
            </td>
		</tr>
		</table>
	</td>
</tr>
<% end if %>