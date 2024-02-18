<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/checkPartnerLog.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<%
dim orderserial
orderserial = requestCheckVar(request("orderserial"),11)

dim oldjumun

dim ojumun
set ojumun = new CJumunMaster
ojumun.FRectOrderSerial = orderserial

if (ojumun.FRectOrderSerial<>"") then
	'// �ֹ� ������
    ojumun.SearchJumunList
end if

if (ojumun.FResultCount<1) and (ojumun.FRectOrderSerial<>"") then
	'// 6���� ����
	oldjumun = "on"
	ojumun.FRectOldJumun = oldjumun
    ojumun.SearchJumunList
end if

if (ojumun.FResultCount<1) then
    dbget.close()	:	response.End
end If

'' 17032324862 CS�� �߰� üũ �ʿ���. 2017/04/20 ----------------
Dim isInStarView   : isInStarView   = FALSE
dim isRecentCsExists : isRecentCsExists = FALSE

isInStarView = (oldjumun = "on")  ''�����ֹ��� ��ǥ��

if (NOT isInStarView) then
    if Not IsNull(ojumun.FMasterItemList(0).Fbeadaldate) And Not IsNull(ojumun.FMasterItemList(0).FIpkumDiv) Then
        If (ojumun.FMasterItemList(0).FIpkumDiv = "8") And (DateDiff("d", ojumun.FMasterItemList(0).Fbeadaldate, Now()) > 10) Then
            isInStarView = TRUE
        end if
    end if
end if

if (isInStarView) then
    isRecentCsExists = Fn_getRecentUpcheCSExsists(orderserial,session("ssBctID"))
    if (isRecentCsExists) then
        isInStarView = FALSE
    end if
end if
'' ---------------------------------------------------------------

if (isInStarView) Then
		'// ���Ϸ� 10�� ���Ŀ��� �������� ǥ�� ����. 2017-04-11, skyer9
		If Not IsNull(ojumun.FMasterItemList(0).FBuyName) then
			ojumun.FMasterItemList(0).FBuyName = Left(ojumun.FMasterItemList(0).FBuyName,1) & "**"
		End If
		If Not IsNull(ojumun.FMasterItemList(0).FReqName) then
			ojumun.FMasterItemList(0).FReqName = Left(ojumun.FMasterItemList(0).FReqName,1) & "**"
		End If
		If Not IsNull(ojumun.FMasterItemList(0).FBuyPhone) Then
			if (Len(ojumun.FMasterItemList(0).FBuyPhone) > 4) then
				ojumun.FMasterItemList(0).FBuyPhone = Left(ojumun.FMasterItemList(0).FBuyPhone, Len(ojumun.FMasterItemList(0).FBuyPhone) - 4) & "****"
			End If
		End If
		If Not IsNull(ojumun.FMasterItemList(0).FBuyHp) Then
			if (Len(ojumun.FMasterItemList(0).FBuyHp) > 4) then
				ojumun.FMasterItemList(0).FBuyHp = Left(ojumun.FMasterItemList(0).FBuyHp, Len(ojumun.FMasterItemList(0).FBuyHp) - 4) & "****"
			End If
		End If
		If Not IsNull(ojumun.FMasterItemList(0).FReqPhone) Then
			if (Len(ojumun.FMasterItemList(0).FReqPhone) > 4) then
				ojumun.FMasterItemList(0).FReqPhone = Left(ojumun.FMasterItemList(0).FReqPhone, Len(ojumun.FMasterItemList(0).FReqPhone) - 4) & "****"
			End If
		End If
		If Not IsNull(ojumun.FMasterItemList(0).FReqHp) Then
			if (Len(ojumun.FMasterItemList(0).FReqHp) > 4) then
				ojumun.FMasterItemList(0).FReqHp = Left(ojumun.FMasterItemList(0).FReqHp, Len(ojumun.FMasterItemList(0).FReqHp) - 4) & "****"
			End If
		End If
		If Not IsNull(ojumun.FMasterItemList(0).FReqZipCode) Then
			ojumun.FMasterItemList(0).FReqZipCode = "*****"
		End If
		If Not IsNull(ojumun.FMasterItemList(0).FReqAddress) Then
			ojumun.FMasterItemList(0).FReqAddress = "(�ּһ���)"
		End If
End If

'// NOTE : �ֹ� ������ ��ü�� ������ ��, session("ssBctID") �� ��ġ�ϴ� �����ϸ� ȭ�鿡 ǥ���Ѵ�.

dim ix, i
dim sellsum
sellsum = 0

dim IsCanNotView, ValidUpcheItem : ValidUpcheItem = False

'// �ֹ� ������
ojumun.SearchJumunDetail orderserial
for ix=0 to ojumun.FJumunDetail.FDetailCount-1
    if ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid <>0 then
	    if UCase(ojumun.FJumunDetail.FJumunDetailList(ix).FMakerid)=UCase(session("ssBctID")) and (ojumun.FJumunDetail.FJumunDetailList(ix).Fisupchebeasong="Y") and (ojumun.FJumunDetail.FJumunDetailList(ix).Fcancelyn<>"Y") then
            if (ojumun.FJumunDetail.FJumunDetailList(ix).FCurrstate<3) or IsNULL(ojumun.FJumunDetail.FJumunDetailList(ix).FCurrstate) then
                IsCanNotView = true
            end if
            ValidUpcheItem = True
        end if
    end if
next

''����ǰ���� �߰� : ��ǰ ���� �� ������.
dim oGift
set oGift = new COrderGift

if (ojumun.FResultCount>0)  then
    oGift.FRectOrderSerial = orderserial
    oGift.FRectMakerid = session("ssBctId")
    oGift.FRectGiftDelivery = "Y"
    oGift.GetOneOrderGiftlist
end if

%>
<% if (Not ValidUpcheItem) then %>
<script language='javascript'>
alert('�ùٸ� �ֹ����� �ƴϰų� �˻��� �� �����ϴ�.');
window.close();
</script>
<% dbget.Close() : response.end %>
<% end if %>

<% if (IsCanNotView) then %>
<script language='javascript'>
alert('�������� �ֹ��̰ų� / �������� �Ǵ� �ֹ�Ȯ���� ���Ͻ� ��ǰ�� �ֽ��ϴ�.');
window.close();
</script>
<% dbget.Close() : response.end %>
<% end if  %>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td>
			<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="200" style="padding:5; border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>"  background="/images/menubar_1px.gif">
						<font color="#333333"><b>�ֹ��󼼳���</b></font>
					</td>
					<td align="right" style="border-bottom:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">

					</td>

				</tr>
			</table>
		</td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
        	<b>�ֹ���ȣ</b> : <%= ojumun.FMasterItemList(0).FOrderSerial %>&nbsp;&nbsp;&nbsp;&nbsp;
        	<b>�����ڸ�</b> : <%= ojumun.FMasterItemList(0).FBuyName %>
		</td>
	</tr>
    <tr>
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ</td>
		<td width="225" bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FOrderSerial %></td>
		<td width="100" bgcolor="<%= adminColor("tabletop") %>"><% if(FALSE) then %>����Ʈ<% end if %></td>
		<td width="225" bgcolor="#FFFFFF"><% if(FALSE) then %><%= ojumun.FMasterItemList(0).FSitename %><% end if %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).JumunMethodName %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�ֹ�����</td>
		<td bgcolor="#FFFFFF"><font color="<%= ojumun.FMasterItemList(0).IpkumDivColor %>"><%= ojumun.FMasterItemList(0).IpkumDivName %></font></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�ֹ���</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FRegDate %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ա���</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FIpkumDate %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FDiscountRate %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">��ҿ���</td>
		<td bgcolor="#FFFFFF"><font color="<%= ojumun.FMasterItemList(0).CancelYnColor %>"><%= ojumun.FMasterItemList(0).CancelYnName %></font></td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td colspan=3 bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FBuyName %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��������ȭ</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FBuyPhone %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�������ڵ���</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FBuyHp %></td>
	</tr>
	<% if (ojumun.FMasterItemList(0).Fjumundiv = "3") then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�������̸���</td>
		<td colspan=3 bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).Fbuyemail %></td>
	</tr>
	<% end if %>
<% ''������� detail.currstate�� %>
<% if (true) or ojumun.FMasterItemList(0).Fipkumdiv > 4 then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td colspan=3 bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FReqName %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��������ȭ</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FReqPhone %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�������ڵ���</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FReqHp %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�������ּ�</td>
		<td colspan="3" bgcolor="#FFFFFF">
		    <%=ojumun.FMasterItemList(0).FReqZipCode%>

		    <% if (FALSE) then %>
		    	<%= ojumun.FMasterItemList(0).FReqZipCode %>
	        <% end if %>
		<br>
		<%= ojumun.FMasterItemList(0).FReqZipAddr %>
		&nbsp;<%= ojumun.FMasterItemList(0).FReqAddress %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ����</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<%= nl2br(ojumun.FMasterItemList(0).FComment) %>
		</td>
	</tr>
	<% If Not IsNULL(ojumun.FMasterItemList(0).Freqdate) then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><b>�ö�� ������</b></td>
		<td colspan="3" bgcolor="#FFFFFF">
		<%= ojumun.FMasterItemList(0).Freqdate %> ��  <%= ojumun.FMasterItemList(0).GetReqTimeText() %><br>
		(�� �ö�� ������ǰ�� ��쿡�� �ش�, �Ϲݻ�ǰ�� �ش� �ȵ�.)
		</td>
	</tr>
	<% end if %>
	<% If Not IsNULL(ojumun.FMasterItemList(0).Fcardribbon) then %>
	<% If ojumun.FMasterItemList(0).Fcardribbon <> 3 then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">ī�� ���� ����</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<% If ojumun.FMasterItemList(0).Fcardribbon = 1 then %>ī��<% elseIf ojumun.FMasterItemList(0).Fcardribbon = 2 then %>����<% else %> ����<% End if %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">ī�帮���޼���</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<% if ojumun.FMasterItemList(0).Ffromname<>"" then %>
		From.<%= nl2br(ojumun.FMasterItemList(0).Ffromname) %><br>
		<% End if %>
		<%= nl2br(ojumun.FMasterItemList(0).Fmessage) %>
		</td>
	</tr>
	<% End if %>
	<% End if %>

<% else %>
	<tr align="center">
		<td colspan=4 bgcolor="#FFFFFF"><font color="blue"><b>��������� [��ü�ֹ��뺸] ���� ���Ŀ� Ȯ�ΰ����մϴ�.</b></font></td>
	</tr>
<% end if %>

</table>


<!--
<% if (FALSE) then %>
<table border="1" cellspacing="0" cellpadding="0" class="a">
	<tr>
		<td width="100">��ۿɼ�</td>
		<td width="200"><%= ojumun.FJumunDetail.BeasongOptionStr %></td>
	</tr>
	<tr>
		<td>��ۺ�</td>
		<td align="right"><%= FormatNumber(ojumun.FJumunDetail.BeasongPay,0) %></td>
	</tr>
</table>
<% end if %>
-->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
        	<b>�ֹ���ǰ����</b>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">��ǰ�ڵ�</td>
		<td width="50">�̹���</td>
		<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
		<td width="35">����</td>
		<td width="50">�ǸŰ���</td>
		<td width="35">���<br>����</td>
		<td width="35">���<br>����</td>
		<td width="120">�����<br>�������</td>
	</tr>
<% for ix=0 to ojumun.FJumunDetail.FDetailCount-1 %>
<% if  ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid <>0 then %>
	<% if UCase(ojumun.FJumunDetail.FJumunDetailList(ix).FMakerid)=UCase(session("ssBctID")) then %>
	<% sellsum = sellsum + ojumun.FJumunDetail.FJumunDetailList(ix).Fitemcost*ojumun.FJumunDetail.FJumunDetailList(ix).FItemNo %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid %></td>
		<td><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid %>" target="_blank"><img src="<%= ojumun.FJumunDetail.FJumunDetailList(ix).FImageSmall %>" border="0"></a></td>
		<td align="left">
			<%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemName %>
			<br>
			<% if ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName <> "" then %>
				<font color="blue">[<%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName %>]</font>
			<% end if %>
		</td>
		<td><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemNo %></td>
		<td align="right"><%= FormatNumber(ojumun.FJumunDetail.FJumunDetailList(ix).Fitemcost,0) %></td>
		<td>
			<% if ojumun.FJumunDetail.FJumunDetailList(ix).Fisupchebeasong="Y" then %>
			<font color="red">��ü</font>
			<% else %>
			10x10
			<% end if %>
		</td>
		<td>
			<font color="<%= ojumun.FJumunDetail.FJumunDetailList(ix).CancelStateColor %>"><%= ojumun.FJumunDetail.FJumunDetailList(ix).CancelStateStr %></font>
		</td>
		<td>
			<acronym title="<%= ojumun.FJumunDetail.FJumunDetailList(ix).Fbeasongdate %>"><%= Left(ojumun.FJumunDetail.FJumunDetailList(ix).Fbeasongdate,10) %></acronym><br>
            <%= ojumun.FJumunDetail.FJumunDetailList(ix).Fsongjangdivname %><br>
            <% if (FALSE) and (ojumun.FJumunDetail.FJumunDetailList(ix).FsongjangDiv = "24") then %>
            <a href="javascript:popDeliveryTrace('<%= ojumun.FJumunDetail.FJumunDetailList(ix).Ffindurl %>','<%= ojumun.FJumunDetail.FJumunDetailList(ix).Fsongjangno %>');"><%= ojumun.FJumunDetail.FJumunDetailList(ix).Fsongjangno %></a>
            <% else %>
            <a target="_blank" href="<%= ojumun.FJumunDetail.FJumunDetailList(ix).Ffindurl + ojumun.FJumunDetail.FJumunDetailList(ix).Fsongjangno %>"><%= ojumun.FJumunDetail.FJumunDetailList(ix).Fsongjangno %></a>
            <% end if %>
		</td>
	</tr>
	<% if (Not IsNULL(ojumun.FJumunDetail.FJumunDetailList(ix).Frequiredetail)) and (ojumun.FJumunDetail.FJumunDetailList(ix).Frequiredetail<>"") then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="8">
			<% if (ojumun.FJumunDetail.FJumunDetailList(ix).FItemNo>1) then %>
				<% for i=0 to ojumun.FJumunDetail.FJumunDetailList(ix).FItemNo-1 %>
					[<%= i+ 1 %>�� ��ǰ ����]<br>
					<%= nl2Br(splitValue(ojumun.FJumunDetail.FJumunDetailList(ix).Frequiredetail,CAddDetailSpliter,i)) %>
					<br>
				<% next %>
			<% else %>
				<%= nl2Br(ojumun.FJumunDetail.FJumunDetailList(ix).Frequiredetail) %>
			<% end if %>
		</td>
	</tr>
	<% end if %>
	<% end if %>
<% end if %>
<% next %>
	<tr align="center" bgcolor="#FFFFFF">
		<td>�հ�</td>
		<td colspan="4" align="right"><%= FormatNumber(sellsum,0) %></td>
		<td colspan="3"></td>
	</tr>
</table>

<p>
<% if oGift.FresultCount>0 then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
    <td width="50" align="center" >����ǰ</td>
    <td>
    <% for ix=0 to oGift.FResultCount -1 %>
        [<%= oGift.FItemList(ix).Fevt_name %>] <%= oGift.FItemList(ix).GetEventConditionStr %><br>
    <% next %>
    </td>
</tr>
</table>
<p>
<% end if %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="50" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">

		</td>
		<td colspan="15">
        	<font color="blue">
        		<b>�� �ڷ�� ����� ���� �����θ� ����ؾ� �մϴ�.<br>
				�̿��� �������� ���� ��,����� å���� �ش� ��ü���� �ֽ��ϴ�.</b>
			</font>
		</td>
	</tr>
</table>

<!-- ǥ �ϴܹ� ��-->

<%
set ojumun = Nothing
set oGift = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
