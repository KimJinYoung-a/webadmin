<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : �������� cs����
' Hieditor : 2011.03.15 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/shopcscenter_order_cls.asp"-->
<!-- #include virtual="/admin/offshop/shopcscenter/cscenter_Function_off.asp"-->

<%
dim csmasteridx ,ioneas,i ,ioneasDetail ,deliverydivname
	csmasteridx = request("csmasteridx")

set ioneas = new corder
	ioneas.FRectMakerID = session("ssBctID")
	ioneas.FRectCsAsID = csmasteridx
	ioneas.fGetOneCSASMaster

if (ioneas.FResultCount<1) then
    response.write "<script>"
    response.write "	alert('��ȿ�� ������ȣ�� �ƴմϴ�.');"
    response.write "	history.back();"
    response.write "</script>"
    response.write dbget.close()	:	response.End
end if

set ioneasDetail= new corder
	ioneasDetail.FRectCsAsID = csmasteridx
	ioneasDetail.fGetCsDetailList

if ioneas.FOneItem.Fdivcd = "A030" then
	deliverydivname = "A/S�Ϸ��ļ�����"
elseif ioneas.FOneItem.Fdivcd = "A031" then
	deliverydivname = "A/S��ü"
end if
%>

<script language='javascript'>

function ViewOrderDetail(frm,orgmasteridx){
	var props = "width=600, height=600, location=no, status=yes, resizable=no,";
	window.open("about:blank", "ViewOrderDetail", props);

    frm.target = 'ViewOrderDetail';
    frm.masteridx.value = orgmasteridx;
    frm.action="/common/offshop/shopcscenter/upche_viewordermaster.asp";
	frm.submit();
}

function SaveFin(frm){
	//alert('��� �غ����Դϴ�.');
	//return;

	if (frm.finishmemo.value.length<1){
		alert('ó�� ������ �Է��� �ּ���.');
		frm.finishmemo.focus();
		return;
	}

	if (frm.songjangdiv.value.length<1){
		alert('����� �ù�縦 �Է��� �ּ���.');
		frm.songjangdiv.focus();
		return;
	}

	if (frm.songjangno.value.length<1){
		alert('����� ��ȣ�� �Է��� �ּ���.');
		frm.songjangno.focus();
		return;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.submit();
	}
}

//��üa/s , ��üa/s(����ȸ��) �ּ��� ����
function popEditCsDelivery(CsAsID){
    var window_width = 600;
    var window_height = 450;

    var popEditCsDelivery=window.open("/admin/offshop/shopcscenter/action/pop_CsDeliveryEdit.asp?CsAsID=" + CsAsID ,"popEditCsDelivery","width=600 height=500 scrollbars=yes resizable=yes");
    popEditCsDelivery.focus();
}

</script>

<% if ioneas.Ftotalcount > 0 then %>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="FFFFFF">
	<tr>
		<td>
			<% getcurrstate_table ioneas.FOneItem.Fcurrstate,ioneas.FOneItem.Fdivcd %>
		</td>
	</tr>
	</table>
<% end if %>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="upche_csprocess.asp">
<input type="hidden" name="orgmasteridx" value="<%= ioneas.FOneItem.forgmasteridx %>">
<input type="hidden" name="finishuser" value="<%= session("ssBctID") %>">
<input type="hidden" name="masteridx" value="<%= ioneas.FOneItem.fmasteridx %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<b>���CS ó���亯</b>
		&nbsp;&nbsp;
    	�ۼ��� : <b><%= ioneas.FOneItem.Fregdate %></b>
    	&nbsp;&nbsp;
    	<% if not IsNULL(ioneas.FOneItem.Ffinishdate) then %>
    	�Ϸ��� : <b><%= ioneas.FOneItem.Ffinishdate %></b>
    	<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ</td>
	<td>
		<%= ioneas.FOneItem.Forderno %>
		<input type="button" class="button" value="�ֹ��󼼺���" onclick="ViewOrderDetail(frmshow,'<%= ioneas.FOneItem.forgmasteridx %>');">
	</td>
	<td width="45%" rowspan="5" valign="top">
		<%
		if ioneas.FOneItem.Fdivcd="A030" then
		%>
			* ��üA/S
			<br><br>��ü���� a/s ��ǰ�� ���� �Ϸ���, �����̳� ���в� �߼��ϴ� ���� �Դϴ�.
			<br>�����̳� ���в� �߼��Ͻ� �ù� ���� ������ �Է��� �ּ���.
		<%
		elseif ioneas.FOneItem.Fdivcd="A031" then
		%>
			* ��üA/S(����ȸ��)
			<br><br>���忡�� ���в� ������ a/s ��ǰ��, ��ü�� �߼��ϴ� ���� �Դϴ�.
		<% end if %>

	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td><%= ioneas.FOneItem.FCustomerName %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td><%= ioneas.FOneItem.FTitle %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td><%= replace(ioneas.FOneItem.Fcontents_jupsu,VbCrlf,"<br>") %></td>
</tr>
<% if (ioneasDetail.FResultCount>0) then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("tabletop") %>">������ǰ</td>
    <td>
        <table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#CCCCCC" class="a">
        <tr bgcolor="<%= adminColor("topbar") %>" align="center">
            <td width="100">��ǰ�ڵ�</td>
            <td width="100">�귣��ID</td>
            <td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
            <td width="50">�ǸŰ�</td>
            <td width="40">����</td>
        </tr>
        <% for i=0 to ioneasDetail.FResultCount-1 %>
        <tr bgcolor="#FFFFFF" align="center">
            <td>
            	<%=ioneasDetail.FItemList(i).fitemgubun%>-<%=FormatCode(ioneasDetail.FItemList(i).fitemid)%>-<%=ioneasDetail.FItemList(i).fitemoption%>
            </td>
            <td>
            	<%=ioneasDetail.FItemList(i).fmakerid%>
            </td>
            <td align="left">
            	<%= ioneasDetail.FItemList(i).Fitemname %>
            	<% if ioneasDetail.FItemList(i).Fitemoptionname<>"" then %>
            	<br>
            	<font color="blue">[<%= ioneasDetail.FItemList(i).Fitemoptionname %>]</font>
            	<% end if %>
            </td>
            <td align="right"><%= FormatNumber(ioneasDetail.FItemList(i).Fsellprice,0) %></td>
            <td align="center"><%= ioneasDetail.FItemList(i).Fitemno %></td>
        </tr>
        <% next %>
        </table>
    </td>
</tr>
<% end if %>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<b>���CS ó������ۼ�</b>
		&nbsp;&nbsp;
		*ó�� ���� �Է½� <font color=red>�����ȣ</font>�� �󼼳����� ������ �ּ���
	</td>
</tr>
<% if ioneas.FOneItem.Fdivcd = "A030" or ioneas.FOneItem.Fdivcd = "A031" then %>
<tr bgcolor="#FFFFFF">
	<td width="80" bgcolor="<%= adminColor("tabletop") %>"><%= deliverydivname %></td>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
		    <td width="50" bgcolor="<%= adminColor("pink") %>">�޴º�</td>
		    <td width="80" bgcolor="#FFFFFF"><%= ioneas.FOneItem.Freqname %></td>
		    <td width="50" bgcolor="<%= adminColor("pink") %>">����ó</td>
		    <td bgcolor="#FFFFFF"><%= ioneas.FOneItem.Freqphone %> / <%= ioneas.FOneItem.Freqhp %></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("pink") %>">�ּ�</td>
		    <td colspan="3" bgcolor="#FFFFFF">
				[<%= ioneas.FOneItem.Freqzipcode %>] <%= ioneas.FOneItem.Freqzipaddr %> &nbsp;<%= ioneas.FOneItem.FReqAddress %>

				<% if (ioneas.FOneItem.Fcurrstate="B001") then %>
					<%
					'/��üa/s �̸鼭 ��ü�ϰ�� ..  ��üa/s(����ȸ��) �̸鼭 ������ ��쿡�� .. �ּ� ���� ����
					if (ioneas.FOneItem.Fdivcd="A030" and C_IS_Maker_Upche) or (ioneas.FOneItem.Fdivcd="A031" and C_IS_SHOP) then
					%>
					    <input class="button" type="button" value="�ּҺ���" onclick="popEditCsDelivery('<%= ioneas.FOneItem.Fasid %>');" >
					<%
					'/�����ڳ� , �������ΰ����� �ϰ�� ���� ����
					elseif C_ADMIN_AUTH or C_OFF_AUTH then
					%>
						 <input class="button" type="button" value="�ּҺ���(�����ڸ��)" onclick="popEditCsDelivery('<%= ioneas.FOneItem.Fasid %>');" >
					<% end if %>
				<% else %>
					<input class="button" type="button" value="�ּҺ���Ұ�" onclick="alert('�������¿����� ���氡�� �մϴ�.');" >
				<% end if %>
		    </td>
		</tr>
		</table>
	</td>
	<td width="45%" rowspan="3" valign="top">
		<%
		if ioneas.FOneItem.Fdivcd="A030" then
		%>
			*ó���������� �Էµ� ������ ����� ��ü�� �����ϴ� �����Դϴ�.
			<br>(���Բ� ���µǴ� ������ �ƴմϴ�.)
			<br>
			<br><font color="red">�����̳� ���в� �����, �ù������� �� �Է� ��Ź�帳�ϴ�</font>
			<br>
			<br><font color="blue">*ó������ �Է¿�û����</font>
			<br>����� :
			<br>��Ÿ���� :
			<br><font color="blue">*�� ������ ī���ϼż�, ó�����뿡 �����ֽø� �����ϰڽ��ϴ�.</font>
		<%
		elseif ioneas.FOneItem.Fdivcd="A031" then
		%>
			*ó���������� �Էµ� ������ ����� ��ü�� �����ϴ� �����Դϴ�.
			<br>(���Բ� ���µǴ� ������ �ƴմϴ�.)
			<br>
			<br><font color="red">��ü�� �����, �ù������� �� �Է� ��Ź�帳�ϴ�</font>
			<br>
			<br><font color="blue">*ó������ �Է¿�û����</font>
			<br>����� :
			<br>��Ÿ���� :
			<br><font color="blue">*�� ������ ī���ϼż�, ó�����뿡 �����ֽø� �����ϰڽ��ϴ�.</font>
		<% end if %>

	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">ó������</td>
	<td>
		<textarea class="textarea" name="finishmemo" cols="60" rows="8" class="a"><%= ioneas.FOneItem.Fcontents_finish %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">���ÿ����</td>
	<td>
		<% drawSelectBoxDeliverCompany "songjangdiv",ioneas.FOneItem.FSongjangdiv %>
		<input type="text" class="text" name="songjangno" value="<%= ioneas.FOneItem.Fsongjangno %>" size="14" maxlength="14">
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<%
		'/��üó���Ϸ� / ����ó���Ϸ� / ����ó���Ϸ� ���°� �ƴҰ��
		if ioneas.FOneItem.Fcurrstate <> "B006" and ioneas.FOneItem.Fcurrstate <> "B007" and ioneas.FOneItem.Fcurrstate <> "B008" then

			'/��üa/s �̸鼭 ��ü�ϰ�� ..  ��üa/s(����ȸ��) �̸鼭 ������ ��쿡�� .. �ּ� ���� ����
			if (ioneas.FOneItem.Fdivcd="A030" and C_IS_Maker_Upche) or (ioneas.FOneItem.Fdivcd="A031" and C_IS_SHOP) then
		%>
				<input type="button" class="button" value="��üó���Ϸ�" onclick="javascript:SaveFin(frm);">
		<%
			'/�����ڳ� , �������ΰ����� �ϰ�� ���� ����
			elseif C_ADMIN_AUTH or C_OFF_AUTH then
		%>
			 	<input type="button" class="button" value="��üó���Ϸ�(�����ڸ��)" onclick="javascript:SaveFin(frm);">
			<% else %>
				<input class="button" type="button" value="�Ϸ�ó���Ұ�" onclick="alert('���� ������ ���氡�� �մϴ�.');" >
			<% end if %>
		<% end if %>

		<input type="button" class="button" value="��Ϻ���" onClick="location.href='/common/offshop/shopcscenter/upche_cslist.asp';">
	</td>
</tr>
</form>
<form name="frmshow" method="post">
	<input type="hidden" name="masteridx" value="">
</form>
</table>
<!-- ǥ �ϴܹ� ��-->

<%
set ioneas = Nothing
set ioneasDetail = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->