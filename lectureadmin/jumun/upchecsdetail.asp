<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->

<%
dim idx
idx = RequestCheckvar(request("idx"),10)

dim ioneas,i
set ioneas = new CCSASList

ioneas.FRectMakerID = session("ssBctID")
ioneas.FRectCsAsID = idx
ioneas.GetOneCSASMaster



if (ioneas.FResultCount<1) then
    response.write "<script>alert('��ȿ�� ������ȣ�� �ƴմϴ�.');</script>"
    response.write dbget.close()	:	response.End
end if

dim ioneasDetail
set ioneasDetail= new CCSASList
ioneasDetail.FRectCsAsID = idx
ioneasDetail.GetCsDetailList

%>
<script language='javascript'>
function ViewOrderDetail(frm){
	var props = "width=600, height=600, location=no, status=yes, resizable=no,";
	window.open("about:blank", "upcheorderpop", props);
    frm.target = 'upcheorderpop';
    frm.action="/designer/common/viewordermaster.asp"
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


	var ret = confirm('���� �Ͻðڽ��ϱ�?');


	if (ret){
		frm.submit();
	}
}

</script>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="upchecs_process.asp">
	<input type="hidden" name="orderserial" value="<%= ioneas.FOneItem.FOrderSerial %>">
	<input type="hidden" name="finishuser" value="<%= session("ssBctID") %>">
	<input type="hidden" name="id" value="<%= ioneas.FOneItem.FID %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<b>���CS ó���亯</b>
			&nbsp;&nbsp;
			�ۼ��� :
	        	<% if(Lcase(ioneas.FOneItem.Fwriteuser)=Lcase(ioneas.FOneItem.FUserID)) then %>
	        	<b>�� ���� �ۼ�</b>
	        	<% else %>
	        	�ٹ����� ������
	        	<b><% end if %></b>
        	&nbsp;&nbsp;
        	�ۼ��� : <b><%= CStr(ioneas.FOneItem.Fregdate) %></b>
        	&nbsp;&nbsp;
        	<% if not IsNULL(ioneas.FOneItem.Ffinishdate) then %>
        	�Ϸ��� : <b><%= CStr(ioneas.FOneItem.Ffinishdate) %></b>
        	<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ</td>
		<td>
			<%= ioneas.FOneItem.Forderserial %>
			<input type="button" class="button" value="�ֹ��󼼺���" onclick="ViewOrderDetail(frm);">
		</td>
		<td width="45%" rowspan="7" valign="top">
			<% if ioneas.FOneItem.Fdivcd="A000" then %> <!-- �±�ȯ ���� -->
				<b>* �±�ȯ ����</b>
			<% elseif ioneas.FOneItem.Fdivcd="A001" then %> <!-- ������߼� ���� -->
				<b>* ������߼� ����</b>
			<% elseif ioneas.FOneItem.Fdivcd="A004" then %> <!-- ��ǰ ���� -->
				<b>* ��ǰ���� ����</b>
				<br>��ǰ������ �ɰ��, ���Բ� �߼��Ͻ� �ù�� ��ȭ��ȣ�� �ȳ��ص帮��,
				<br>��ǰ�� ������ �ù�縦 ���� <font color="blue">���ҹݼ�</font>�� ���ֽõ��� �ȳ��� �ص帮�� �ֽ��ϴ�.
				<br><font color="blue">���� ��ǰ�� ���, ���ҹݼ����� �պ���ۺ� ������ �ݾ��� ���Բ� ȯ���ص帮��,
				<br>������ �ݾ��� ��ü���곻���� �ڵ����� ��ϵ˴ϴ�.</font>
				<br><font color="red">(�� 2,000�� / �պ� 4,000�� ����)</font>
				<br>
				<br>�ݼۻ�ǰ�� �����ϸ�, ��������� Ȯ���Ͻ� ��,
				<br>�Ʒ��� ó�����뿡 ������ �����ֽø�, �����Ϳ� ������ ���޵Ǹ�,
				<br>�����Ϳ��� ��ǰ���ó�� �� ��ȯ���� �����մϴ�.
				<br>
				<br>*ó�����μ���
				<br>1.����
				<br>2.��ü�Ϸ�ó�� --> �����Ϳ� ó����� ����
				<br>3.�����ͿϷ�ó�� --> ������ ó����� �ȳ� �� ���Ϲ߼�
			<% elseif ioneas.FOneItem.Fdivcd="A006" then %> <!-- ���� ���ǻ��� ���� -->
				<b>* ���� ���ǻ��� ����</b>
				<br>�ֹ��� Ȯ�� ��, ������ �ֹ����� ������ ��û�ϼ��� ���,
				<br>���� ���ǻ������� ��ϵ˴ϴ�.
				<br>ex)���������/��ǰ����/��ǰ�ɼǺ���
				<br>
				<br><font color="red">�ٹ����� �����Ϳ��� ������ ���ɿ��� Ȯ���� ���� �����帳�ϴ�.</font>
			<% else %>

			<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td><%= ioneas.FOneItem.FCustomerName %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">�����̵�</td>
		<td><%= ioneas.FOneItem.FUserID %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td><%= ioneas.FOneItem.FTitle %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="70" bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td><%= ioneas.FOneItem.Fgubun01Name %>&gt;&gt;<%= ioneas.FOneItem.Fgubun02Name %></td>
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
	            <td width="50">�̹���</td>
	            <td width="50">��ǰ�ڵ�</td>
	            <td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
	            <td width="50">�ǸŰ�</td>
	            <td width="40">����</td>
	        </tr>
	        <% for i=0 to ioneasDetail.FResultCount-1 %>
	        <tr bgcolor="#FFFFFF" align="center">
	            <td><img src="<%= ioneasDetail.FItemList(i).FSmallImage %>" width="50"></td>
	            <td><%= ioneasDetail.FItemList(i).FItemID %></td>
	            <td align="left">
	            	<%= ioneasDetail.FItemList(i).Fitemname %>
	            	<% if ioneasDetail.FItemList(i).Fitemoptionname<>"" then %>
	            	<br>
	            	<font color="blue">[<%= ioneasDetail.FItemList(i).Fitemoptionname %>]</font>
	            	<% end if %>
	            </td>
	            <td align="right"><%= FormatNumber(ioneasDetail.FItemList(i).Fitemcost,0) %></td>
	            <td align="center"><%= ioneasDetail.FItemList(i).Fitemno %></td>
	        </tr>
	        <% next %>
	        </table>
	    </td>
	</tr>
	<% end if %>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<b>���CS ó������ۼ�</b>
			&nbsp;&nbsp;
			*ó�� ���� �Է½� <font color=red>�����ȣ</font>�� �󼼳����� ������ �ּ���
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">ó������</td>
		<td>
			<textarea class="textarea" name="finishmemo" cols="60" rows="8" class="a"><%= ioneas.FOneItem.Fcontents_finish %></textarea>
		</td>
		<td width="45%" rowspan="2" valign="top">
			<% if ioneas.FOneItem.Fdivcd="A000" then %> <!-- �±�ȯ ���� -->
				*ó���������� �Էµ� ������ �����Ϳ� ���޵Ǵ� �����Դϴ�.
				<br>(���Բ� ���µǴ� ������ �ƴմϴ�.)
				<br>
				<br><font color="red">�±�ȯ��ǰ �����, �ù������� �� �Է� ��Ź�帳�ϴ�.</font>
				<br>
				<br><font color="blue">*ó������ �Է¿�û����</font>
				<br>����� :
				<br>��Ÿ���� :
				<br><font color="blue">*�� ������ ī���ϼż�, ó�����뿡 �����ֽø� �����ϰڽ��ϴ�.</font>
			<% elseif ioneas.FOneItem.Fdivcd="A001" then %> <!-- ������߼� ���� -->
				*ó���������� �Էµ� ������ �����Ϳ� ���޵Ǵ� �����Դϴ�.
				<br>(���Բ� ���µǴ� ������ �ƴմϴ�.)
				<br>
				<br><font color="red">�±�ȯ��ǰ �����, �ù������� �� �Է� ��Ź�帳�ϴ�.</font>
				<br>
				<br><font color="blue">*ó������ �Է¿�û����</font>
				<br>����� :
				<br>��Ÿ���� :
				<br><font color="blue">*�� ������ ī���ϼż�, ó�����뿡 �����ֽø� �����ϰڽ��ϴ�.</font>
			<% elseif ioneas.FOneItem.Fdivcd="A004" then %> <!-- ��ǰ ���� -->
				*ó���������� �Էµ� ������ �����Ϳ� ���޵Ǵ� �����Դϴ�.
				<br>(���Բ� ���µǴ� ������ �ƴմϴ�.)
				<br>
				<br><font color="red">��ǰ��ǰ �԰� �Ϸ� ��, ó������ �Է°� �Բ� �Ϸ�ó�� ��Ź�帳�ϴ�.</font>
				<br>
				<br><font color="blue">*ó������ �Է¿�û����</font>
				<br>��ǰ��� : ������ / ����
				<br>��ǰ���� : �ҷ���ǰ / ����ǰ
				<br>ȯ�Ұ��� : ����� + ���¹�ȣ + �����ָ�(������ ÷���� ���)
				<br>��Ÿ���� :
				<br><font color="blue">*�� ������ ī���ϼż�, ó�����뿡 �����ֽø� �����ϰڽ��ϴ�.</font>
			<% elseif ioneas.FOneItem.Fdivcd="A006" then %> <!-- ���� ���ǻ��� ���� -->
				*ó���������� �Էµ� ������ �����Ϳ� ���޵Ǵ� �����Դϴ�.
				<br>(���Բ� ���µǴ� ������ �ƴմϴ�.)
				<br>
				<br><font color="red">�����Ϳ��� ��û�� ������ǻ��׿� ���� ó�������� �˷��ֽñ� �ٶ��ϴ�.</font>
				<br>�߼� ��, �� ������ Ȯ���ϼ��� ��쿡��, �̹ݿ� ���� �Ϸ�ó�� ��Ź�帳�ϴ�.
			<% else %>

			<% end if %>



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
		<% if ioneas.FOneItem.Fcurrstate="B007" then %>

		<% else %>
			<input type="button" class="button" value="�Ϸ�ó��" onclick="javascript:SaveFin(frm);">
	    <% end if %>
			<input type="button" class="button" value="��Ϻ���" onClick="location.href='/designer/jumunmaster/upchecslist.asp';">
		</td>
	</tr>
	</form>
</table>

<!-- ǥ �ϴܹ� ��-->

<%
set ioneas = Nothing
set ioneasDetail = Nothing
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->