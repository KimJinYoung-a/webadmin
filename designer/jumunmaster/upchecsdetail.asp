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
idx = requestCheckVar(request("idx"), 32)

'// ===========================================================================
dim ioneas,i
set ioneas = new CCSASList

ioneas.FRectMakerID = session("ssBctID")
ioneas.FRectCsAsID = idx
ioneas.GetOneCSASMaster

'==============================================================================
''ȯ������
dim orefund

set orefund = New CCSASList

orefund.FRectCsAsID = requestCheckVar(request("idx"), 32)

orefund.GetOneRefundInfo



'// ===========================================================================
if (ioneas.FResultCount<1) then
    response.write "<script>alert('��ȿ�� ������ȣ�� �ƴմϴ�.');</script>"
    response.write dbget.close()	:	response.End
end if

dim ioneasDetail
set ioneasDetail= new CCSASList
ioneasDetail.FRectCsAsID = idx
ioneasDetail.GetCsDetailList


'// ===========================================================================
dim sqlStr

if IsNull(ioneas.FOneItem.Fconfirmdate) then
	sqlStr = " update [db_cs].[dbo].tbl_new_as_list set confirmdate = getdate() where id = " + CStr(idx) + " "
	dbget.Execute sqlStr
end if


'// ===========================================================================
dim IsChangeReturn
dim ioneRefasDetail

dim ioneRefas, IsRefASExist, refasid
dim chulgoyn
dim receiveyn
dim receiveonly

dim divcd, refdivcd

chulgoyn = "N"
receiveyn = ""
IsRefASExist = False
receiveonly = requestCheckVar(request("receiveonly"), 32)
set ioneRefas = new CCSASList
IsChangeReturn = False
refasid = 0

divcd = ioneas.FOneItem.FDivCD
if ((ioneas.FOneItem.FDivCD = "A000") or (ioneas.FOneItem.FDivCD = "A100")) then
	'// �±�ȯ���, ��ǰ���� �±�ȯ���

	if (ioneas.FOneItem.Fcurrstate >= "B006") then
		chulgoyn = "Y"
	end if

	ioneRefas.FRectMakerID = session("ssBctID")
	ioneRefas.FRectCsRefAsID = idx
	ioneRefas.GetOneCSASMaster

	refdivcd = ioneRefas.FOneItem.FDivCD

	if (ioneRefas.FResultCount>0) then
	    IsRefASExist = True
	    refasid = ioneRefas.FOneItem.FID

	    if (ioneRefas.FOneItem.Fcurrstate >= "B006") then
	    	receiveyn = "Y"
	    else
	    	receiveyn = "N"
	    end if
	end if

	IsChangeReturn = (ioneas.FOneItem.FDivCD = "A100")				'// ��ǰ���� �±�ȯ���

	set ioneRefasDetail = new CCSASList

	if (IsChangeReturn) then
		ioneRefasDetail.FRectMakerID = session("ssBctID")
		ioneRefasDetail.FRectCsRefAsID = idx
		ioneRefasDetail.GetCsDetailList
	end if
end if


'// ===========================================================================
dim currrowspan
dim tmpStr

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>

function ViewOrderDetail(frm){
	var props = "width=600, height=600, location=no, status=yes, resizable=no,";
	window.open("about:blank", "upcheorderpop", props);
    frm.target = 'upcheorderpop';
    frm.action="/designer/common/viewordermaster.asp"
	frm.submit();

}

function SaveReceiveFin(frm) {
	var ret = confirm('���� �Ͻðڽ��ϱ�?');


	if (ret){
		frm.submit();
	}
}

function trim(value) {
 return value.replace(/^\s+|\s+$/g,"");
}

function SaveFin(frm){
	//alert('��� �غ����Դϴ�.');
	//return;
	var val;

	frm.finishmemo.value = trim(frm.finishmemo.value);
	if (frm.finishmemo.value.length<1){
		alert('ó�� ������ �Է��� �ּ���.');
		frm.finishmemo.focus();
		return;
	}

	<% if (ioneas.FOneItem.FDivCD = "A100") then %>

		if (frm.customerrealbeasongpay) {
			if (frm.customerrealbeasongpay.value == "") {
				frm.customerrealbeasongpay.value = "0";
			}

			if (frm.customerrealbeasongpay.value*0 != 0) {
				alert("�� �߰� ��ۺ�� ���ڸ� �����մϴ�.");
				frm.customerrealbeasongpay.focus();
				return;
			}

			if (frm.customerrealbeasongpay.value != "0") {
				frm.customerreceiveyn.value = "Y";
			} else {
				frm.customerreceiveyn.value = "N";
			}
		}

	<% end if %>

	if ($("#needChkYN_N").val()) {
		if ($("#needChkYN_N").prop("checked") === false && $("#needChkYN_Y").prop("checked") === false) {
			alert("��ÿϷ� ���θ� �����ϼ���.");
			$("#needChkYN_N").focus();
			return;
		}

		if ($("#needChkYN_N").prop("checked") === true) {
			<% if (ioneas.FOneItem.FDivCD = "A000") then %>
			if ($("#needRefChkYN_N").prop("checked") === false && $("#needRefChkYN_Y").prop("checked") === false) {
				alert("��ȯȸ�������� ��ÿϷ� ���θ� �����ϼ���.");
				$("#needRefChkYN_N").focus();
				return;
			}
			<% End If %>
			if (confirm("�������� Ȯ������ ���� ��� �Ϸ�ó���˴ϴ�.\n\n��� �����Ͻðڽ��ϱ�?") === false) {
				return;
			}
		}
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');


	if (ret){
		frm.submit();
	}
}

function SetReceiveYes(frm) {
	// not used
	var ret = confirm('ȸ���Ϸ� ó�� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.receiveyn[0].checked = true;

		frm.submit();
	}
}



function GetRadioValue(obj) {
	for (var i=0; i < obj.length; i++) {
		if (obj[i].checked) {
			return obj[i].value;
		}
	}

	return "";
}

</script>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="upchecs_process.asp">
	<input type="hidden" name="orderserial" value="<%= ioneas.FOneItem.FOrderSerial %>">
	<input type="hidden" name="finishuser" value="<%= session("ssBctID") %>">
	<input type="hidden" name="id" value="<%= ioneas.FOneItem.FID %>">
	<input type="hidden" name="refasid" value="<%= refasid %>">
	<input type="hidden" name="receiveonly" value="<%= receiveonly %>">
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
		<td width="80" bgcolor="<%= adminColor("tabletop") %>" height="25">�ֹ���ȣ</td>
		<td>
			<%= ioneas.FOneItem.Forderserial %>
			<input type="button" class="button" value="�ֹ��󼼺���" onclick="ViewOrderDetail(frm);">
		</td>
		<td width="45%" rowspan="7" valign="top">
			<% if (ioneas.FOneItem.Fdivcd="A000") or (ioneas.FOneItem.Fdivcd="A012") or (ioneas.FOneItem.Fdivcd="A100") or (ioneas.FOneItem.Fdivcd="A112") then %> <!-- �±�ȯ ���� -->
				<b>* �±�ȯ ����</b>
			<% elseif ioneas.FOneItem.Fdivcd="A001" then %> <!-- ������߼� ���� -->
				<b>* ������߼� ����</b>
			<% elseif ioneas.FOneItem.Fdivcd="A004" then %> <!-- ��ǰ ���� -->
				<b>* ��ǰ���� ����</b>
				<br>��ǰ������ �ɰ��, ���Բ� �߼��Ͻ� �ù�� ��ȭ��ȣ�� �ȳ��ص帮��,
				<br>��ǰ�� ������ �ù�縦 ���� <font color="blue">���ҹݼ�</font>�� ���ֽõ��� �ȳ��� �ص帮�� �ֽ��ϴ�.
				<br><font color="blue">���� ��ǰ�� ���, �귣�� ��ǰ �Ϻι�ǰ�� ��� ��ǰ��ۺ�, ��ü ��ǰ�� ��� ���ҹݼ����� �պ���ۺ� ������ �ݾ��� ���Բ� ȯ���ص帮��,
				<br>������ �ݾ��� ��ü���곻���� �ڵ����� ��ϵ˴ϴ�.</font>
				<br><font color="red">(�� 2,500�� / �պ� 5,000�� ����)</font>
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
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td><%= ioneas.FOneItem.FCustomerName %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�����̵�</td>
		<td><%= ioneas.FOneItem.FUserID %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td><%= ioneas.FOneItem.FTitle %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="70" bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td><%= ioneas.FOneItem.Fgubun01Name %>&gt;&gt;<%= ioneas.FOneItem.Fgubun02Name %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="50">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td>
			<%
			tmpStr = replace(ioneas.FOneItem.Fcontents_jupsu,"<","&lt;")
			tmpStr = replace(tmpStr,">","&gt;")
			tmpStr = replace(tmpStr,VbCrlf,"<br>")
			%>
			<%= tmpStr %>
		</td>
	</tr>
	<% if (ioneasDetail.FResultCount>0) then %>
	<tr bgcolor="#FFFFFF">
	    <td bgcolor="<%= adminColor("tabletop") %>">������ǰ</td>
	    <td>
	        <table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#CCCCCC" class="a">
	        <tr bgcolor="<%= adminColor("topbar") %>" align="center" height="25">
	            <td width="50"></td>
	            <td width="50">�̹���</td>
	            <td width="50">��ǰ�ڵ�</td>
	            <td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
	            <td width="50">�ǸŰ�</td>
	            <td width="40">����</td>
	        </tr>
		<% if (IsChangeReturn) then %>
	    	<% for i=0 to ioneRefasDetail.FResultCount-1 %>
		        <tr bgcolor="#FFFFFF" align="center">
	   	            <td>���ֹ�</td>
		            <td><img src="<%= ioneRefasDetail.FItemList(i).FSmallImage %>" width="50"></td>
		            <td><%= ioneRefasDetail.FItemList(i).FItemID %></td>
		            <td align="left">
		            	<%= ioneRefasDetail.FItemList(i).Fitemname %>
		            	<% if ioneRefasDetail.FItemList(i).Fitemoptionname<>"" then %>
		            	<br>
		            	<font color="blue">[<%= ioneRefasDetail.FItemList(i).Fitemoptionname %>]</font>
		            	<% end if %>
		            </td>
		            <td align="right"><%= FormatNumber(ioneRefasDetail.FItemList(i).Fitemcost,0) %></td>
		            <td align="center"><%= ioneRefasDetail.FItemList(i).Fitemno %></td>
		        </tr>
	        <% next %>
	        <% for i=0 to ioneasDetail.FResultCount-1 %>
		        <tr bgcolor="#FFFFFF" align="center">
    	            <td>����&gt;</td>
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
        <% else %>
	        <% for i=0 to ioneasDetail.FResultCount-1 %>
		        <tr bgcolor="#FFFFFF" align="center">
    	            <td>������ǰ</td>
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
        <% end if %>
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
	<tr bgcolor="#FFFFFF" height="25">
	<% if (receiveonly = "Y") then %><!-- (�±�ȯ���, ��ǰ���� �±�ȯ���) + (�±�ȯȸ�� ��ϵ� ���) + (�±�ȯȸ�� �Է�) -->
		<td width="130" height="120" bgcolor="<%= adminColor("tabletop") %>">��� ó������</td>
		<td>
			<%= nl2br(ioneas.FOneItem.Fcontents_finish) %>
		</td>
	<% else %>
		<td width="130" bgcolor="<%= adminColor("tabletop") %>">ó������</td>
		<td>
			<textarea class="textarea" name="finishmemo" cols="60" rows="8" class="a"><%= ioneas.FOneItem.Fcontents_finish %></textarea>
		</td>
	<% end if %>

		<td width="45%" rowspan="20" valign="top">
			<% if (ioneas.FOneItem.Fdivcd="A000") or (ioneas.FOneItem.Fdivcd="A100") then %> <!-- �±�ȯ ���� -->
				<% if (receiveonly = "Y") then %>
					*ó���������� �Էµ� ������ �����Ϳ� ���޵Ǵ� �����Դϴ�.
					<br>(���Բ� ���µǴ� ������ �ƴմϴ�.)
					<br>
					<br><font color="red">�� �߰���ۺ� �ִ°�� ���ɾ��� �� �Է� ��Ź�帳�ϴ�.</font>
					<br>
					<br><font color="blue">*ó������ �Է¿�û����</font>
					<br>��Ÿ���� :
					<br><font color="blue">*�� ������ ī���ϼż�, ó�����뿡 �����ֽø� �����ϰڽ��ϴ�.</font>
				<% else %>
					*ó���������� �Էµ� ������ �����Ϳ� ���޵Ǵ� �����Դϴ�.
					<br>(���Բ� ���µǴ� ������ �ƴմϴ�.)
					<br>
					<br><font color="red">�±�ȯ��ǰ �����, �ù������� �� �Է� ��Ź�帳�ϴ�.</font>
					<br>
					<br><font color="blue">*ó������ �Է¿�û����</font>
					<br>����� :
					<br>��Ÿ���� :
					<br><font color="blue">*�� ������ ī���ϼż�, ó�����뿡 �����ֽø� �����ϰڽ��ϴ�.</font>
				<% end if %>
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


	<%
	'[�ڵ�����]
	'------------------------------------------------------------------------------
	'A008			�ֹ����
	'
	'A004			��ǰ����(��ü���)
	'A010			ȸ����û(�ٹ����ٹ��)
	'
	'A001			������߼�
	'A002			���񽺹߼�
	'
	'A200			��Ÿȸ��
	'
	'A000			�±�ȯ���
	'A100			��ǰ���� �±�ȯ���
	'
	'A009			��Ÿ����
	'A006			�������ǻ���
	'A700			��ü��Ÿ����
	'
	'A003			ȯ��
	'A005			�ܺθ�ȯ�ҿ�û
	'A007			ī��,��ü,�޴�����ҿ�û
	'
	'A011			�±�ȯȸ��(�ٹ����ٹ��)
	'A012			�±�ȯ��ǰ(��ü���)

	'A111			��ǰ���� �±�ȯȸ��(�ٹ����ٹ��)
	'A112			��ǰ���� �±�ȯ��ǰ(��ü���)
	%>

	<% if (receiveonly <> "Y") then %>
	<% if InStr(",A000,A100,A001,A002,A009,A006,A012,A004,", divcd) > 0 then %>
	<% if (divcd = "A004") then %>
	<tr bgcolor="#FFFFFF" height="30">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ǰ����</td>
		<td>
			<%= ioneas.FOneItem.Fgubun01Name %>&gt;&gt;<%= ioneas.FOneItem.Fgubun02Name %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="30">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ǰ��ۺ�</td>
		<td>
			<% if (orefund.FOneItem.Frefunddeliverypay<>0) then %>
			ȯ�ҽ� ��ǰ��ۺ� <%= FormatNumber(orefund.FOneItem.Frefunddeliverypay*-1, 0) %> �� ������ ȯ��
			<% else %>
			����
			<% end if %>
		</td>
	</tr>
	<% end if %>
	<tr bgcolor="#FFFFFF" height="30">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ÿϷ� ����</td>
		<td>
			<% if (ioneas.FOneItem.FneedChkYN="F") then %>
			<input type="radio" id="needChkYN_F" name="needChkYN" value="F" <%= CHKIIF(ioneas.FOneItem.FneedChkYN="F", "checked", "") %> > ������ Ȯ�ο�
			<% else %>
			<input type="radio" id="needChkYN_N" name="needChkYN" value="N" <%= CHKIIF(ioneas.FOneItem.FneedChkYN="N", "checked", "") %> > ��ÿϷ�(������ Ȯ�� ���ʿ�)
			<input type="radio" id="needChkYN_Y" name="needChkYN" value="Y" <%= CHKIIF(ioneas.FOneItem.FneedChkYN="Y", "checked", "") %> > ������ Ȯ�ο�û
			<% end if %>
		</td>
	</tr>
	<% If (divcd = "A000") Then %>
	<tr bgcolor="#FFFFFF" height="30">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ȯȸ��</td>
		<td>
			<input type="radio" id="needRefChkYN_N" name="needRefChkYN" value="N" <%= CHKIIF(receiveyn = "Y", "checked", "") %> > ��ÿϷ�(��ǰȸ���Ϸ� �Ǿ���)
			<input type="radio" id="needRefChkYN_Y" name="needRefChkYN" value="Y" > ��ǰȸ�� ����
		</td>
	</tr>
	<% end if %>
	<% end if %>
	<% end if %>

	<% if (receiveonly = "Y") then %>

		<!-- ============================ (�±�ȯ���, ��ǰ���� �±�ȯ���) + (�±�ȯȸ�� ��ϵ� ���) + (�±�ȯȸ�� �Է�) -->
		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%= adminColor("tabletop") %>" height="30">��� �����</td>
			<td>
				<%= DeliverDivCd2Nm(ioneas.FOneItem.FSongjangdiv) %>
				<%= ioneas.FOneItem.Fsongjangno %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">��� ����</td>
			<td>
				<% if (chulgoyn = "Y") then %>
					���Ϸ�
				<% elseif (chulgoyn = "N") then %>
					<font color="blue">�������</font>
				<% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">ȸ�� ó������</td>
			<td>
				<textarea class="textarea" name="finishmemo" cols="60" rows="8" class="a"><%= ioneRefas.FOneItem.Fcontents_finish %></textarea>
			</td>
		</tr>
		<% if InStr(",A000,A100,A001,A002,A009,A006,A012,A004,", refdivcd) > 0 then %>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ÿϷ� ����</td>
			<td>
				<% if (ioneRefas.FOneItem.FneedChkYN="F") then %>
				<input type="radio" id="needChkYN_F" name="needChkYN" value="F" <%= CHKIIF(ioneRefas.FOneItem.FneedChkYN="F", "checked", "") %> > ������ Ȯ�ο�
				<% else %>
				<input type="radio" id="needChkYN_N" name="needChkYN" value="N" <%= CHKIIF(ioneRefas.FOneItem.FneedChkYN="N", "checked", "") %> > ��ÿϷ�(������ Ȯ�� ���ʿ�)
				<input type="radio" id="needChkYN_Y" name="needChkYN" value="Y" <%= CHKIIF(ioneRefas.FOneItem.FneedChkYN="Y", "checked", "") %> > ������ Ȯ�ο�û
				<% end if %>
			</td>
		</tr>
		<% end if %>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">�ù�����</td>
			<td>
            	<% if ioneRefas.FOneItem.IsRequireSongjangNO then %>
				<%
				Select Case ioneRefas.FOneItem.FsongjangRegGubun
					Case "U"
						Response.Write("��ü ����")
					Case "C"
						Response.Write("����������")
					Case "T"
						Response.Write("���� ����")
					Case Else
						Response.Write ioneRefas.FOneItem.FsongjangRegGubun
				End Select
				%>
		        <% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">������Է�</td>
			<td>
            	<% if ioneRefas.FOneItem.IsRequireSongjangNO then %>
				<%= ioneRefas.FOneItem.FsongjangRegUserID %>
		        <% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">���ÿ����</td>
			<td>
				<% drawSelectBoxDeliverCompany "songjangdiv",ioneRefas.FOneItem.FSongjangdiv %>
				<input type="text" class="text" name="songjangno" value="<%= ioneRefas.FOneItem.Fsongjangno %>" size="14" maxlength="14">
			</td>
		</tr>
		<% if (ioneas.FOneItem.Fdivcd="A100") then %>
			<tr bgcolor="#FFFFFF">
				<td bgcolor="<%= adminColor("tabletop") %>">�� �߰���ۺ�(����)</td>
				<td>
					<input type="text" class="text_ro" name="customeraddbeasongpay" value="<%= ioneas.FOneItem.Fcustomeraddbeasongpay %>" size="10" ReadOnly >
					&nbsp;
		    	    <select class="select" name="customeraddmethod" class="text" disabled>
			    	    <option value="">����
			    	    <option value="1" <% if (ioneas.FOneItem.Fcustomeraddmethod = "1") then %>selected<% end if %>>�ڽ�����
			    	    <option value="2" <% if (ioneas.FOneItem.Fcustomeraddmethod = "2") then %>selected<% end if %>>�ù�� ���δ�
			    	    <option value="9" <% if (ioneas.FOneItem.Fcustomeraddmethod = "9") then %>selected<% end if %>>ȯ�Ҿ׿��� ����
			    	    <option value="5" <% if (ioneas.FOneItem.Fcustomeraddmethod = "5") then %>selected<% end if %>>��Ÿ
		    	    </select>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td bgcolor="<%= adminColor("tabletop") %>">�� �߰���ۺ�(Ȯ��)</td>
				<input type="hidden" name="customerreceiveyn" value="<%= ioneas.FOneItem.Fcustomerreceiveyn %>">
				<td>
					<input type="text" class="text" name="customerrealbeasongpay" value="<%= ioneas.FOneItem.Fcustomerrealbeasongpay %>" size="10"> * �ڽ����� �� ��� ��ü���� Ȯ���� �ݾ�
				</td>
			</tr>
		<% end if %>
		<!-- ============================ (�±�ȯ���, ��ǰ���� �±�ȯ���) + (�±�ȯȸ�� ��ϵ� ���) + (�±�ȯȸ�� �Է�) -->

	<% else %>

		<!-- ============================ �̿� -->
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">�ù�����</td>
			<td>
            	<% if ioneas.FOneItem.IsRequireSongjangNO then %>
				<%
				Select Case ioneas.FOneItem.FsongjangRegGubun
					Case "U"
						Response.Write("��ü ����")
					Case "C"
						Response.Write("����������")
					Case "T"
						Response.Write("���� ����")
					Case Else
						Response.Write ioneas.FOneItem.FsongjangRegGubun
				End Select
				%>
		        <% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" height="30">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>">������Է�</td>
			<td>
            	<% if ioneas.FOneItem.IsRequireSongjangNO then %>
				<%= ioneas.FOneItem.FsongjangRegUserID %>
		        <% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%= adminColor("tabletop") %>" height="25">���ÿ����</td>
			<td>
				<% drawSelectBoxDeliverCompany "songjangdiv",ioneas.FOneItem.FSongjangdiv %>
				<input type="text" class="text" name="songjangno" value="<%= ioneas.FOneItem.Fsongjangno %>" size="14" maxlength="14">
			</td>
		</tr>
		<% if (IsRefASExist) then %>
			<tr bgcolor="#FFFFFF" height="30">
				<td width="130" height="120" bgcolor="<%= adminColor("tabletop") %>">ȸ�� ó������</td>
				<td>
					<%= nl2br(ioneRefas.FOneItem.Fcontents_finish) %>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF" height="30">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">�ù�����</td>
				<td>
					<% if ioneRefas.FOneItem.IsRequireSongjangNO then %>
					<%
					Select Case ioneRefas.FOneItem.FsongjangRegGubun
						Case "U"
							Response.Write("��ü ����")
						Case "C"
							Response.Write("����������")
						Case "T"
							Response.Write("���� ����")
						Case Else
							Response.Write ioneRefas.FOneItem.FsongjangRegGubun
					End Select
					%>
					<% end if %>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF" height="30">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">������Է�</td>
				<td>
					<% if ioneRefas.FOneItem.IsRequireSongjangNO then %>
					<%= ioneRefas.FOneItem.FsongjangRegUserID %>
					<% end if %>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF" height="30">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">ȸ�� �����</td>
				<td>
					<%= DeliverDivCd2Nm(ioneRefas.FOneItem.FSongjangdiv) %>
					&nbsp;
					<%= ioneRefas.FOneItem.Fsongjangno %>
				</td>
			</tr>
			<% if (ioneas.FOneItem.Fdivcd="A100") then %>
				<tr bgcolor="#FFFFFF">
					<td  height="30" bgcolor="<%= adminColor("tabletop") %>">�� �߰���ۺ�(����)</td>
					<td>
						<% if Not IsNull(ioneas.FOneItem.Fcustomeraddbeasongpay) then %>
							<%= FormatNumber(ioneas.FOneItem.Fcustomeraddbeasongpay, 0) %> ��
							(
							<% if (ioneas.FOneItem.Fcustomeraddmethod = "1") then %>
								�ڽ�����
							<% elseif (ioneas.FOneItem.Fcustomeraddmethod = "2") then %>
								�ù�� ���δ�
							<% elseif (ioneas.FOneItem.Fcustomeraddmethod = "9") then %>
								ȯ�Ҿ׿��� ����
							<% elseif (ioneas.FOneItem.Fcustomeraddmethod = "5") then %>
								��Ÿ
							<% end if %>
							)
						<% end if %>
					</td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td  height="30" bgcolor="<%= adminColor("tabletop") %>">�� �߰���ۺ�(Ȯ��)</td>
					<td>
						<% if Not IsNull(ioneas.FOneItem.Fcustomerrealbeasongpay) then %>
							<%= FormatNumber(ioneas.FOneItem.Fcustomerrealbeasongpay, 0) %> ��
						<% end if %>
						&nbsp;
						 * �ڽ����� �� ��� ��ü���� Ȯ���� �ݾ�
					</td>
				</tr>
			<% end if %>
			<tr bgcolor="#FFFFFF" height="30">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">ȸ�� ����</td>
				<td>
					<% if (receiveyn = "Y") then %>
						ȸ���Ϸ�
					<% elseif (receiveyn = "N") then %>
						<font color="blue">ȸ������</font>
					<% end if %>
				</td>
			</tr>
		<% end if %>
		<!-- ============================ �̿� -->

	<% end if %>

	</form>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<b>��ü �߰� ����</b>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="130" height="25" bgcolor="<%= adminColor("tabletop") %>">ȸ����ۺ�</td>
		<td>
			<%= FormatNumber(orefund.FOneItem.Frefunddeliverypay*-1, 0) %> ��
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="130" height="25" bgcolor="<%= adminColor("tabletop") %>">�߰������ۺ�</td>
		<td>
			<%= FormatNumber(ioneas.FOneItem.Fadd_upchejungsandeliverypay, 0) %> ��
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="130" height="25" bgcolor="<%= adminColor("tabletop") %>">�߰��������</td>
		<td>
			<%= ioneas.FOneItem.Fadd_upchejungsancause %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="130" height="25" bgcolor="<%= adminColor("tabletop") %>">�������ۺ�</td>
		<td>
			<b><%= FormatNumber((ioneas.FOneItem.Fadd_upchejungsandeliverypay + orefund.FOneItem.Frefunddeliverypay*-1), 0) %> ��</b>
		</td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="35" bgcolor="FFFFFF">
		<td colspan="15" align="center">

			<% if (IsRefASExist) and (receiveonly = "Y") then %>

				<% if ioneRefas.FOneItem.Fcurrstate="B007" then %>

				<% else %>
					<input type="button" class="button" value="ȸ���Ϸ�ó��" onclick="javascript:SaveFin(frm);">
				<% end if %>

			<% else %>

				<% if ioneas.FOneItem.Fcurrstate="B007" then %>

				<% else %>
				    <% if ((ioneas.FOneItem.Fdivcd = "A000") or (ioneas.FOneItem.Fdivcd = "A100")) and (IsRefASExist) then %>
					<input type="button" class="button" value="���ó��" onclick="javascript:SaveFin(frm);">
					<% else %>
					<input type="button" class="button" value="�Ϸ�ó��" onclick="javascript:SaveFin(frm);">
					<% end if %>
				<% end if %>

			<% end if %>

			<input type="button" class="button" value="��Ϻ���" onClick="location.href='/designer/jumunmaster/upchecslist.asp';">
		</td>
	</tr>
</table>

<!-- ǥ �ϴܹ� ��-->

<%
set ioneas = Nothing
set ioneasDetail = Nothing
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
