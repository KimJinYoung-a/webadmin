<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : �������� ���
' Hieditor : 2011.02.25 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim ojumun ,i
set ojumun = new cupchebeasong_list
	ojumun.FRectDesignerID = session("ssBctId")
	ojumun.fDesignerDateBaljuList()

%>

<script language='javascript'>

//��ü����
function switchCheckBox(comp){
    var frm = comp.form;

	if(frm.detailidx.length>1){
		for(i=0;i<frm.detailidx.length;i++){
			frm.detailidx[i].checked = comp.checked;
			AnCheckClick(frm.detailidx[i]);
		}
	}else{
		frm.detailidx.checked = comp.checked;
		AnCheckClick(frm.detailidx);
	}
}

//���� �ֹ� Ȯ��
function CheckNBaljusu(){
	var frm = document.frmbalju;
	var pass = false;

    if(frm.detailidx.length>1){
    	for (var i=0;i<frm.detailidx.length;i++){
    	    pass = (pass||frm.detailidx[i].checked);
    	}
    }else{
        pass = frm.detailidx.checked;
    }

	if (!pass) {
		alert("���� �ֹ��� �����ϴ�.");
		return;
	}

	var ret = confirm("���� �ֹ��� Ȯ�� �Ͻðڽ��ϱ�?");

	if (ret){
 		frm.action="/common/offshop/beasong/upche_selectbaljulist.asp";
		frm.submit();

	}
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left" bgcolor="#FFFFFF">
		<input type="radio" name="" value="" checked >��ۿ�û ����Ʈ
		<!-- <input type="radio" name="" value="">��û���� �ֹ�����Ʈ(�ֹ����� ����) -->
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>

<br>
�� ��ۿ�û���� ��, [���ÿ�ûȮ��]�� Ŭ���Ͻø�, ���ּ� ����� �����մϴ�.
<br>���ּ� ������� ���Ͻ� ���, [�������̹�۸���Ʈ]�� �̿��Ͻñ� �ٶ��ϴ�.
<br>(��ûȮ���� �ϼž� ������� Ȯ���� �����մϴ�.)

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		<input type="button" class="button" value="���ÿ�ûȮ��" onclick="CheckNBaljusu()">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmbalju" method="post">
<tr bgcolor="FFFFFF">
	<td height="25" colspan="15">
		�˻���� : <b><% = ojumun.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
	<td>IDX</td>
	<td>�����ֹ���ȣ</td>
	<td>������</td>
	<td>��ǰ�ڵ�</td>
	<td>��ǰ��<font color="blue">&nbsp;[�ɼ�]</font></td>
	<td>���ް�</td>
	<td>�ǸŰ�</td>
	<td>����</td>
	<td>��ۿ�û��</td>
	<td>��������<!--�ֹ��뺸��--></td>
	<td>�����</td>
</tr>
<% if ojumun.ftotalcount > 0 then %>
<% for i=0 to ojumun.ftotalcount-1 %>
<tr align="center" class="a" bgcolor="#FFFFFF">
	<td>
	    <!-- detail Index -->
		<input type="checkbox" name="detailidx"  onClick="AnCheckClick(this);" value="<% =ojumun.fitemlist(i).fdetailidx %>">
	</td>
	<td><%= ojumun.fitemlist(i).fdetailidx %></td>
	<td><%= ojumun.fitemlist(i).forderno %></td>
	<td><%= ojumun.fitemlist(i).FReqname %></td>
	<td><%= ojumun.fitemlist(i).fitemgubun %>-<%= CHKIIF(ojumun.fitemlist(i).FitemID>=1000000,Format00(8,ojumun.fitemlist(i).FitemID),Format00(6,ojumun.fitemlist(i).FitemID)) %>-<%= ojumun.fitemlist(i).fitemoption %></td>
	<td align="left">
		<%= ojumun.fitemlist(i).FItemname %>
		<% if (ojumun.fitemlist(i).fitemoptionname<>"") then %>
		<font color="blue">[<%= ojumun.fitemlist(i).fitemoptionname %>]</font>
		<% end if %>
	</td>
	<td><%= FormatNumber(ojumun.fitemlist(i).fsuplyprice,0) %></td>
	<td><%= FormatNumber(ojumun.fitemlist(i).fsellprice,0) %></td>
	<td><%= ojumun.fitemlist(i).FItemno %></td>
	<td><acronym title="<%= ojumun.fitemlist(i).Fregdate %>"><%= left(ojumun.fitemlist(i).Fregdate,10) %></acronym></td>
	<td><acronym title="<%= ojumun.fitemlist(i).Fbaljudate %>"><%= left(ojumun.fitemlist(i).Fbaljudate,10) %></acronym></td>
    <td>
        <% if IsNULL(ojumun.fitemlist(i).Fbaljudate) then %>
        	D+0
        <% elseif datediff("d",(left(ojumun.fitemlist(i).Fbaljudate,10)) , (left(now,10)) )>2 then %>
        	<font color="red"><b>D+<%= datediff("d",(left(ojumun.fitemlist(i).Fbaljudate,10)) , (left(now,10)) ) %></b></font>
        <% else %>
        	D+<%= datediff("d",(left(ojumun.fitemlist(i).Fbaljudate,10)) , (left(now,10)) ) %>
        <% end if %>
    </td>
</tr>

<% next %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
	<% end if %>


    </form>
</table>


<%
set ojumun = Nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->