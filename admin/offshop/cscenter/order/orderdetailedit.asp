<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ������
' Hieditor : 2011.03.09 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim i ,detailidx, masteridx ,ojumunDetail
	detailidx = requestCheckVar(request("detailidx"),10)

set ojumunDetail = new COrder
	ojumunDetail.frectdetailidx = detailidx
	ojumunDetail.fSearchOneJumunDetail()

	if ojumunDetail.ftotalcount > 0 then
		masteridx = ojumunDetail.FOneItem.fmasteridx
	else
		response.write "<script language='javascript'>"
		response.write "	alert('�ֹ������� ���̺� ���� �����ϴ�');"
		response.write "	self.close();"
		response.write "</script>"
		response.end
	end if

dim ojumun
set ojumun = new COrder
	if (masteridx <> "") then
	    ojumun.FRectmasteridx = masteridx
	    ojumun.fQuickSearchOrderMaster()
	end if

%>

<script language='javascript'>

window.resizeTo(600,600);

var oldConfirmDate = "";
var oldBeasongDate = "";

function CheckConfirmDate(comp){
    if (comp.value==""){
        oldConfirmDate = comp.form.upcheconfirmdate.value;
        oldBeasongDate = comp.form.beasongdate.value;
        comp.form.upcheconfirmdate.value = "";
    }else{
        if (oldConfirmDate!=""){
            comp.form.upcheconfirmdate.value = oldConfirmDate;
        }

        if (oldBeasongDate!=""){
            comp.form.beasongdate.value = oldBeasongDate;
        }
    }
}

//����
function EditDetail(detailidx,mode,comp){
    var frm = document.frm;

	if(mode=="currstate"){
		<% if ojumun.FOneItem.FIpkumdiv<4 then %>
		    alert('�����Ϸ� �̻� ���º��氡��.');
		    return;
		<% end if %>
    }else if(mode=="songjangdiv"){
        if (frm.songjangdiv.value.length<1){
            alert('�ù�縦 �����ϼ���.');
			frm.songjangdiv.focus();
			return;
        }

        if (!IsDigit(frm.songjangno.value)){
			alert('������ȣ�� ���ڴ� �����մϴ�.');
			frm.songjangdiv.focus();
			return;
		}
	}else if (mode=="itemno"){

	}else{
		return;
	}

	frm.mode.value=mode;

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

</script>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#CCCCCC" class="a">
<form name="frm" method="post" action="/admin/offshop/cscenter/order/order_process.asp">
<input type="hidden" name="detailidx" value="<%= ojumunDetail.FOneItem.Fdetailidx %>">
<input type="hidden" name="masteridx" value="<%= ojumunDetail.FOneItem.Fmasteridx %>">
<input type="hidden" name="mode" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ֹ����������� ����</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">Idx.</td>
	<td><%= ojumunDetail.FOneItem.Fdetailidx %></td>
	<td width="120"></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
	<td>
		<%= ojumunDetail.FOneItem.fitemgubun%>-<%=CHKIIF(ojumunDetail.FOneItem.fitemid>=1000000,Format00(8,ojumunDetail.FOneItem.fitemid),Format00(6,ojumunDetail.FOneItem.fitemid))%>-<%=ojumunDetail.FOneItem.fitemoption %>
	</td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">�ɼ��ڵ�</td>
	<td><%= ojumunDetail.FOneItem.Fitemoption %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
	<td><%= ojumunDetail.FOneItem.Fitemname %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">�ɼǸ�</td>
	<td><%= ojumunDetail.FOneItem.Fitemoptionname %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">�귣�� ID</td>
	<td><%= ojumunDetail.FOneItem.Fmakerid %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">�ǸŰ�</td>
	<td>
		<%= ojumunDetail.FOneItem.fsellprice %>
		<% if (ojumunDetail.FOneItem.fsellprice<>ojumunDetail.FOneItem.FCurrsellcash) then %>
			(���ǸŰ�:<%= ojumunDetail.FOneItem.FCurrsellcash %>)
		<% end if %>
	</td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">���ż���</td>
	<td><input type="text" class="text" name="itemno" value="<%= ojumunDetail.FOneItem.Fitemno %>" size="3" maxlength="9">��</td>
	<td>
	    <% if (C_ADMIN_AUTH) then %>
	    	<input type="button" class="button" value="��������" <%= CHKIIF(ojumun.FOneItem.FIpkumdiv>6,"disabled","") %> onclick="javascript:EditDetail(<%= ojumunDetail.FOneItem.Fdetailidx %>,'itemno',frm.itemno)">
	    <% end if %>
	</td>
</tr>
<% if ojumunDetail.FOneItem.Fisupchebeasong="Y" then %>
<tr bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">�ֹ�����</td>
	<td>
		<select class="select" name="currstate" class="text">
    		<option value="" <% if ojumunDetail.FOneItem.Fcurrstate="" then response.write "selected" %> onChange="CheckConfirmDate(this);">��Ȯ��
    		<option value="2" <% if ojumunDetail.FOneItem.Fcurrstate="2" then response.write "selected" %> onChange="CheckConfirmDate(this);">��ü�뺸
    		<option value="3" <% if ojumunDetail.FOneItem.Fcurrstate="3" then response.write "selected" %> onChange="CheckConfirmDate(this);">�ֹ�Ȯ��
    		<option value="7" <% if ojumunDetail.FOneItem.Fcurrstate="7" then response.write "selected" %> onChange="CheckConfirmDate(this);">���Ϸ�
		</select>
	</td>
	<td>
		<input type="button" class="button" value="Ȯ�λ��¼���" onclick="javascript:EditDetail(<%= ojumunDetail.FOneItem.Fdetailidx %>,'currstate',frm.currstate)">
	</td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">�ֹ��뺸��<br>(������)</td>
	<td><%= ojumun.FOneItem.FBaljuDate %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">��üȮ����</td>
	<td><input type="text" class="text" name="upcheconfirmdate" value="<%= ojumunDetail.FOneItem.Fupcheconfirmdate %>" size="21" maxlength="19"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
	<td><input type="text" class="text" name="beasongdate" value="<%= ojumunDetail.FOneItem.Fbeasongdate %>" size="21" maxlength="19"></td>
	<td></td>
</tr>

<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">�ù�����</td>
	<td>
		<% drawSelectBoxDeliverCompany "songjangdiv",ojumunDetail.FOneItem.Fsongjangdiv %>
		<input type="text" class="text" name="songjangno" value="<%= ojumunDetail.FOneItem.Fsongjangno %>" size="20" maxlength="20">
	</td>
	<td><input type="button" class="button" value="�ù���������" onclick="javascript:EditDetail(<%= ojumunDetail.FOneItem.Fdetailidx %>,'songjangdiv',frm.songjangdiv)"></td>
</tr>

<% else %>

<tr bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">�ֹ��뺸��<br>(������)</td>
	<td><%= ojumun.FOneItem.FBaljuDate %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
	<td><input type="text" class="text" name="beasongdate" value="<%= ojumunDetail.FOneItem.Fbeasongdate %>" size="21" maxlength="19"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("tabletop") %>">�ù�����</td>
	<td>
		<% drawSelectBoxDeliverCompany "songjangdiv",ojumunDetail.FOneItem.Fsongjangdiv %>
		<input type="text" class="text" name="songjangno" value="<%= ojumunDetail.FOneItem.Fsongjangno %>" size="20" maxlength="20">
	</td>
	<td></td>
</tr>
<% end if %>
</form>
</table>

<%
set ojumun       = Nothing
set ojumunDetail = Nothing
%>
<!-- #include virtual="/admin/offshop/cscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->