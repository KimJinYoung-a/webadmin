<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������
' History : �̻� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
dim i
dim idx, orderserial
idx = requestCheckVar(request("idx"),10)

dim ojumunDetail
set ojumunDetail = new CJumunMaster
ojumunDetail.SearchOneJumunDetail idx

if (ojumunDetail.Fresultcount<1) then
    ojumunDetail.FRectOldJumun = "on"
    ojumunDetail.SearchOneJumunDetail idx
end if

orderserial = ojumunDetail.FJumunDetail.FOrderSerial

dim ojumun
set ojumun = new COrderMaster

if (orderserial <> "") then
    ojumun.FRectOrderSerial = orderserial
    ojumun.QuickSearchOrderMaster
end if


if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster
end if

if ojumun.FResultCount<1 then
	response.write "�ش�Ǵ� �ֹ����� �����ϴ�."
	dbget.close() : response.end
end if

%>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
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

function EditDetail(detailidx,mode,comp) {
<% if ojumunDetail.FRectOldJumun = "on" or ojumun.FRectOldOrder = "on" then %>
    alert('���ų��� �����Ұ�.');
    return;
<% end if %>
    var frm = document.frm;

	if (mode=="buycash") {
		if (!IsDigit(comp.value)) {
			alert("���԰��� ���ڸ� �����մϴ�.");
			comp.focus();
			return;
		}
	}else if (mode=="reducedPrice") {
		if (!IsDigit(comp.value)) {
			alert("�������� ���ڸ� �����մϴ�.");
			comp.focus();
			return;
		}
	}else if (mode=="itemcost") {
		if (!IsDigit(comp.value)) {
			alert("�������� ���ڸ� �����մϴ�.");
			comp.focus();
			return;
		}
    }else if (mode=="itemcostCouponNotApplied") {
		if (!IsDigit(comp.value)) {
			alert("�ǸŰ��� ���ڸ� �����մϴ�.");
			comp.focus();
			return;
		}
	}else if(mode=="isupchebeasong") {
	    if (frm.isupchebeasong.value=="Y") {
	        if (frm.omwdiv.value!="U") {
	            alert("���Ա��а� ��۱����� ��ġ���� �ʽ��ϴ�.");
	            return;
	        }
	    }else{
	        if (frm.omwdiv.value=="U") {
	            alert("���Ա��а� ��۱����� ��ġ���� �ʽ��ϴ�.");
	            return;
	        }
	    }

        if (frm.omwdiv.value=="U") {
            if ((frm.odlvType.value=="1")||(frm.odlvType.value=="4")) {
                alert("���Ա��а� ��۱����� ��ġ���� �ʽ��ϴ�.");
	            return;
            }
        }else{
            if ((frm.odlvType.value!="1")&&(frm.odlvType.value!="4")) {
                alert("���Ա��а� ��۱����� ��ġ���� �ʽ��ϴ�.");
	            return;
            }
        }

	}else if(mode=="songjangdiv") {
		// �߸��� �����ȣ ���� �� EMS�ù�� ������ ���Ե� �����ȣ.
	    if (frm.songjangdiv.value.length<1) {
            alert("�ù�縦 �����ϼ���.\n\n�����ȣ�� ���� ��� ��Ÿ�� �Է��ϼ���.");
			frm.songjangdiv.focus();
			return;
        }

        if (frm.songjangno.value == '') {
			alert("������ȣ�� �Է��ϼ���.\n\n�����ȣ�� ���� ��� ��Ÿ�� �Է��ϼ���.");
			frm.songjangno.focus();
			return;
		}

		if (frm.applyallitem) {
			if (frm.applyallitem.checked == true) {
				if (confirm("�ش� ��ü [��ü��ǰ] �� ���� �����ȣ�� �Է��մϴ�.\n\n�����Ͻðڽ��ϱ�?") != true) {
					return;
				}
			}
		}

	}else if(mode=="currstate") {

	<% if ojumun.FOneItem.FIpkumdiv<4 then %>
	    alert("�����Ϸ� �̻� ���º��氡��.");
	    return;
	<% end if %>

		if (frm.currstate.value == '7') {
			// �߸��� �����ȣ ���� �� EMS�ù�� ������ ���Ե� �����ȣ.
			if (frm.songjangdiv.value.length<1) {
				alert("�ù�縦 �����ϼ���.\n\n�����ȣ�� ���� ��� ��Ÿ�� �Է��ϼ���.");
				frm.songjangdiv.focus();
				return;
			}

			if (frm.songjangno.value == '') {
				alert("������ȣ�� �Է��ϼ���.\n\n�����ȣ�� ���� ��� ��Ÿ�� �Է��ϼ���.");
				frm.songjangno.focus();
				return;
			}

			/*
			if (frm.applyallitem) {
				if (frm.applyallitem.checked == true) {
					alert('�Ѱ��� ��ǰ�� ���ؼ��� �Է°����մϴ�.');
					frm.applyallitem.checked = false;
				}
			}
			*/
		}

	}else if (mode=="requiredetail") {

	}else if (mode=="itemno") {

	}else if (mode=="vatinclude") {
		if (comp.value == "") {
			alert("���������� �����ϼ���.");
			return;
		}
    }else if (mode=="recalcmaster") {

	}else if (mode=="jungsan") {

    }else if (mode=="10x10logistics") {

	}else if (mode=="balju") {

    }else if (mode=="updmastercoupon") {

    }else if (mode=="ipkumdate") {

    }else if (mode=="additemid") {

	}else{
		return;
	}

	frm.mode.value=mode;

	if (confirm("���� �Ͻðڽ��ϱ�?")) {
		frm.submit();
	}
}

function PopCurrentItemStock(itemid, itemoption) {
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?itemgubun=10&itemid=" + itemid + "&itemoption=" + itemoption,"PopCurrentItemStock","width=1000,height=600,resizable=yes,scrollbars=yes")
	popwin.focus();
}

</script>

<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#CCCCCC" class="a">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ֹ����������� ����</b>
		</td>
	</tr>

	<form name="frm" method="post" action="orderedit_process.asp">
	<input type="hidden" name="detailidx" value="<%= ojumunDetail.FJumunDetail.Fdetailidx %>">
	<input type="hidden" name="orderserial" value="<%= ojumunDetail.FJumunDetail.FOrderSerial %>">
	<input type="hidden" name="presongjangno" value="<%= ojumunDetail.FJumunDetail.Fsongjangno %>">
	<input type="hidden" name="presongjangdiv" value="<%= ojumunDetail.FJumunDetail.Fsongjangdiv %>">
	<input type="hidden" name="mode" value="">
	<tr bgcolor="#FFFFFF" height="25">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">Idx.</td>
		<td><%= ojumunDetail.FJumunDetail.Fdetailidx %></td>
		<td width="120"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
		<td>
			<%= ojumunDetail.FJumunDetail.Fitemid %>
			&nbsp;
			<input type="button" class="button" value="��ǰ�������Ȳ" onClick="PopCurrentItemStock('<%= ojumunDetail.FJumunDetail.Fitemid %>', '<%= ojumunDetail.FJumunDetail.Fitemoption %>')">
		</td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�ɼ��ڵ�</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemoption %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemname %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�ɼǸ�</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemoptionname %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�귣�� ID</td>
		<td><%= ojumunDetail.FJumunDetail.Fmakerid %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">���ֻ���</td>
		<td>

		</td>
		<td>
		    <input type="button" class="button" value="�����Ϸ���ȯ" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'balju')">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�ǸŰ�</td>
		<td>
            <input type="text" class="text" name="itemcostCouponNotApplied" value="<%= ojumunDetail.FJumunDetail.FitemcostCouponNotApplied %>" size="7" maxlength="9">
			<% if (ojumunDetail.FJumunDetail.Fitemcost<>ojumunDetail.FJumunDetail.FCurrsellcash) then %>
				(���ǸŰ�:<%= ojumunDetail.FJumunDetail.FCurrsellcash %>)
			<% end if %>
		</td>
		<td>
            <% if C_ADMIN_AUTH or C_CSPowerUser or C_MngPart then %>
		    <input type="button" class="button" value="�ǸŰ�����" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'itemcostCouponNotApplied',frm.itemcostCouponNotApplied)">
		    <% end if %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ������</td>
		<td>
            <input type="text" class="text" name="itemcost" value="<%= ojumunDetail.FJumunDetail.Fitemcost %>" size="7" maxlength="9">
		</td>
		<td>
            <% if C_ADMIN_AUTH or C_CSPowerUser or C_MngPart then %>
		    <input type="button" class="button" value="����������" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'itemcost',frm.itemcost)">
		    <% end if %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">���ʽ�������</td>
		<td>
            <input type="text" class="text" name="reducedPrice" value="<%= ojumunDetail.FJumunDetail.FreducedPrice %>" size="7" maxlength="9">
			<% if (ojumunDetail.FJumunDetail.Fitemcost<>ojumunDetail.FJumunDetail.FCurrsellcash) then %>
				(���ǸŰ�:<%= ojumunDetail.FJumunDetail.FCurrsellcash %>)
			<% end if %>
            * ���ξ� �ٹ����ٺδ�
		</td>
		<td>
            <% if C_ADMIN_AUTH or C_CSPowerUser or C_MngPart then %>
		    <input type="button" class="button" value="����������" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'reducedPrice',frm.reducedPrice)">
		    <% end if %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">���԰�</td>
		<td>
			<input type="text" class="text" name="buycash" value="<%= ojumunDetail.FJumunDetail.Fbuycash %>" size="7" maxlength="9">
			<% if ojumunDetail.FJumunDetail.Fitemcost<>0 then %>
			(<%= 100-CLng(ojumunDetail.FJumunDetail.Fbuycash/ojumunDetail.FJumunDetail.Fitemcost*10000/100) %> %)
			<% end if %>
			<% if (ojumunDetail.FJumunDetail.Fbuycash<>ojumunDetail.FJumunDetail.FCurrbuycash) then %>
				(�����԰�:<%= ojumunDetail.FJumunDetail.FCurrbuycash %>)
			<% end if %>
            * ������ �Ǵ� ȸ���� ���� ����
        </td>
		<td>
		    <% if C_ADMIN_AUTH or C_CSPowerUser or C_MngPart then %>
		    <input type="button" class="button" value="���԰�����" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'buycash',frm.buycash)">
		    <% end if %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td>
			<select class="select" name="vatinclude" class="text">
        		<option value="">����</option>
        		<option value="Y" <% if ojumunDetail.FJumunDetail.Fvatinclude="Y" then response.write "selected" %> >����</option>
        		<option value="N" <% if ojumunDetail.FJumunDetail.Fvatinclude="N" then response.write "selected" %> >�鼼</option>
    		</select>
			* ���곻���� �ִ� ��� �����Ұ�!
    	</td>
		<td>
			<% if (C_ADMIN_AUTH) then %>
			<input type="button" class="button" value="�������м���" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'vatinclude',frm.vatinclude)">
			<% end if %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">���ϸ���</td>
		<td><%= ojumunDetail.FJumunDetail.Fmileage %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">���ż���</td>
		<td><input type="text" class="text" name="itemno" value="<%= ojumunDetail.FJumunDetail.Fitemno %>" size="3" maxlength="9">��</td>
		<td>
		    <% if (C_ADMIN_AUTH) then %>
		    <input type="button" class="button" value="��������" <%= CHKIIF(ojumun.FOneItem.FIpkumdiv>6,"disabled","") %> onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'itemno',frm.itemno)">
		    <% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">��۱���</td>
		<td>
			<select class="select" name="isupchebeasong">
	    		<option value="Y" <% if ojumunDetail.FJumunDetail.Fisupchebeasong="Y" then response.write "selected" %> >��ü���
	    		<option value="N" <% if ojumunDetail.FJumunDetail.Fisupchebeasong="N" then response.write "selected" %> >��ü���
			</select>

	        <select class="select" name="omwdiv">
	            <option value="M" <%= chkIIF(ojumunDetail.FJumunDetail.Fomwdiv="M","selected","") %> >����
	            <option value="W" <%= chkIIF(ojumunDetail.FJumunDetail.Fomwdiv="W","selected","") %> >��Ź
	            <option value="U" <%= chkIIF(ojumunDetail.FJumunDetail.Fomwdiv="U","selected","") %> >��ü
	        </select>

	        <select class="select" name="odlvType">
	            <option value="1" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="1","selected","") %> >��ü���
	            <option value="2" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="2","selected","") %> >��ü���
	            <option value="4" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="4","selected","") %> >��ü����
	            <option value="5" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="5","selected","") %> >��ü����
	            <option value="7" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="7","selected","") %> >��ü����
	            <option value="9" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="9","selected","") %> >��ü����
	            <option value="6" <%= chkIIF(ojumunDetail.FJumunDetail.FodlvType="6","selected","") %> >�������
	        </select>
		</td>
		<td>
		    <% if (C_ADMIN_AUTH) then %>
		    <input type="button" class="button" value="��۱��м���" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'isupchebeasong',frm.isupchebeasong)" >
		    <% end if %>
		</td>
	</tr>

	<% if ojumunDetail.FJumunDetail.Fisupchebeasong="Y" then %>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">�ֹ�����</td>
		<td>
			<select class="select" name="currstate" class="text">
                <% if (ojumunDetail.FJumunDetail.Fcurrstate = 7) then %>
                <option value="7to3" <% if ojumunDetail.FJumunDetail.Fcurrstate="7" then response.write "selected" %> onChange="CheckConfirmDate(this);">������� ��ȯ</option>
                <% else %>
        		<option value="" <% if ojumunDetail.FJumunDetail.Fcurrstate="" then response.write "selected" %> onChange="CheckConfirmDate(this);">��Ȯ��
        		<option value="2" <% if ojumunDetail.FJumunDetail.Fcurrstate="2" then response.write "selected" %> onChange="CheckConfirmDate(this);">��ü�뺸
        		<option value="3" <% if ojumunDetail.FJumunDetail.Fcurrstate="3" then response.write "selected" %> onChange="CheckConfirmDate(this);">�ֹ�Ȯ��
        		<option value="7" <% if ojumunDetail.FJumunDetail.Fcurrstate="7" then response.write "selected" %> onChange="CheckConfirmDate(this);">���Ϸ�
                <% end if %>
    		</select>
			* ���곻���� �ִ� ��� �����Ұ�!
    	</td>
		<td><input type="button" class="button" value="Ȯ�λ��¼���" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'currstate',frm.currstate)"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">�ֹ��뺸��<br>(������)</td>
		<td><%= ojumun.FOneItem.FBaljuDate %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">��üȮ����</td>
		<td><input type="text" class="text" name="upcheconfirmdate" value="<%= ojumunDetail.FJumunDetail.Fupcheconfirmdate %>" size="21" maxlength="19"></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
		<td><input type="text" class="text" name="beasongdate" value="<%= ojumunDetail.FJumunDetail.Fbeasongdate %>" size="21" maxlength="19"></td>
		<td></td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">�ù�����</td>
		<td>
			<% drawSelectBoxDeliverCompany "songjangdiv",ojumunDetail.FJumunDetail.Fsongjangdiv %>
			<input type="text" class="text" name="songjangno" value="<%= ojumunDetail.FJumunDetail.Fsongjangno %>" size="17" maxlength="20">
			<br /><input type="checkbox" name="applyallitem" value="Y"> ����ǰ����
		</td>
		<td><input type="button" class="button" value="�ù���������" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'songjangdiv',frm.songjangdiv)"></td>
	</tr>

	<% else %>

	<% if (C_ADMIN_AUTH or session("ssBctId") = "hasora" or session("ssBctId") = "boyishP" or session("ssBctId") = "oesesang52" or session("ssBctId") = "rabbit1693") then %>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">�ֹ�����</td>
		<td>
			<select class="select" name="currstate" class="text">
                <% if (ojumunDetail.FJumunDetail.Fcurrstate = 7) then %>
                <option value="7to3" <% if ojumunDetail.FJumunDetail.Fcurrstate="7" then response.write "selected" %> onChange="CheckConfirmDate(this);">������� ��ȯ</option>
                <% else %>
				<option value="" <% if ojumunDetail.FJumunDetail.Fcurrstate="" then response.write "selected" %> onChange="CheckConfirmDate(this);">��Ȯ��</option>
        		<option value="2" <% if ojumunDetail.FJumunDetail.Fcurrstate="2" then response.write "selected" %> onChange="CheckConfirmDate(this);">�����뺸</option>
        		<option value="3" <% if ojumunDetail.FJumunDetail.Fcurrstate="3" then response.write "selected" %> onChange="CheckConfirmDate(this);">�ֹ�Ȯ��</option>
        		<option value="7" onChange="CheckConfirmDate(this);">���Ϸ�</option>
                <% end if %>
    		</select>
			* skyer9,tozzinet,hasora,boyishP,oesesang52, rabbit1693 only
			<br>* �߰��� ����� �ʿ�
			<br>* ���곻���� �ִ� ��� �����Ұ�!
    	</td>
		<td><input type="button" class="button" value="Ȯ�λ��¼���" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'currstate',frm.currstate)"></td>
	</tr>
    <% elseif (C_ADMIN_AUTH or C_CSPowerUser) then %>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">�ֹ�����</td>
		<td>
            <% if (ojumunDetail.FJumunDetail.Fcurrstate = 7) then %>
            <select class="select" name="currstate" class="text">
        		<option value="7to3" <% if ojumunDetail.FJumunDetail.Fcurrstate="7" then response.write "selected" %> onChange="CheckConfirmDate(this);">������� ��ȯ</option>
    		</select>
            <% else %>
            ����ǰ���� ��밡��
            <input type="hidden" name="currstate" value="<%= ojumunDetail.FJumunDetail.Fcurrstate %>">
            <% end if %>
			* CS������ ����
			<br>* ���곻���� �ִ� ��� �����Ұ�!
    	</td>
		<td>
            <% if (ojumunDetail.FJumunDetail.Fcurrstate = 7) then %>
            <input type="button" class="button" value="Ȯ�λ��¼���" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'currstate',frm.currstate)">
            <% end if %>
        </td>
	</tr>
	<% else %>
	<input type="hidden" name="currstate" value="<%= ojumunDetail.FJumunDetail.Fcurrstate %>">
	<% end if %>

	<tr bgcolor="#FFFFFF" >
		<td width="80" bgcolor="<%= adminColor("tabletop") %>">�ֹ��뺸��<br>(������)</td>
		<td><%= ojumun.FOneItem.FBaljuDate %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
		<td><input type="text" class="text" name="beasongdate" value="<%= ojumunDetail.FJumunDetail.Fbeasongdate %>" size="21" maxlength="19"></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">�ù�����</td>
		<td>
			<% drawSelectBoxDeliverCompany "songjangdiv",ojumunDetail.FJumunDetail.Fsongjangdiv %>
			<input type="text" class="text" name="songjangno" value="<%= ojumunDetail.FJumunDetail.Fsongjangno %>" size="20" maxlength="20">
			<br /><input type="checkbox" name="applyallitem" value="Y"> ����ǰ����
		</td>
		<td><input type="button" class="button" value="�ù���������" onclick="javascript:EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'songjangdiv',frm.songjangdiv)"></td>
	</tr>

	<% end if %>

	<% if ojumunDetail.FJumunDetail.Foitemdiv="06" then %>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">�ֹ����۹���</td>
		<td>
		    <% if ojumunDetail.FJumunDetail.FItemNo=1 then %>
		    <textarea readonly class="textarea" name="requiredetail" cols="40" rows="3"><%= ojumunDetail.FJumunDetail.Frequiredetail %></textarea>
		    <% else %>
		    <% for i=0 to ojumunDetail.FJumunDetail.FItemNo-1 %>
		    <textarea readonly class="textarea" name="requiredetail<%=i%>" cols="40" rows="3"><%= splitValue(ojumunDetail.FJumunDetail.Frequiredetail,CAddDetailSpliter,i) %></textarea>
		    <% next %>
		    <% end if %>
		</td>
		<td>
		    <input type="button" class="button" value="�ֹ����۹�������" onclick="EditRequireDetail('<%= orderserial %>','<%= ojumunDetail.FJumunDetail.Fdetailidx %>')">
		</td>
	</tr>
	<% end if %>

    <tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">������������</td>
		<td>
            ������������ ����
		</td>
		<td>
		    <input type="button" class="button" value="����" onclick="EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'updmastercoupon',null)">
		</td>
	</tr>

    <% if C_ADMIN_AUTH then %>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">�ֹ�������</td>
		<td>
		</td>
		<td>
		    <input type="button" class="button" value="����[������]" onclick="EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'recalcmaster',null)">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td>
		</td>
		<td>
		    <input type="button" class="button" value="üũ[������]" onclick="EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'jungsan',null)">
		</td>
	</tr>
    <tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">��ۺ�</td>
		<td>
            10x10logistics ��ۺ��߰�
		</td>
		<td>
		    <input type="button" class="button" value="�߰�[������]" onclick="EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'10x10logistics',null)">
		</td>
	</tr>
    <tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">�Ա���</td>
		<td>
            <input type="text" class="text" name="ipkumdate" value="<% ''CHKIIF(ojumun.FOneItem.Fipkumdate="", "", FormatDate(ojumun.FOneItem.Fipkumdate, "0000.00.00 00:00:00")) %>" size="21" maxlength="30">
		</td>
		<td>
		    <input type="button" class="button" value="����[������]" onclick="EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'ipkumdate',null)" disabled>
		</td>
	</tr>
    <tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�߰�</td>
		<td>
            ���ֹ� : <input type="text" class="text" name="orgorderserial" value="" size="10" maxlength="30">
            <br />
            ������ : <input type="text" class="text" name="orgdetailidx" value="" size="10" maxlength="30">
		</td>
		<td>
		    <input type="button" class="button" value="����[������]" onclick="EditDetail(<%= ojumunDetail.FJumunDetail.Fdetailidx %>,'additemid',null)">
		</td>
	</tr>
    <% end if %>
	</form>
</table>


<%
set ojumun       = Nothing
set ojumunDetail = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
