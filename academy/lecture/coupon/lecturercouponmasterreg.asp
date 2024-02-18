<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� ����
' History : 2010.10.11 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/lecturer/lecturercouponcls.asp" -->
<%
dim lecturercouponidx ,oitemcouponmaster ,IsEditMode, IsExpiredCoupon
	lecturercouponidx = requestCheckVar(request("lecturercouponidx"),9)
	if lecturercouponidx="" then lecturercouponidx=0

set oitemcouponmaster = new ClecturerCouponMaster
	oitemcouponmaster.FRectlecturerCouponIdx = lecturercouponidx
	oitemcouponmaster.GetOnelecturerCouponMaster()

IsEditMode = (CStr(lecturercouponidx)<>"0")
%>

<script language='javascript'>

function OpenCouponMaster(){
	frmcoupon.mode.value="opencoupon";

	if (confirm('������ ���� �Ͻðڽ��ϱ�?')){
		frmcoupon.submit();
	}
}

function reserveCouponMaster(){
	frmcoupon.mode.value="reservecoupon";

	if (confirm('���������� ���� �Ͻðڽ��ϱ�?')){
		frmcoupon.submit();
	}

}

var alertCnt = 0;
function AlertMarginChange(){
	if (alertCnt==0){
		alert('���� ������ �����Ͻø� ����ǰ ��ü�� ���� �˴ϴ�.');
		alertCnt++;
	}
}

function CloseCouponMaster(){
	frmcoupon.mode.value="closecoupon";

	if (confirm('!! ���� ����� ������ ���� ���� �˴ϴ�.\n\n������ ���� ���� �Ͻðڽ��ϱ�?')){
		frmcoupon.submit();
	}
}

function fninput(v){

	var ele = document.getElementById('marginlayer');

	if (v==20){
		ele.style.display="";
	}else {
		ele.style.display="none";
	}
}

function SaveCouponMaster(frm, isEditMode){
	if (frmcoupon.lecturercouponname.value.length<2){
		alert('�������� �Է��� �ּ���.');
		frmcoupon.lecturercouponname.focus();
		return;
	}

    if ((!frmcoupon.couponGubun[0].checked)&&(!frmcoupon.couponGubun[1].checked)){
        alert('���� ������ �����ϼ���..');
		frmcoupon.couponGubun[0].focus();
		return;
    }

    if (frmcoupon.couponGubun[1].checked){
        alert('���� ���� ����� �ý�����  ���� ���!');
    }

	if (frmcoupon.lecturercouponvalue.value.length<1){
		alert('���� �ݾ� �Ǵ� �������� �Է��� �ּ���.');
		frmcoupon.lecturercouponvalue.focus();
		return;
	}

	if (!IsDigit(frmcoupon.lecturercouponvalue.value)){
		alert('���� �ݾ� �Ǵ� �������� ���ڸ� �����մϴ�.');
		frmcoupon.lecturercouponvalue.focus();
		return;
	}


	if ((!frmcoupon.lecturercoupontype[0].checked)&&(!frmcoupon.lecturercoupontype[1].checked)){
		alert('���� Ÿ���� ������ �ּ���.');
		frmcoupon.lecturercouponvalue.focus();
		return;
	}

    //if ((frmcoupon.lecturercoupontype[2].checked)&&(frmcoupon.lecturercouponvalue.value!='2000')){
	//	alert('������ ������ ���ξ��� 2000�� �Դϴ�.');
	//	frmcoupon.lecturercouponvalue.focus();
	//	return;
	//}

	//if ((frmcoupon.lecturercoupontype[2].checked)&&!(frmcoupon.margintype.value=='20'||frmcoupon.margintype.value=='50'||frmcoupon.margintype.value=='80')){
	//	alert('������ ���� �߱޽� �ݹݺδ�, �������� �Ǵ� ������500��ü�δ����� �������ּ���.');
	//	frmcoupon.margintype.focus();
	//	return;
	//}

	if (frmcoupon.lecturercouponstartdate.value.length!=10){
		alert('���� �߱� �������� �Է��� �ּ���.');
		frmcoupon.lecturercouponstartdate.focus();
		return;
	}

	if (frmcoupon.lecturercouponstartdate2.value.length!=8){
		alert('���� �߱� �������� �Է��� �ּ���.');
		frmcoupon.lecturercouponstartdate2.focus();
		return;
	}

	if (frmcoupon.lecturercouponexpiredate.value.length!=10){
		alert('���� �߱� �������� �Է��� �ּ���.');
		frmcoupon.lecturercouponexpiredate.focus();
		return;
	}

	if (frmcoupon.lecturercouponexpiredate2.value.length!=8){
		alert('���� �߱� �������� �Է��� �ּ���.');
		frmcoupon.lecturercouponexpiredate2.focus();
		return;
	}

	if (frmcoupon.margintype.value.length<1){
		alert('���� ������ ������ �ּ���.');
		frmcoupon.margintype.focus();
		return;
	}

	if (frmcoupon.margintype.value==20){
		if (frmcoupon.defaultmargin.value.length<1){
			alert('������ �Է��� �ּ���.');
			frmcoupon.defaultmargin.focus();
			return;
		}

	}

    if (isEditMode){
        if (confirm('���� �Ͻðڽ��ϱ�?')){
    		frmcoupon.submit();
    	}
    }else{
    	if (confirm('���� �Ͻðڽ��ϱ�?')){
    		frmcoupon.submit();
    	}
    }
}

</script>

<!---- �˻� ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		������ȣ : <input type="text" name="lecturercouponidx" value="<%= lecturercouponidx %>" Maxlength="12" size="12" readonly >
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>	
</form>
</table>
<!---- /�˻� ---->

<br>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#BABABA">
<form name="frmcoupon" method="post" action="lecturercoupon_Process.asp">
<input type="hidden" name="lecturercouponidx" value="<%= lecturercouponidx %>">
<input type="hidden" name="mode" value="couponmaster">
<tr bgcolor="#DDDDFF">
	<td width="100">������</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="lecturercouponname" value="<%= oitemcouponmaster.FOneItem.Flecturercouponname %>" size="40" maxlength="30"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="100">��������</td>
	<td bgcolor="#FFFFFF">
	    <input type="radio" name="couponGubun" value="C" <%= ChkIIF(oitemcouponmaster.FOneItem.FcouponGubun="C","checked","") %> >�Ϲ�
	    <!--<input type="radio" name="couponGubun" value="T" <%= ChkIIF(oitemcouponmaster.FOneItem.FcouponGubun="T","checked","") %> >Ÿ��(E-mailƯ��)-->
	    <input type="radio" name="couponGubun" value="P" <%= ChkIIF(oitemcouponmaster.FOneItem.FcouponGubun="P","checked","") %> >�����ι߱�(����Ʈ �߱� �Ұ� : �ý����� ����)
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="lecturercouponvalue" value="<%= oitemcouponmaster.FOneItem.Flecturercouponvalue %>" size="6">
		<input type="radio" name="lecturercoupontype" value="1" <% if oitemcouponmaster.FOneItem.Flecturercoupontype="1" then response.write "checked" %> > %
		<input type="radio" name="lecturercoupontype" value="2" <% if oitemcouponmaster.FOneItem.Flecturercoupontype="2" then response.write "checked" %> > ��
		<!--<input type="radio" name="lecturercoupontype" value="3" <% if oitemcouponmaster.FOneItem.Flecturercoupontype="3" then response.write "checked" %> > ��۷��������� (2000 �Է�)-->
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >����Ⱓ</td>
	<td bgcolor="#FFFFFF">
	<input type="text" class="text" name="lecturercouponstartdate" value="<%= Left(oitemcouponmaster.FOneItem.Flecturercouponstartdate,10) %>" size="10" maxlength="10">
	<input type="text" class="text_ro" name="lecturercouponstartdate2" value="<%= ChkIIF(oitemcouponmaster.FOneItem.Flecturercouponstartdate<>"",Right(oitemcouponmaster.FOneItem.Flecturercouponstartdate,8),"00:00:00") %>" size="8" maxlength="8">
	<a href="javascript:calendarOpen(frmcoupon.lecturercouponstartdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	~
	<input type="text" class="text" name="lecturercouponexpiredate" value="<%= Left(oitemcouponmaster.FOneItem.Flecturercouponexpiredate,10) %>" size="10" maxlength="10">
	<input type="text" class="text_ro" name="lecturercouponexpiredate2" value="<%= ChkIIF(oitemcouponmaster.FOneItem.Flecturercouponexpiredate<>"",Right(oitemcouponmaster.FOneItem.Flecturercouponexpiredate,8),"23:59:59") %>" size="8" maxlength="8">
	<a href="javascript:calendarOpen(frmcoupon.lecturercouponexpiredate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	<br>(<%= Left(now(),10) %> 00:00:00)  ~  (<%= Left(now(),10) %> 23:59:59)
	<br><font color="#808080">(�� ���� �̹� �ٿ�ε��� ������ ����Ⱓ�� ������� �ʽ��ϴ�. ���� �Ⱓ �����ÿ� �������ּ���.)</font>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >�⺻ ��������</td>
	<td bgcolor="#FFFFFF">
		<select name="margintype" onchange="AlertMarginChange();fninput(this.value);">
		<!--<option value="">---����--- -->
		<!--<option value="30" <% if oitemcouponmaster.FOneItem.Fmargintype="30" then response.write "selected" %> >���ϸ���-->
		<option value="60" <% if oitemcouponmaster.FOneItem.Fmargintype="60" then response.write "selected" %> >��ü�δ�
		<!--<option value="50" <% if oitemcouponmaster.FOneItem.Fmargintype="50" then response.write "selected" %> >�ݹݺδ�
		<option value="10" <% if oitemcouponmaster.FOneItem.Fmargintype="10" then response.write "selected" %> >�ΰŽ��δ�
		<option value="20" <% if oitemcouponmaster.FOneItem.Fmargintype="20" then response.write "selected" %> >��������
		<option value="00" <% if oitemcouponmaster.FOneItem.Fmargintype="00" then response.write "selected" %> >��ǰ��������
		<option value="90" <% if oitemcouponmaster.FOneItem.Fmargintype="90" then response.write "selected" %> >20%��ü���
		<option value="80" <% if oitemcouponmaster.FOneItem.Fmargintype="80" then response.write "selected" %> >������(500��ü�δ�)-->
		</select>
		<span id="marginlayer" style="display:<% IF oitemcouponmaster.FOneItem.Fmargintype<>"20" Then response.write "none" %>"><input type="text" class="text" name="defaultmargin" value="<%=oitemcouponmaster.FOneItem.FDefaultMargin%>" size="3" maxlength="3" onChange="AlertMarginChange();">%</span>
		<font color="#808080">(��ǰ���� �������� �ٸ� ��� ������ ���� �����մϴ�.)</font>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >��������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="lecturercouponexplain" value="<%= oitemcouponmaster.FOneItem.Flecturercouponexplain %>" size="60" maxlength="50">
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >�߱� ����</td>
	<td bgcolor="#FFFFFF">
	<%= oitemcouponmaster.FOneItem.GetOpenStateName %>
	<% if (oitemcouponmaster.FOneItem.Flecturercouponidx>0) then %>
    	<% if (oitemcouponmaster.FOneItem.IsOpenAvailCoupon) then %>
    		--&gt;<input type="button" value="����" onclick="OpenCouponMaster();" class="button">
    	<% elseif (oitemcouponmaster.FOneItem.Fopenstate="0")  then %>
    		--&gt;<input type="button" value="�߱޿���" onclick="reserveCouponMaster();" class="button">
    	<% elseif (oitemcouponmaster.FOneItem.Fopenstate="9")  then %>

    	<% else %>
    	--&gt;<input type="button" value="�߱ް�������" onclick="CloseCouponMaster();" class="button">
    	(������ 12�� 15�п� �ڵ� ����˴ϴ�.)
    	<% end if %>
    <% end if %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >�����</td>
	<td bgcolor="#FFFFFF">
		<%= oitemcouponmaster.FOneItem.Fregdate %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<% if (IsEditMode) then %>
	    <% if (oitemcouponmaster.FOneItem.Fopenstate="0") then %>
	    <td colspan="2" align="center"><input type="button" value="�� ��" onclick="SaveCouponMaster(frmcoupon, true)" class="button"></td>
	    <% elseif (Not oitemcouponmaster.FOneItem.IsOpenAvailCoupon) then %>
	    <td colspan="2" align="center"><input type="button" value="�� ��" onclick="SaveCouponMaster(frmcoupon, true)" class="button" Disabled ></td>
	    <% else %>
	    <td colspan="2" align="center"><input type="button" value="�� ��" onclick="SaveCouponMaster(frmcoupon, true)" class="button"></td>
	    <% end if %>
	<% else %>
	<td colspan="2" align="center"><input type="button" value="�� ��" onclick="SaveCouponMaster(frmcoupon, false)" class="button"></td>
	<% end if %>
</tr>
</form>
</table>

<%
	set oitemcouponmaster = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->