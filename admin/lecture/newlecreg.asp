<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual ="/lib/classes/partnerusercls.asp" -->
<!-- #include virtual="/lib/classes/lecture_itemregcls.asp"-->
<%
Sub SelectBoxDesignerItem1()
	dim query1
	%>
	<select name="tempid" onchange="TnDesignerNMargineAppl(this.value);">
	<option value=''>-- ��ü���� --</option>
	<%

	query1 = "select c.userid, c.coname from [db_user].[dbo].tbl_user_c c" + vbcrlf
	query1 = query1 + " where c.userdiv='14'" + vbcrlf

	rsget.Open query1,dbget,1

	if  not rsget.EOF  then
		rsget.Movefirst

		do until rsget.EOF
			response.write ("<option value='" & rsget("userid") & "," &rsget("coname") & "'>" & rsget("userid") & " (" & rsget("coname") & ")</option>")
			rsget.MoveNext
		loop

	end if

	rsget.close

	response.write("</select>")
End Sub

%>
<script language="JavaScript">
<!--

function checkform(form) {
//alert('����(4�� 22��) ���� ���� ����� ������ ���ε尡 �Ұ��� �մϴ�. ��ø� ��ٷ� �ּ���');
//return;

	var limitynv = "";
	var optionv="";
	var aa="";
	var bb="";
	var cc="";
	var dd="";
	var ee="";
	var ff="";
	var gg="";
	var hh="";

	aa=document.getElementById("imgmainload");
	bb=document.getElementById("imgbasicload");
	dd=document.getElementById("imgadd1load");
	ee=document.getElementById("imgadd2load");
	ff=document.getElementById("imgadd3load");
	gg=document.getElementById("imgadd4load");
	hh=document.getElementById("imgadd5load");

	for (var i = 0; i < form.limityn.length; i++) {
	if ( form.limityn[i].checked) {
		 limitynv = form.limityn[i].value
	   }
	}

	//for(var i=0; i<document.itemreg.realopt.options.length; i++) {
		//optionv += (document.itemreg.itemoptionnameno.value + document.itemreg.itemoptioncode.options[i].value + ",")
	//	optionv += (document.itemreg.realopt.options[i].value + ",")
	 //}

	if (form.cd1.value == ""){
	  alert("ī�װ��� �������ּ���!");
	  form.cd1.focus();
	  return;
	}

	if (form.cd2.value == ""){
	  alert("ī�װ��� �������ּ���!");
	  form.cd2.focus();
	  return;
	}

	if (form.cd3.value == ""){
	  alert("ī�װ��� �������ּ���!");
	  form.cd3.focus();
	  return;
	}

//	if (i2ndcate.style.display=="inline"){
//		if (form.stylegubun.value == ""){
//		  alert("��Ÿ�� ������ �������ּ���!");
//		  return;
//		}
//
//		if (form.itemstyle.value == ""){
//		  alert("��Ÿ���� �������ּ���!");
//		  return;
//		}
//	}

	if(form.itemname.value == ""){
	  alert("��ǰ���� �Է����ּ���!");
	  form.itemname.focus();
	  return;
	}
//	else if(form.itemsource.value.length<1){
//	  alert("��ǰ������ �Է����ּ���!");
//	  form.itemsource.focus();
//	  return;
//	}
//	else if(form.itemsize.value.length<1){
//	  alert("��ǰ����� �Է����ּ���!");
//	  form.itemsize.focus();
//	  return;
//	}
//	else if(form.sourcearea.value == ""){
//	  alert("�������� �Է����ּ���!");
//	  form.sourcearea.focus();
//	  return;
//	}
//	else if(form.makename.value == ""){
//	  alert("�����縦 �Է����ּ���!");
//	  form.makename.focus();
//	  return;
//	}
//	else if(form.keywords.value == ""){
//	  alert("�˻� Ű���带 �Է����ּ���!");
//	  form.keywords.focus();
//	  return;
//	}
	else if(!IsDigit(form.sellcash.value)){
	  alert("�ǸŰ��� ���ڸ� �����մϴ�.");
	  form.sellcash.focus();
	  return;
	}
	else if(form.buycash.value == ""){
	  alert("���ް��� �Է����ּ���!");
	  form.buycash.focus();
	  return;
	}
	else if(!IsDigit(form.buycash.value)){
	  alert("���ް��� ���ڸ� �����մϴ�.");
	  form.buycash.focus();
	  return;
	}
	else if(limitynv == "Y" && form.limitno.value == ""){
	  alert("���������� �Է����ּ���!");
	  form.limitno.focus();
	  return;
	}
	else if(limitynv == "Y" && !IsDigit(form.limitno.value)){
	  alert("���������� ���ڸ� �����մϴ�.");
	  form.limitno.focus();
	  return;
	}
	//else if(form.itemoptionname.value == "" || optionv == ""){
	//  alert("�ɼ��� �������ּ���!");
	//  form.itemoptionname.focus();
	//}

	else if(aa.fileSize > 150000){
		alert("���ϻ������ 150Kbyte�� �ѱ�� �� �����ϴ�...");
		form.imgmain.focus();
	}
	else if(aa.width > 610){
		alert("�������� 600�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgmain.focus();
	}
	else if(aa.height > 410){
		alert("�������� 400�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgmain.focus();
	}

//---------------------------------------------------------
	else if(bb.fileSize > 150000){
		alert("���ϻ������ 150Kbyte�� �ѱ�� �� �����ϴ�...");
		form.imgbasic.focus();
	}
	else if(bb.width > 410){
		alert("�������� 400�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgbasic.focus();
	}
	else if(bb.height > 410){
		alert("�������� 400�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgbasic.focus();
	}
//---------------------------------------------------------
	else if(dd.fileSize > 150000){
		alert("���ϻ������ 150Kbyte�� �ѱ�� �� �����ϴ�...");
		form.imgadd1.focus();
	}
	else if(dd.width > 610){
		alert("�������� 600�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgadd1.focus();
	}
	else if(dd.height > 410){
		alert("�������� 400�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgadd1.focus();
	}
//---------------------------------------------------------
	else if(ee.fileSize > 150000){
		alert("���ϻ������ 150Kbyte�� �ѱ�� �� �����ϴ�...");
		form.imgadd2.focus();
	}
	else if(ee.width > 610){
		alert("�������� 600�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgadd2.focus();
	}
	else if(ee.height > 410){
		alert("�������� 400�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgadd2.focus();
	}
//---------------------------------------------------------
	else if(ff.fileSize > 150000){
		alert("���ϻ������ 150Kbyte�� �ѱ�� �� �����ϴ�...");
		form.imgadd3.focus();
	}
	else if(ff.width > 610){
		alert("�������� 600�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgadd3.focus();
	}
	else if(ff.height > 410){
		alert("�������� 400�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgadd3.focus();
	}
//---------------------------------------------------------
	else if(gg.fileSize > 150000){
		alert("���ϻ������ 150Kbyte�� �ѱ�� �� �����ϴ�...");
		form.imgadd4.focus();
	}
	else if(gg.width > 610){
		alert("�������� 600�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgadd4.focus();
	}
	else if(gg.height > 410){
		alert("�������� 400�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgadd4.focus();
	}
//---------------------------------------------------------
	else if(hh.fileSize > 150000){
		alert("���ϻ������ 150Kbyte�� �ѱ�� �� �����ϴ�...");
		form.imgadd5.focus();
	}
	else if(hh.width > 610){
		alert("�������� 600�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgadd5.focus();
	}
	else if(hh.height > 410){
		alert("�������� 400�ȼ��� �ѱ�� �� �����ϴ�...");
		form.imgadd5.focus();
	}

//---------------------------------------------------------
	//else if(form.imglist.value == ""){
	//  alert("����Ʈ�̹����� �������ּ���!");
	//  form.imglist.focus();
	//}
	//else if(form.imgsmall.value == ""){
	//  alert("�����̹����� �������ּ���!");
	//  form.imgsmall.focus();
	//}

    else{
		if(confirm("��ǰ�� �ø��ðڽ��ϱ�?") == true){
		//form.itemoptioncode2.value=optionv;
		//alert(form.itemoptioncode2.value);
<!--		form.submit();-->
		}
	}
}


function CalcuAuto(frm){
	var imargin, isellcash, ibuycash;
	var isellvat, ibuyvat, imileage;
	imargin = frm.margin.value;
	isellcash = frm.sellcash.value;

	isvatinclude = frm.vatinclude.value;

	if (imargin.length<1){
		alert('������ �Է��ϼ���.');
		frm.margin.focus();
		return;
	}

	if (isellcash.length<1){
		alert('�ǸŰ��� �Է��ϼ���.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDouble(imargin)){
		alert('������ ���ڷ� �Է��ϼ���.');
		frm.margin.focus();
		return;
	}

	if (!IsDigit(isellcash)){
		alert('�ǸŰ��� ���ڷ� �Է��ϼ���.');
		frm.sellcash.focus();
		return;
	}

	if (isvatinclude=='Y'){
		isellvat = parseInt(parseInt(1/11 * parseInt(isellcash)));
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		ibuyvat = parseInt(parseInt(1/11 * parseInt(ibuycash)));
		imileage = parseInt(isellcash*0.01) ;
	}else{
		isellvat = 0;
		ibuycash = isellcash - parseInt(isellcash*imargin/100);
		ibuyvat = 0;
		imileage = parseInt(isellcash*0.01) ;
	}

	frm.sellvat.value = isellvat;
	frm.buycash.value = ibuycash;
	frm.buyvat.value = ibuyvat;
	frm.mileage.value = imileage;
}

function TnDesignerNMargineAppl(str){
	var varArray;
	varArray = str.split(',');

	document.itemreg.designerid.value = varArray[0];
	document.itemreg.lecturerid.value = varArray[0];
	document.itemreg.lecturer.value = varArray[1];


}

function TnBasicItemInfo(){
	window.open("/admin/lecture/basic_lecture_info.asp","option_win","width=300,height=200,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes");
}

//-->
</script>

<form name="itemreg" method="post" action="http://partner.10x10.co.kr/admin/shopmaster/lecture_itemreg_upload_bywebadmin.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="itemoptioncode2">
<input type="hidden" name="designerid" value="">

<!-- itemreg.asp ��������-->
<input type="hidden" name="itemsource" value=""><!-- ��ǰ���� -->
<input type="hidden" name="itemsize" value=""><!-- ��ǰ������ -->
<input type="hidden" name="sourcearea" value=""><!-- ������ -->
<input type="hidden" name="makename" value=""><!-- ������ -->
<input type="hidden" name="mwdiv" value="U"><!-- ������Ź����, ��ü������(U)-->
<input type="hidden" name="vatinclude" value="N"><!-- ����, �鼼 ����, N-->
<input type="hidden" name="deliverytype" value="5"><!-- ��۱���, N-->
<input type="hidden" name="limityn" value="Y"><!-- �����Ǹű���, Y-->
<input type="hidden" name="pojangok" value="Y"><!-- ���尡�ɿ���, N-->
<input type="hidden" name="sellyn" value="N"><!-- �Ǹſ���, Y-->
<input type="hidden" name="dispyn" value="N"><!-- ���ÿ���, N-->
<input type="hidden" name="isusing" value="N"><!-- ��뿩��, N-->
<input type="hidden" name="usinghtml" value="N"><!-- HTML�������, N-->
<input type="hidden" name="itemcontent" value=""><!-- ������ ����-->
<input type="hidden" name="ordercomment" value=""><!-- �ֹ��� ���ǻ��� -->
<input type="hidden" name="designercomment" value=""><!-- ��ü�ڸ�Ʈ -->
<!-- itemreg.asp -->

<table width="750" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#3d3d3d">

<tr>
	<td width="100%">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td colspan="4" style="padding-left:20"><a href="javascript:TnBasicItemInfo();"><font color="red">�⺻Ʋ����</font></a></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120" style="spacing-left:1px">���� �� ���� <font color="red">(*)</font></td>
				<td colspan="3"><input type="text" name="yyyymm" value="" size="7" maxlength="7" class="input_b">(<%= Left(now(),7) %>)</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">�귣�� <font color="red">(*)</font></td>
				<td><% SelectBoxDesignerItem1 %></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">���� ī�װ� <font color="red">(*)</font></td>
				<td>
					<select name="cd1">
						<option value="95">���þ���[95]</option>
					</select>
					<select name="cd2">
						<option value="20">College[20]</option>
					</select>
					<select name="cd3">
						<option value="10">College[10]</option>
					</select>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">���¸� <font color="red">(*)</font></td>
				<td><input type="text" name="itemname" maxlength="64" size="50" class="input_b"></td>
			</tr>
		</table>
	</td>
</tr>


<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">�˻�Ű���� <font color="red">(*)</font></td>
		<td colspan="3"><input type="text" name="keywords" maxlength="50" size="50" class="input_b" value="����,��ī����,������,�ø���">&nbsp;(�޸��α��� ex: Ŀ��,Ƽ����,����)</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">����</td>
				<td colspan="3" >
					<input type="text" name="margin" maxlength="32" size="5" class="input_b" value="50">%
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">�ǸŰ�(�Һ��ڰ�) <font color="red">(*)</font></td>
				<td colspan="3"><input type="text" name="sellcash" maxlength="16" size="16" class="input_b">��&nbsp;&nbsp;<input type="text" name="sellvat" maxlength="32" size="10" class="input_b">&nbsp;&nbsp;<font color="red"><input type="button" value="���ް� �ڵ� ���" class="button" onclick="CalcuAuto(itemreg);"></font></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">���԰� <font color="red">(*)</font></td>
				<td colspan="3">
					<input type="text" name="buycash" maxlength="16" size="16" class="input_b">��&nbsp;&nbsp;<input type="text" name="buyvat" maxlength="32" size="10" class="input_b"> (<b>�ΰ��� ���԰�</b>�� �Է��� �ּ���.)
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">���ϸ��� <font color="red">(*)</font></td>
				<td colspan="3"><input type="text" name="mileage" maxlength="32" size="10" class="input_b"> (�⺻ �ǸŰ��� 1%)</td>
			</tr>
		</table>
	</td>
</tr>

<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">�ҼӾ��̵�</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecturerid" value=""  class="input_b"size="30" maxlength="32"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">�����</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecturer" value=""  class="input_b"size="30" maxlength="32"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">���º�</td>
				<td bgcolor="#FFFFFF" width="250">
					<input type="text" name="lecsum" value="" class="input_b" size="12" maxlength="12">
					<input type="checkbox" name="matinclude">��������
				</td>
				<td bgcolor="#DDDDFF" width="120">����</td>
				<td bgcolor="#FFFFFF"><input type="text" name="matsum" value=""  class="input_b" size="12" maxlength="12"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">���񼳸�</td>
				<td bgcolor="#FFFFFF"><input type="text" name="matdesc" value=""  class="input_b" size="90" maxlength="128"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">���</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecspace" size="30" value=""  class="input_b"maxlength="64"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">����Ƚ��</td>
				<td bgcolor="#FFFFFF"><input type="text" name="leccount" value=""  class="input_b"size="6" maxlength="12"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">���ǽð�</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lectime" value=""  class="input_b"size="20" maxlength="12"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">�Ѱ��ǽð�</td>
				<td bgcolor="#FFFFFF"><input type="text" name="tottime" value=""  class="input_b"size="6" maxlength="12"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">���ǱⰣ(�ֱ�)</td>
				<td bgcolor="#FFFFFF"><input type="text" name="lecperiod" value=""  class="input_b"size="30" maxlength="64">(ex : ���� �ݿ��� ���~���)</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120"><font color="red">*</font>��������<font color="red">*</font></td>
				<td bgcolor="#FFFFFF" width="160"><input type="text" name="limitno" maxlength="32" style="background-color:#FFFFFF;" class="input_b">(��)</td>
				<td bgcolor="#DDDDFF" width="120">�����ο�</td>
				<td bgcolor="#FFFFFF"><input type="text" name="properperson" value="" class="input_b" size="6" maxlength="12"></td>
				<td bgcolor="#DDDDFF" width="120">�ּ��ο�</td>
				<td bgcolor="#FFFFFF" ><input type="text" name="minperson" value="" class="input_b" size="6" maxlength="12"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">��������</td>
				<td bgcolor="#FFFFFF" width="250"><input type="text" name="reservestart" value="" class="input_b" size="15" maxlength="10" onclick="calender_open('reservestart');"></td>
				<td bgcolor="#DDDDFF" width="120">���ึ����</td>
				<td bgcolor="#FFFFFF"><input type="text" name="reserveend" value="" class="input_b" size="15" maxlength="10" onclick="calender_open('reserveend');"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">���³���<br>(Ŀ��ŧ��)</td>
				<td bgcolor="#FFFFFF">
					<table border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a" >
						<tr bgcolor="#DDDDFF">
							<td>1��</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate01" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate01');">~<input type="text" name="lecdate01_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate01_end');">(2004-06-06 14:00:00)</td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>2��</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate02" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate02');">~<input type="text" name="lecdate02_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate02_end');"></td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>3��</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate03" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate03');">~<input type="text" name="lecdate03_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate03_end');"></td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>4��</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate04" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate04');">~<input type="text" name="lecdate04_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate04_end');"></td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>5��</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate05" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate05');">~<input type="text" name="lecdate05_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate05_end');"></td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>6��</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate06" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate06');">~<input type="text" name="lecdate06_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate06_end');"></td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>7��</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate07" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate07');">~<input type="text" name="lecdate07_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate07_end');"></td>
						</tr>
						<tr bgcolor="#DDDDFF">
							<td>8��</td>
							<td bgcolor="#FFFFFF"><input type="text" name="lecdate08" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate08');">~<input type="text" name="lecdate08_end" value="" class="input_b" size="20" maxlength="19" onclick="calender_open('lecdate08_end');"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">���°���</td>
				<td bgcolor="#FFFFFF"><textarea name="leccontents" class="input_b" rows="10" cols="80"></textarea></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">Ŀ��ŧ���Ұ�</td>
				<td bgcolor="#FFFFFF"><textarea name="leccurry" class="input_b" rows="10" cols="80"></textarea></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">��Ÿ����</td>
				<td bgcolor="#FFFFFF"><textarea name="lecetc" class="input_b" rows="10" cols="80"></textarea></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">��������</td>
				<td bgcolor="#FFFFFF">
				&nbsp;&nbsp;&nbsp;
				<input type="radio" name="regfinish" value="N" > ������
				<input type="radio" name="regfinish" value="Y" checked > ��������
				</td>
			</tr>
		</table>
	</td>
</tr>

<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">��뿩��</td>
				<td bgcolor="#FFFFFF">
				&nbsp;&nbsp;&nbsp;
				<input type="radio" name="isusing" value="Y" checked > �����(������)
				<input type="radio" name="isusing" value="N"  > ������(���þ���)
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#DDDDFF" width="120">��뿩��</td>
				<td bgcolor="#FFFFFF"><input type="button" value="��������" onclick="CheckForm();return false;">&nbsp;&nbsp;&nbsp;</td>
			</tr>
		</table>
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->