<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim makerid
makerid = session("ssBctID")


dim opartner
set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid

if makerid<>"" then
	opartner.GetOnePartnerNUser
end if

dim ooffontract
set ooffontract = new COffContractInfo
ooffontract.FRectDesignerID = makerid

if makerid<>"" then
	ooffontract.GetPartnerOffContractInfo
end if

dim i

''DefaultCenterMwdiv  
dim DefaultCenterMwdiv
DefaultCenterMwdiv = GetDefaultItemMwdivByBrand(makerid)
%>
<script language='javascript'>
function AttachImage(comp,filecomp){
	comp.src=filecomp.value;
}

function ChangeBrand(comp){
	location.href="?makerid=" + comp.value;
}

function CheckAddItem(frm){
	if ((frm.itemgubun[0].checked==false)&&(frm.itemgubun[1].checked==false)){
		alert('��ǰ������ �����ϼ���.');
		return;
	}

	if (frm.cd1.value.length<1){
		alert('ī�װ��� �����ϼ���.');
		return;
	}

	if (frm.makerid.value.length<1){
		alert('�귣�带 �����ϼ���.');
		return;
	}

	if (frm.shopitemname.value.length<1){
		alert('��ǰ���� �Է��ϼ���.');
		frm.shopitemname.focus();
		return;
	}

	if ((frm.extbarcode.value.length>0) && (frm.extbarcode.value.length<10)){
		alert('���ڵ� ���̰� �ʹ� ª���ϴ�.  ���� ���ڵ尡 �ִ°�츸 �Է��� �ּ���');
		frm.extbarcode.focus();
		return;
	}

	if (!IsDigit(frm.shopitemprice.value)){
		alert('�ǸŰ��� ���ڸ� �����մϴ�.');
		frm.shopitemprice.focus();
		return;
	}


//	if (!IsDigit(frm.discountsellprice.value)){
//		alert('���� �ǸŰ��� ���ڸ� �����մϴ�.');
//		frm.discountsellprice.focus();
//		return;
//	}


	if (!IsDigit(frm.shopsuplycash.value)){
		alert('��ü ���԰��� ���ڸ� �����մϴ�.');
		frm.shopsuplycash.focus();
		return;
	}

	if (!IsDigit(frm.shopbuyprice.value)){
		alert('�� ���ް��� ���ڸ� �����մϴ�.');
		frm.shopbuyprice.focus();
		return;
	}

	if (((frm.shopsuplycash.value!=0)||(frm.shopbuyprice.value!=0))){
		if (!confirm('!! �⺻ ��� ������ �ٸ� ��쿡�� ���԰� ���ް��� �Է� �ϼž� �մϴ�. \n\n��� �Ͻðڽ��ϱ�?')){
			return;
		}
	}

    if (frm.file1.value.length<1){
		alert('�̹����� �Է��� �ּ��� - �ʼ� �����Դϴ�.');
		frm.file1.focus();
		return;
	}
    
//	if (frm.ioffimgmain.src.length<1){
//		alert('�̹����� �Է��� �ּ��� - �ʼ� �����Դϴ�.');
//		frm.file1.focus();
//		return;
//	}

//	if (!((/(.jpg|.jpeg)$/i).test(frm.ioffimgmain.src))){
//	  alert('�⺻ �̹����� jpg ���ϸ� �����մϴ�.');
//	}

	var ret = confirm('�������� ���� ��ǰ���� ��� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.submit();
	}
}


// ============================================================================
// ī�װ����
function editCategory(cdl,cdm,cds){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cds=" + cds ;

	popwin = window.open('/common/module/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function setCategory(cd1,cd2,cd3,cd1_name,cd2_name,cd3_name){
	var frm = document.frmedit;
	frm.cd1.value = cd1;
	frm.cd2.value = cd2;
	frm.cd3.value = cd3;
	frm.cd1_name.value = cd1_name;
	frm.cd2_name.value = cd2_name;
	frm.cd3_name.value = cd3_name;
}
</script>
<table border=0 cellspacing=1 cellpadding=2 width=460 class="a" bgcolor=#FFFFFF>
<tr>
	<td>&gt;&gt;�������� ��ǰ ���</td>
</tr>
</table>

<table border=0 cellspacing=1 cellpadding=2 width=460 class="a" bgcolor=#3d3d3d>
<% if application("Svr_Info")="Dev" then %>	
<form name="frmedit" method=post action="http://testpartner.10x10.co.kr/linkweb/dooffitemimageeditwithdata.asp" enctype="MULTIPART/FORM-DATA">
<% else %>
<form name="frmedit" method=post action="http://partner.10x10.co.kr/linkweb/dooffitemimageeditwithdata.asp" enctype="MULTIPART/FORM-DATA">
<% end if %>
<input type=hidden name=mode value="addnewoffitem">
<input type=hidden name=makerid value="<%= makerid %>">

<tr bgcolor="#FFDDDD">
	<td width=100>�귣��������</td>
	<td bgcolor="#FFFFFF" colspan=5><a href="javascript:PopUpcheInfo('<%= makerid %>');"><%= makerid %></a> (<%= opartner.FOneItem.Fsocname_kor %>,<%= opartner.FOneItem.FCompany_name %>)
	</td>
</tr>
<tr bgcolor="#FFDDDD">
	<td width=100 >�¶���</td>
	<td bgcolor="#FFFFFF" colspan=5><%= opartner.FOneItem.GetMWUName %> &nbsp;&nbsp; <%= opartner.FOneItem.Fdefaultmargine %> %</td>
</tr>

<tr bgcolor="#FFDDDD">
	<td width=100>��������-����</td>
	<td bgcolor="#FFFFFF" colspan=5>
		<table border=0 cellspacing=0 cellpadding=0 class=a width=80%>
		<tr>
			<td ><a href="javascript:editOffDesinger('streetshop000','<%= makerid %>')"><b>��������ǥ</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop000") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop000") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="1") then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= makerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td width=60><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td width=60><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
</tr>
<tr bgcolor="#FFDDDD">
	<td width=100>��������-����</td>
	<td bgcolor="#FFFFFF" colspan=5>
		<table border=0 cellspacing=0 cellpadding=0 class=a width=80%>
		<tr>
			<td ><a href="javascript:editOffDesinger('streetshop800','<%= makerid %>')"><b>����������ǥ</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop800") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop800") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="3") then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= makerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td ><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="5") then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= makerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td ><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>��ǰ����</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type="radio" name="itemgubun" value="90" checked >������ �����ǰ(90) &nbsp;
	<input type="radio" name="itemgubun" value="70">�Ҹ�ǰ(70)
	<br><font color="#AAAAAA">(90������������, 80�̺�Ʈ ,70�Ҹ�ǰ, 95���������������Ǹ�)</font>
	</td>
</tr>
<tr bgcolor="#DDDDFF" >
	<td width=100 >ī�װ�</td>
	<td bgcolor="#FFFFFF" colspan=5>
	  <input type="hidden" name="cd1" value="">
	  <input type="hidden" name="cd2" value="">
	  <input type="hidden" name="cd3" value="">

      <input type="text" name="cd1_name" value="" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd2_name" value="" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd3_name" value="" size="12" readonly style="background-color:#E6E6E6">

      <input type="button" value="����" onclick="editCategory(frmedit.cd1.value,frmedit.cd2.value,frmedit.cd3.value);">
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>��ǰ��</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=text name="shopitemname" value="" size=40 maxlength=30 class="input_01" >
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>�ɼǸ�</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=hidden name="shopitemoptionname" value="">
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>������ڵ�</td>
	<td bgcolor="#FFFFFF" colspan=5><input type=text name="extbarcode" value="" size=20 maxlength=20 class="input_01" >(�ִ� ��츸 ���)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>�������</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=radio name=isusing value="Y" checked >�����
	<input type=radio name=isusing value="N">������
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>���͸��Ա���</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=radio name=centermwdiv value="W" <%= ChkIIF(DefaultCenterMwdiv<>"M","checked","") %> >Ư��
	<input type=radio name=centermwdiv value="M" <%= ChkIIF(DefaultCenterMwdiv="M","checked","") %>>����
	</td>
</tr>
<tr bgcolor="#DDDDFF" >
	<td width=100 >��������</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=radio name=vatinclude value="Y" checked >����
	<input type=radio name=vatinclude value="N">�鼼
	</td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td width=100 align="left" rowspan="3">���ݼ���</td>
	<td bgcolor="#FFFFFF" >�ǸŰ�</td>
	<td bgcolor="#FFFFFF" >���԰�</td>
	<td bgcolor="#FFFFFF" >���ް�</td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td bgcolor="#FFFFFF"><input type=text name="shopitemprice" value="" size=8 maxlength=9 class="input_right" ></td>
	<td bgcolor="#FFFFFF"><input type=text name="shopsuplycash" value="0" size=8 maxlength=9 class="input_right" style="background-color : #DDDDDD" readonly ></td>
	<td bgcolor="#FFFFFF" ><input type=text name="shopbuyprice" value="0" size=8 maxlength=9 class="input_right" style="background-color : #DDDDDD" readonly ></td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td bgcolor="#FFFFFF" ></td>
	<td bgcolor="#FFFFFF" colspan="3">(0 �ΰ�� �⺻���� ���� �ڵ� ������)</td>
</tr>

</tr>
<tr bgcolor="#DDDDFF">
	<td width=100 valign=top>������ǰ<br>�̹���</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=file name=file1 class="input_01" size=20 onchange="AttachImage(ioffimgmain,this)" >(400 x 400 px)
	<br>(�⺻ �̹����� �� <b>jpg</b> ���Ϸ� �÷��ֽñ� �ٶ��ϴ�.)
	<img name="ioffimgmain" src="" width=340 height=340>
	</td>
</tr>
</form>
<tr bgcolor="#FFFFFF">
	<td colspan=6 align=center><input type=button value=" ��  �� " onclick="CheckAddItem(frmedit)" class="input_01"></td>
</tr>
</table>

<%
set opartner = Nothing
set ooffontract = Nothing
%>

<script language='javascript'>
alert('�¶��ο��� �Ǹ��ϰų� �Ǹ� ������ ��ǰ�� ���� ������� ���ñ� �ٶ��ϴ�. \n ��ǰ�� ���ߵ�ϵǾ� �������� ������� �ֽ��ϴ�. \n �¶��λ�ǰ ����� �Ϸ������� �������� ��ǰ���� �ڵ� ��ϵ˴ϴ�. \n\n �� �ʿ��Ѱ��(�������ο����� �Ǹ��Ͻ� ��� ��)�� ������ּ���');
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->