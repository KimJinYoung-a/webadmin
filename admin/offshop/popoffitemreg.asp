<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������λ�ǰ ���
' Hieditor : 2009.04.07 ������ ����
'			 2010.06.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
function drawOffContractBrandChangeEvent(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>" onchange="ChangeBrand(this)">
     <option value='' <%if selectedId="" then response.write " selected"%>>����</option><%
   query1 = " select c.userid, c.socname_kor from [db_user].[dbo].tbl_user_c c "
   query1 = query1 & " , [db_shop].[dbo].tbl_shop_designer s"
   query1 = query1 & " where c.userid = s.makerid "
   query1 = query1 & " and s.shopid='streetshop000'"
   query1 = query1 & " order by c.userid"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"&rsget("userid")&" ["&db2html(rsget("socname_kor"))&"]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Function

dim makerid , i
	makerid = requestCheckVar(request("makerid"),32)

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
	if ((frm.itemgubun[0].checked==false) && (frm.itemgubun[1].checked==false) && (frm.itemgubun[2].checked==false) && (frm.itemgubun[3].checked==false)){
		alert('��ǰ������ �����ϼ���.');
		return;
	}

	// ����ǰüũ
	var isgiftproduct = false;
	if (frm.itemgubun[2].checked == true) {
		isgiftproduct = true;
	}

	if (frm.makerid.value.length<1){
		alert('�귣�带 �����ϼ���.');
		return;
	}

	if (frm.cd1.value.length<1){
		alert('ī�װ��� �����ϼ���.');
		return;
	}

	if (frm.shopitemname.value.length<1){
		alert('��ǰ���� �Է��ϼ���.');
		frm.shopitemname.focus();
		return;
	}

	if ((frm.extbarcode.value.length>0) && (frm.extbarcode.value.length<10)){
		alert('���ڵ� ���̰� �ʹ� ª���ϴ�. ���� ���ڵ尡 �ִ°�츸 �Է��� �ּ���' );
		frm.extbarcode.focus();
		return;
	}

	if (frm.itemgubun[3].checked==true) {
        if (frm.shopitemprice.value ==''){
			alert("�ǸŰ��� �Է����ּ���.");
			frm.shopitemprice.focus();
			return;
		}
		
        if (frm.shopitemprice.value.substr(0,1) != '-'){
			frm.shopitemprice.value = "-"+frm.shopitemprice.value
		}							
	}else if (frm.itemgubun[2].checked==true) {
        if (frm.shopitemprice.value > 0){
			alert("����ǰ�� �ǸŰ��� 0���Ͽ��� �մϴ�.");
			frm.shopitemprice.focus();
			return;
		}

        if (frm.shopitemprice.value ==''){
			alert("�ǸŰ��� �Է����ּ���.");
			frm.shopitemprice.focus();
			return;
		}
	}else{
		if (!IsDigit(frm.shopitemprice.value)){
			alert('�ǸŰ��� ���ڸ� �����մϴ�.');
			frm.shopitemprice.focus();
			return;
		}	
	}
				
//	if (!IsDigit(frm.discountsellprice.value)){
//		alert('���� �ǸŰ��� ���ڸ� �����մϴ�.');
//		frm.discountsellprice.focus();
//		return;
//	}

	if (isgiftproduct == true) {
		if (frm.shopitemname.value.match(/����ǰ/) != null) {
			alert("����ǰ ������ ��ǰ�� �ڵ��Էµ˴ϴ�. ����ǰ ������ ���켼��.");
			return;
		}

		//if (frm.shopitemprice.value*1 != 0) {
		//	alert("����ǰ�� �ǸŰ��� 0������ �����ؾ� �մϴ�.");
		//	return;
		//}

		//if (frm.orgsellprice.value*1 != 0) {
		//	alert("����ǰ�� �Һ��ڰ��� 0������ �����ؾ� �մϴ�.");
		//	return;
		//}
	}

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

<% if application("Svr_Info") <> "Dev" then %>
	if (frm.file1.value.length<1){
		alert('�̹����� �Է��� �ּ��� - �ʼ� �����Դϴ�.');
		frm.file1.focus();
		return;
	}
<% end if %>

	var ret = confirm('�߰��Ͻðڽ��ϱ�?');

	if (ret) {
		if (isgiftproduct == true) {
			frm.shopitemname.value = "[����ǰ] " + frm.shopitemname.value;
		}

		frm.submit();
	}
}

// ī�װ����
function editCategory(cdl,cdm,cdn){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cdn=" + cdn ;

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

<table border=0 cellspacing=1 cellpadding=2 width="100%" class="a" bgcolor="#FFFFFF">
<tr>
	<td>&gt;&gt;�������� ��ǰ ���</td>
</tr>
</table>

<table border=0 cellspacing=1 cellpadding=2 width="100%" class="a" bgcolor="#3d3d3d">
<% if application("Svr_Info")="Dev" then %>
	<form name="frmedit" method="post" action="http://testpartner.10x10.co.kr/linkweb/dooffitemimageeditwithdata.asp" enctype="MULTIPART/FORM-DATA">
<% else %>
	<form name="frmedit" method="post" action="http://partner.10x10.co.kr/linkweb/dooffitemimageeditwithdata.asp" enctype="MULTIPART/FORM-DATA">
<% end if %>
<input type="hidden" name="mode" value="addnewoffitem">
<tr bgcolor="#FFDDDD">
	<td width=100>�귣�� ����</td>
	<td bgcolor="#FFFFFF" colspan=5><% drawOffContractBrandChangeEvent "makerid",makerid  %>
	</td>
</tr>
<% if makerid<>"" and opartner.FResultCount > 0 then %>
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
		<table border=0 cellspacing=0 cellpadding=0 class="a" width="80%">
		<tr>
			<td ><a href="javascript:editOffDesinger('streetshop000','<%= makerid %>')"><b>��������ǥ</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop000") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop000") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="1")  then %>
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
		<table border=0 cellspacing=0 cellpadding=0 class="a" width="80%">
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
	<input type="radio" name="itemgubun" value="80">����ǰ(80)
	<input type="radio" name="itemgubun" value="60">���α�(60)
	<br><font color="#AAAAAA">(90������������, 80����ǰ ,70�Ҹ�ǰ, 95���������������Ǹ� ,60���α�)</font>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100 >ī�װ�</td>
	<td bgcolor="#FFFFFF" colspan=5>
	  <input type="hidden" name="cd1" value="">
	  <input type="hidden" name="cd2" value="">
	  <input type="hidden" name="cd3" value="">

      <input type="text" name="cd1_name" value="" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd2_name" value="" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd3_name" value="" size="12" readonly style="background-color:#E6E6E6">

      <input type="button" value="����" onclick="editCategory(frmedit.cd1.value,frmedit.cd2.value,frmedit.cd3.value);" class="button">
	</td>
</tr>
<tr bgcolor="#DDDDFF" height="50">
	<td width=100>��ǰ��</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type="text" name="shopitemname" value="" size=40 maxlength=40 class="input_01" ><br>
	* ����ǰ�� ��ǰ�� "[����ǰ]" ������ �ڵ����� �ٽ��ϴ�.
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>�ɼǸ�</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type="text" name="shopitemoptionname" size=40 maxlength=40 value="" class="input_01">
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>������ڵ�</td>
	<td bgcolor="#FFFFFF" colspan=5><input type=text name="extbarcode" value="" size=20 maxlength=20 class="input_01" >(�ִ� ��츸 ���)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>�������</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type="radio" name="isusing" value="Y" checked >�����
	<input type="radio" name="isusing" value="N">������
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>���͸��Ա���</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type="radio" name="centermwdiv" value="W" <%= ChkIIF(DefaultCenterMwdiv<>"M","checked","") %> >Ư��
	<input type="radio" name="centermwdiv" value="M" <%= ChkIIF(DefaultCenterMwdiv="M","checked","") %>>����
	</td>
</tr>
<tr bgcolor="#DDDDFF" >
	<td width=100 >��������</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type="radio" name="vatinclude" value="Y" checked >����
	<input type="radio" name="vatinclude" value="N">�鼼
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
	<td bgcolor="#FFFFFF"><input type=text name="shopsuplycash" value="0" size=8 maxlength=9 class="input_right" ></td>
	<td bgcolor="#FFFFFF" ><input type=text name="shopbuyprice" value="0" size=8 maxlength=9 class="input_right" ></td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td bgcolor="#FFFFFF" ></td>
	<td bgcolor="#FFFFFF" colspan="2" align="left">
		* 0�ΰ�� �⺻���� ���� ������<br>
		* ����ǰ�� ��� ���� ������ �������
	</td>
</tr>

</tr>
<tr bgcolor="#DDDDFF">
	<td width=100 valign=top>������ǰ<br>�̹���</td>
	<td bgcolor="#FFFFFF" colspan=5 align="left">
		<input type="file" name="file1" class="input_01" size=20 onchange="AttachImage(ioffimgmain,this)" >(400 x 400 px)
		<br>(�⺻ �̹����� �� <b>jpg</b> ���Ϸ� �÷��ֽñ� �ٶ��ϴ�.)
		<img name="ioffimgmain" src="" width=340 height=340>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=6 align="center"><input type="button" value=" ��  �� " onclick="CheckAddItem(frmedit)" class="input_01"></td>
</tr>
<% end if %>
</form>
</table>

<%
set opartner = Nothing
set ooffontract = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->