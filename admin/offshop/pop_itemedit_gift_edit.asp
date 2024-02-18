<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ����ǰ ���
' Hieditor : 2013.01.15 �̻� ����
'			2017.04.13 �ѿ�� ����(���Ȱ���ó��)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim itemgubun,itemid, itemoption, barcode ,i ,makerid ,ioffitem ,opartner ,ooffontract ,IsOnlineItem
dim editmode , CenterMwDiv ,offList ,offSmall ,OnlineSailYn , IsDirectIpchulContractExistsBrand
dim shopitemname ,shopitemoptionname ,cd1 ,cd2 ,cd3 ,cd1_name ,cd2_name ,cd3_name ,orgsellprice ,shopitemprice
dim shopsuplycash ,shopbuyprice ,isusing ,vatinclude ,extbarcode ,imageList ,offmain ,OnlineOrgprice
dim OnlineBuycash, mwDiv ,OnlineSellcash ,regdate ,updt
	makerid = requestCheckVar(request("makerid"),32)
	barcode	  = requestCheckVar(request("barcode"),32)

editmode = FALSE

'//�����ϰ��
if barcode <> "" and not(isnull(barcode)) then
	editmode = TRUE

	itemgubun = Left(barcode,2)
	itemid	  = CLng(Mid(barcode,3,6))
	itemoption = Right(barcode,4)

	set ioffitem  = new COffShopItem
		ioffitem.FRectItemgubun = itemgubun
		ioffitem.FRectItemId = itemid
		ioffitem.FRectItemOption = itemoption
		ioffitem.GetOffOneItem

	if ioffitem.FResultCount > 0 then
		makerid = ioffitem.FOneItem.Fmakerid
		Barcode = ioffitem.FOneItem.GetBarcode
		shopitemname = ioffitem.FOneItem.Fshopitemname
		shopitemoptionname = ioffitem.FOneItem.Fshopitemoptionname
		cd1 = ioffitem.FOneItem.FCateCDL
		cd2 = ioffitem.FOneItem.FCateCDM
		cd3 = ioffitem.FOneItem.FCateCDS
		cd1_name = ioffitem.FOneItem.FCateCDLName
		cd2_name = ioffitem.FOneItem.FCateCDMName
		cd3_name = ioffitem.FOneItem.FCateCDSName
		orgsellprice = ioffitem.FOneItem.FShopItemOrgprice
		shopitemprice = ioffitem.FOneItem.Fshopitemprice
		shopsuplycash = ioffitem.FOneItem.Fshopsuplycash
		shopbuyprice = ioffitem.FOneItem.Fshopbuyprice
		ItemGubun = ioffitem.FOneItem.FItemGubun
		isusing = ioffitem.FOneItem.Fisusing
		CenterMwDiv = ioffitem.FOneItem.FCenterMwDiv
		vatinclude = ioffitem.FOneItem.Fvatinclude
		extbarcode = ioffitem.FOneItem.Fextbarcode
		imageList = ioffitem.FOneItem.FimageList
		offmain = ioffitem.FOneItem.FOffImgMain
		offList = ioffitem.FOneItem.FOffImgList
		offSmall = ioffitem.FOneItem.FOffImgSmall
		OnlineSailYn = ioffitem.FOneItem.FOnlineSailYn
		OnlineOrgprice = ioffitem.FOneItem.FOnlineOrgprice
		OnlineBuycash = ioffitem.FOneItem.FOnlineBuycash
		mwDiv = ioffitem.FOneItem.FmwDiv
		OnlineSellcash = ioffitem.FOneItem.FOnlineSellcash
		regdate = ioffitem.FOneItem.Fregdate
		updt = ioffitem.FOneItem.Fupdt

		if left(Barcode,2) <> "80" and left(Barcode,2) <> "85" then
			response.write "<script type='text/javascript'>"
			response.write "	alert('�߸��� �����Դϴ�.');"
			response.write "</script>"
			dbget.close()	:	response.end
		end if
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('�ش�Ǵ� ��ǰ�� �����ϴ�');"
		'response.write "	self.close();"
		response.write "</script>"
		dbget.close()	:	response.end
	end if

	IsOnlineItem = (itemgubun="10")

'/�űԵ��
else
	if makerid <> "" then
		CenterMwDiv = GetDefaultItemMwdivByBrand(makerid)

		shopitemprice = "0"
		orgsellprice = "0"
	end if
end if

set opartner = new CPartnerUser
    opartner.FRectDesignerID = makerid

    if makerid <> "" then
    	opartner.GetOnePartnerNUser
    else
		opartner.FResultCount = 0
	end if

set ooffontract = new COffContractInfo
    ooffontract.FRectDesignerID = makerid

    if makerid <> "" then
		ooffontract.GetPartnerOffContractInfo
	end if

function drawOffContractBrandChangeEvent(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>" onchange="ChangeBrand(this)">
     <option value='' <%if selectedId="" then response.write " selected"%>>����</option><%
   query1 = " select c.userid, c.socname_kor"
   query1 = query1 & " from [db_user].[dbo].tbl_user_c c with (nolock)"
   query1 = query1 & " join [db_shop].[dbo].tbl_shop_designer s with (nolock)"
   query1 = query1 & " 		on s.shopid='streetshop000'"
   query1 = query1 & " where c.userid = s.makerid"
   query1 = query1 & " order by c.userid"

	'response.write query1 & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open query1, dbget, adOpenForwardOnly, adLockReadOnly

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

if vatinclude = "" then vatinclude = "Y"
if isusing = "" then isusing = "Y"
'C_IS_SHOP = TRUE
%>

<script type='text/javascript'>

//�űԵ�϶� �귣�� ����
function ChangeBrand(comp){
	location.href="?makerid=" + comp.value;
}

//����
function EditItem(frm){
	var tmpitemgubuncheck = '';
	<% if editmode then %> var editmode = true; <% else %> var editmode = false; <% end if %>

	//��ǰ���� ���ð� üũ
	if (editmode){
		tmpitemgubuncheck = frm.itemgubun.value;
	}else{
		var itemgubun = document.getElementsByName("itemgubun");
		for(var i=0; i < itemgubun.length ; i++){
			if (itemgubun[i].checked){
				tmpitemgubuncheck = frm.itemgubun[i].value;
			}
		}
	}

	if (!editmode){
		if (tmpitemgubuncheck == ''){
			alert('��ǰ������ �����ϼ���.');
			return;
		}
	}

	if (frm.shopitemname.value.length<1){
		alert('��ǰ���� �Է��ϼ���.');
		frm.shopitemname.focus();
		return;
	} else {
		// Ư������ ����
		frm.shopitemname.value = frm.shopitemname.value.replace(/['"\\\|]/gi, "");
	}

	if (editmode){
	    if (frm.orgsellprice.value.length<1){
			alert('�Һ��ڰ��� �Է��ϼ���.');
			frm.orgsellprice.focus();
			return;
		}
	}

	if (frm.shopitemprice.value.length<1){
		alert('�ǸŰ��� �Է��ϼ���.');
		frm.shopitemprice.focus();
		return;
	}

	if (frm.shopsuplycash.value.length<1){
		alert('���԰��� �Է��ϼ���.');
		frm.shopsuplycash.focus();
		return;
	}

	if (editmode != true) {
        frm.itemoptioncode2.value = "";
        frm.itemoptioncode3.value = "";

		var optiont = "";
		var optionv = "";
		var optioncnt = 0;
		if (frm.useoptionyn[0].checked == true) {
			if (tmpitemgubuncheck == "80") {
				alert("�������� ����ǰ�� ���� �ɼ��� ����� �� �����ϴ�.\n\n�׽�Ʈ ������!!");
				return;
			}

			for (var i = 0; i < frm.etcOpt.length; i++) {
				// Ư������ ����
				frm.etcOpt[i].value = frm.etcOpt[i].value.replace(/['"\\\|]/gi, "");

				if (frm.etcOpt[i].value != "") {
					optioncnt = optioncnt + 1;
					var s = "0000" + optioncnt;

					optiont += (frm.etcOpt[i].value + "|");
					optionv += s.substring(s.length - 4) + "|";
				}
			}

			if (optioncnt < 2) {
				alert("�ɼ��� �ΰ� �̻��̾�� �մϴ�.");
				return;
			}
		}

		frm.itemoptioncode2.value = optionv;
        frm.itemoptioncode3.value = optiont;
	}

	if (frm.shopitemprice.value > 0){
		alert("����ǰ�� �ǸŰ��� 0���Ͽ��� �մϴ�.");
		frm.shopitemprice.focus();
		return;
	} if (editmode){
		if (frm.orgsellprice.value > 0){
			alert("����ǰ�� �Һ��ڰ� 0���Ͽ��� �մϴ�.");
			frm.orgsellprice.focus();
			return;
		}
	} if (editmode){
		if (frm.shopitemname.value.match(/^\[����ǰ\] /) == null) {
			alert("����ǰ ������ ������ �� �����ϴ�.");
			return;
		}
	} if (!editmode){
		if (frm.shopitemname.value.match(/����ǰ/) != null) {
			alert("����ǰ ������ ��ǰ�� �ڵ��Էµ˴ϴ�. ����ǰ ������ ���켼��.");
			return;
		}
	}

	if (!editmode){
		if (((frm.shopsuplycash.value!=0)||(frm.shopbuyprice.value!=0))){
			if (!confirm('!! ���忡 �Ǹ��ϴ� ��쿡�� ������ް��� �Է� �ϼž� �մϴ�. \n\n��� �Ͻðڽ��ϱ�?')){
				return;
			}
		}
	}

	if (editmode){
		if (frm.tmpoffmain.value.length<1 && frm.file1.value.length<1){
			alert('�̹����� �Է��� �ּ��� - �ʼ� �����Դϴ�.');
			frm.file1.focus();
			return;
		}
	}else{
		if (frm.file1.value.length<1){
			alert('�̹����� �Է��� �ּ��� - �ʼ� �����Դϴ�.');
			frm.file1.focus();
			return;
		}
	}

	var ret = 0;
	for (i=0; i< document.getElementsByName("centermwdiv").length; i++){
		if (document.getElementsByName("centermwdiv")[i].checked == true){
			ret = ret + 1;
		}
	}
	if (ret == 0){
		alert("���� ���� ������ ���� �ϼ���.");
		return;
	}

    if ((!frm.vatinclude[0].checked)&&(!frm.vatinclude[1].checked)){
        alert('���� ������ ���� �ϼ���.');
		frm.vatinclude[0].focus();
		return;
    }

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		if (frm.shopitemname.value.match(/����ǰ/) == null) {
			frm.shopitemname.value = "[����ǰ] " + frm.shopitemname.value;
		}

		frm.submit();
	}
}

function PopUpcheInfo(v){
	window.open("/admin/lib/popbrandinfoonly.asp?designer=" + v,"popupcheinfo","width=640 height=580 scrollbars=yes resizable=yes");
}

function editOffDesinger(shopid,designerid){
	var popwin = window.open("/admin/lib/popshopupcheinfo.asp?shopid=" + shopid + "&designer=" + designerid,"popshopupcheinfo","width=640 height=540");
	popwin.focus();
}

function TnCheckOptionYN(frm){
	if (frm.useoptionyn[0].checked == true) {
	    // �ɼǻ��
		document.all.optlist.style.display="";
		document.all.optname.style.display="none";

	} else {
	    // �ɼǾ���
		document.all.optlist.style.display="none";
		document.all.optname.style.display="";
    }
}

function InsertOptionWithGubun(ioptTypeName, ft, fv) {
	var frm = document.frmedit;

	//�ɼǰ��� �������� ������ skip ,����ɼ��ΰ�� ����
	if (fv!="0000"){
		for (i=0;i<frm.realopt.length;i++){
			if (frm.realopt[i].value==fv){
				return;
			}
		}
	}

    frm.optTypeNm.value = ioptTypeName;
	frm.elements['realopt'].options[frm.realopt.options.length] = new Option(ft, fv);
}

function popNormalOptionAdd() {
	popwin = window.open('/common/module/normalitemoptionadd.asp' ,'popNormalOptionAdd','width=540,height=260,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ���õ� �ɼ� ����
function delItemOptionAdd()
{
	var frm = document.frmedit;
	var sidx = frm.realopt.options.selectedIndex;

	if(sidx<0){
		alert("������ �ɼ��� �������ֽʿ�.");
	}else{
	    for(i=0; i<frm.realopt.options.length; i++){
    		if(frm.realopt.options[i].selected){
    			frm.realopt.options[i] = null;
    			i=i-1;
    		}
    	}

		if (frm.realopt.options.length<1){
		    frm.optTypeNm.value = '';
		}
	}
}

// ī�װ����
function editCategory(cdl,cdm,cdn){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cdn=" + cdn ;

	popwin = window.open('/common/module/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//ī�װ� ����
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

<!-- ����Ʈ ���� -->
>>����ǰ ���
<form name="frmedit" method="post" action="<%=uploadImgUrl%>/linkweb/offshop/item/itemedit_off.asp" enctype="MULTIPART/FORM-DATA" style="margin:0px;">
<input type="hidden" name="editmode" value="<%=editmode%>">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="barcode" value="<%=barcode%>">
<input type="hidden" name="offmain" value="<%=offmain%>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">
<input type="hidden" name="itemoptioncode2">
<input type="hidden" name="itemoptioncode3">
<input type="hidden" name="regtype" value="giftitem">

<input type="hidden" name="cd1" value="">
<input type="hidden" name="cd2" value="">
<input type="hidden" name="cd3" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

<% if NOT(editmode) then %>
<tr bgcolor="<%= adminColor("pink") %>">
	<td width="100" height="30">�귣��ID</td>
	<td bgcolor="#FFFFFF">
		<% drawSelectBoxDesignerwithName "imakerid",makerid  %> �ؽű� ����Ͻ� ����ǰ�� �귣�带 ������ �ּ���.
	</td>
</tr>
<% if (makerid = "") or (opartner.FResultCount < 1) then %>
<tr bgcolor="<%= adminColor("pink") %>">
	<td colspan="2" bgcolor="#FFFFFF" align="center">
		<input type="button" class="button" value="�˻�" onclick="ChangeBrand(document.frmedit.imakerid);">
	</td>
</tr>
<% end if %>
<%
end if

'// �귣�� ������ ������� �������� �ʰ�, ������ �귣�� �����ϵ���..
if makerid = "" then dbget.close() : response.write "</table>" : response.end

'// �߸��� �귣��
if opartner.FResultCount < 1 then
	response.write "<script>alert('�߸��� �귣���Դϴ�.');</script>"
	dbget.close() : response.write "</table>" : response.end
end if

%>

<tr bgcolor="<%= adminColor("pink") %>" height="30">
	<td width=100>�귣��������</td>
	<td bgcolor="#FFFFFF">
		<a href="javascript:PopUpcheInfo('<%= makerid %>');"><%= makerid %></a> (<%= opartner.FOneItem.Fsocname_kor %>,<%= opartner.FOneItem.FCompany_name %>)
	</td>
</tr>
<% if (editmode) then %>
<input type="hidden" name="itemgubun" value="<%=itemgubun%>">
<tr bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="100">��ǰ�ڵ�</td>
	<td bgcolor="#FFFFFF">
		<%= Barcode %>
		<%if left(Barcode,2) = "10" then %>
			�¶��ΰ����ǰ
		<% elseif left(Barcode,2) = "90" then %>
			�������������ǰ
		<% elseif left(Barcode,2) = "95" then %>
			���������������ǸŻ�ǰ
		<% elseif left(Barcode,2) = "85" then %>
			ON����ǰ
		<% elseif left(Barcode,2) = "80" then %>
			OFF����ǰ
		<% elseif left(Barcode,2) = "70" then %>
			�Ҹ�ǰ
		<% end if %>
		<br><font color="#AAAAAA">(85ON����ǰ, 80OFF����ǰ)</font>
	</td>
</tr>
<% else %>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100 height="30">��ǰ����</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="itemgubun" value="85" <% if itemgubun = "85" then response.write " checked" %>>ON����ǰ(85)
		<input type="radio" name="itemgubun" value="80" <% if itemgubun = "80" then response.write " checked" %> disabled>OFF����ǰ(80)
	</td>
</tr>
<% end if %>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="30">��ǰ��</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="shopitemname" value="<%= shopitemname %>" size="80" maxlength="90">
		<br>�� ��ǰ�� "[����ǰ]" ������ �ڵ����� �ٽ��ϴ�.
	</td>
</tr>
<% if NOT(editmode) then %>
<tr bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td height="30">�ɼǱ���</td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="useoptionyn" value="Y" onClick="TnCheckOptionYN(this.form);" disabled>�ɼǻ����</label>&nbsp;&nbsp;
		<label><input type="radio" name="useoptionyn" value="N" onClick="TnCheckOptionYN(this.form);" checked>�ɼǻ�����</label>
	</td>
</tr>
<!----- ���� �ɼ� DIV ----->
<tr bgcolor="<%= adminColor("tabletop") %>" id="optname" height="30">
    <td height="30">�ɼǸ�</td>
  	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="shopitemoptionname" value="<%= shopitemoptionname %>" size="40" maxlength="40">
  	</td>
</tr>
<!----- ���� �ɼ� DIV ----->
<tr bgcolor="<%= adminColor("tabletop") %>" id="optlist" style="display:none" height="30">
    <td height="30">�ɼ� ����</td>
  	<td bgcolor="#FFFFFF" align="left">

		<table width="440" border="0" cellspacing="1" cellpadding="2" align="left" class="a"  bgcolor="#3d3d3d" >
		<% for i = 1 to 10 %>
		<tr bgcolor="#FFFFFF" align="center">
			<td>�ɼǸ� <%= i %> </td>
			<td align="center"><input type="text" class="text" name="etcOpt" size="20" maxlength="20"></td>
		</tr>
		<% next %>
		</table>

  	</td>
</tr>
<% else %>
<tr bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>�ɼǱ���</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="shopitemoptionname" value="<%= shopitemoptionname %>" size="40" maxlength="40">
	</td>
</tr>
<% end if %>
<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td>�Һ��ڰ�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="orgsellprice" value="<%= orgsellprice %>" size=8 maxlength=9 readonly>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td>�ǸŰ�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="shopitemprice" value="<%= shopitemprice %>" size=8 maxlength=9 readonly>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td>���԰�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="shopsuplycash" value="<%= shopsuplycash %>" size=8 maxlength=9 class="input_right"> �ؼ��� ������ ���� �������
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td>������ް�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="shopbuyprice" value="<%= shopbuyprice %>" size=8 maxlength=9 class="input_right" > �ؼ��� ������ ���� �������
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>�������</td>
	<td bgcolor="#FFFFFF">
		<% if isusing = "Y" then %>
		<input type=radio name=isusing value="Y" checked >�����
		<input type=radio name=isusing value="N">������
		<% else %>
		<input type=radio name=isusing value="Y"  >�����
		<input type=radio name=isusing value="N" checked >������
		<% end if %>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>���͸��Ա���</td>
	<td bgcolor="#FFFFFF">
		<%
		' �űԵ�Ͻÿ��� ������ ��Ź���� ����.	2023.06.23 �̹����̻�� ��û
		if not(editmode) then
		%>
			<input type="radio" name="centermwdiv" value="W" checked >��Ź
		<% else %>
			<input type="radio" name="centermwdiv" value="W" <%= ChkIIF(centermwdiv="W","checked","") %> >��Ź
			<input type="radio" name="centermwdiv" value="M" <%= ChkIIF(centermwdiv="M","checked","") %> >����
		<% end if %>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>��������</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="vatinclude" value="Y" <%= ChkIIF(vatinclude = "Y","checked","") %>  >����
		<input type="radio" name="vatinclude" value="N" <%= ChkIIF(vatinclude = "N","checked","") %> > <font color="<%= ChkIIF(vatinclude = "N","blue","#000000") %>">�鼼</font>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>�̹���</td>
	<td bgcolor="#FFFFFF">
		<% if IsOnlineItem then %>
			<img src="<%= imageList %>" width="50" height="50">
		<% else %>
			<input type="file" name="file1" class="button" size=20 >
			<Br>�� �⺻ �̹����� �ݵ�� 400x400 , jpg ���Ϸ� �÷��ֽñ� �ٶ��ϴ�.
			<Br>�� 400x400 �̹����� ���� �Ͻø�, �ڵ����� 100x100 , 50x50 �� ���� �˴ϴ�.
			<input type="hidden" name="tmpoffmain" value="<%= offmain %>">
   				<% IF offmain <> "" THEN %>
	   				<BR><img src="<%=offmain%>" border="0" width=400 height=400> 400x400
   				<% END IF %>
   				<% if offlist <> "" then %>
   					<BR><img src="<%=offlist%>" border="0" width=100 height=100> 100x100
   				<% end if %>
   				<% if offsmall <> "" then %>
   					<BR><img src="<%=offsmall%>" border="0" width=50 height=50> 50x50
   				<% end if %>
		<% end if %>
	</td>
</tr>
<% if editmode then %>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>�����</td>
	<td bgcolor="#FFFFFF"><%= regdate %></td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>����������</td>
	<td bgcolor="#FFFFFF"><%= updt %></td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align=center>
		<input type="button" class="button" value="<% if editmode then %>����<% else %>�ű�����<% end if %>" onclick="EditItem(frmedit)">
	</td>
</tr>
</table>
</form>

<%
set opartner = Nothing
set ooffontract = Nothing
set ioffitem = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->