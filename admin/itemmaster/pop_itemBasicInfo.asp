<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ����
' History : ������ ����
'			2016.07.06 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, oitem, oitemvideo
dim makerid, rentalItemFlag

itemid = requestCheckvar(request("itemid"),10)
makerid = requestCheckvar(request("makerid"),32)
menupos = requestCheckvar(request("menupos"),10)
if (itemid = "") then
    response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if

set oitem = new CItem

oitem.FRectItemID = itemid
oitem.GetOneItem

Set oitemvideo = New CItem
oitemvideo.FRectItemId = itemid
oitemvideo.FRectItemVideoGubun = "video1"
oitemvideo.GetItemContentsVideo

''������ǰ ��� ����
dim strItemRelation
strItemRelation = GetItemRelationStr(itemid)

'���ϸ���
dim sailmargine, orgmargine, margine

''����
if oitem.FOneItem.Fsailprice<>0 then
	sailmargine = 100-CLng(oitem.FOneItem.Fsailsuplycash/oitem.FOneItem.Fsailprice*100*100)/100
else
	sailmargine = 0
end if

if oitem.FOneItem.Forgprice<>0 then
	orgmargine = 100-CLng(oitem.FOneItem.Forgsuplycash/oitem.FOneItem.Forgprice*100*100)/100
else
	orgmargine = 0
end if

if oitem.FOneItem.Fsellcash<>0 then
	margine = 100-CLng(oitem.FOneItem.Fbuycash/oitem.FOneItem.Fsellcash*100*100)/100
else
	margine = 0
end if

'// ��Ż ��ǰ�� �ϴ� �׽�Ʈ�� ��Ź ������ ������
If C_ADMIN_AUTH Then
	rentalItemFlag = true
Else
	rentalItemFlag = true
End If
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript' src="/js/ckeditor/ckeditor.js"></script>
<script type="text/javascript">
$(function(){
	// �ε��� ��ǰ�Ӽ� ���� ���
	printItemAttribute();
});

function printItemAttribute() {
	var arrDispCd="";
	$("input[name='catecode']").each(function(i){
		if(i>0) arrDispCd += ",";
		arrDispCd += $(this).val();
	});
	$.ajax({
		url: "/common/module/act_ItemAttribSelect.asp?itemid=<%=itemid%>&arrDispCate="+arrDispCd,
		cache: false,
		success: function(message) {
			$("#lyrItemAttribAdd").empty().append(message);
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

function UseTemplate() {
	window.open("/common/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
}

function popMultiLangEdit(iid) {
	window.open("/common/item/pop_MultiLangItemCont.asp?itemid="+iid+"&lang=EN", "multiLang_win", "width=600, height=500, scrollbars=yes, resizable=yes");
}

// ============================================================================
// ī�װ����
function editCategory(cdl,cdm,cds){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cds=" + cds ;

	popwin = window.open('/common/module/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function setCategory(cd1,cd2,cd3,cd1_name,cd2_name,cd3_name){
	var frm = document.itemreg;
	frm.cd1.value = cd1;
	frm.cd2.value = cd2;
	frm.cd3.value = cd3;
	frm.cd1_name.value = cd1_name;
	frm.cd2_name.value = cd2_name;
	frm.cd3_name.value = cd3_name;
}


// ============================================================================
// �����ϱ�
function SubmitSave() {
	if (itemreg.designerid.value == ""){
		alert("��ü�� �����ϼ���.");
		itemreg.designer.focus();
		return;
	}

	if (!$("input[name='isDefault'][value='y']").length&&$("input[name='isDefault']").length){
		alert("[�⺻] ���� ī�װ��� �����ϼ���.\n�� [�߰�] ���� ī�װ��� ���� �� �����ϴ�.");
		return;
	}

	// ī�װ� �������� �˻�
	if(tbl_Category.rows.length>0)	{
		if(tbl_Category.rows.length>1)	{
			var chk=0;
			for(l=0;l<document.all.cate_div.length;l++)	{
				if(document.all.cate_div[l].value=="D") chk++;
			}
			if(chk==0) {
				alert("ī�װ��� �⺻ ī�װ��� �������ּ���.\n�ر⺻ ī�װ��� �ʼ��׸��Դϴ�.");
				return;
			} else if(chk>1) {
				alert("ī�װ��� �⺻ ī�װ��� �Ѱ��� �������ּ���.");
				return;
			}
		}
		else {
			if(document.all.cate_div.length){
				if(document.all.cate_div[0].value!="D") {
					alert("ī�װ��� �⺻ ī�װ��� �������ּ���.\n�ر⺻ ī�װ��� �ʼ��׸��Դϴ�.");
					return;
				}
			} else {
				if(document.all.cate_div.value!="D") {
					alert("ī�װ��� �⺻ ī�װ��� �������ּ���.\n�ر⺻ ī�װ��� �ʼ��׸��Դϴ�.");
					return;
				}
			}
		}
	} else {
		alert("ī�װ��� �������ּ���.");
		return;
	}

    if (validate(itemreg)==false) {
        return;
    }
    
    //��ü��۸� �ֹ����� ����.
    <% if oitem.FOneItem.Fmwdiv <> "U" then %>
    if (itemreg.itemdiv[1].checked){
        alert('�ֹ� ���ۻ�ǰ�� ��ü����ΰ�츸 �����մϴ�.');
        itemreg.itemdiv[0].focus();
        return;
    }
    <%else%>//����,��Ź�� �ܵ�(����) �ֹ����� ����
    	if(itemreg.reserveItemTp[1].checked){
    		if(!confirm("�ܵ�(����)���Ż�ǰ�� �ٸ� ��ǰ�� ���� ���Ű� �Ұ��մϴ�.\n�ܵ� ���Ż�ǰ���� �����Ͻðڽ��ϱ�?")){
				itemreg.reserveItemTp[0].focus();
				return;
			};
    	}
    <% end if %>
    
    //��ǰ�� ����üũ �߰� 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("��ǰ���� �ִ� 64byte ���Ϸ� �Է����ּ���.(�ѱ�32�� �Ǵ� ����64��)");
		itemreg.itemname.focus();
		return;
	}

	//��ǰ ǰ������
    if (!itemreg.infoDiv.value){
        alert('��ǰ�� �ش��ϴ� ǰ���� �������ֽʽÿ�.');
        itemreg.infoDiv.focus();
        return;
    } else if(itemreg.infoDiv.value=="35") {
    	if(!itemreg.itemsource.value) {
	        alert('��ǰ�� ������ �Է����ּ���.');
	        itemreg.itemsource.focus();
	        return;
    	}
    	if(!itemreg.itemsize.value) {
	        alert('��ǰ�� ũ�⸦ �Է����ּ���.');
	        itemreg.itemsize.focus();
	        return;
    	}
    }

	//������������.
    if (itemreg.safetyYn[0].checked){
  		if($("#real_safetynum").val() == ""){
  			alert("�������������� �����ϰ� ������ȣ�� �Է��� �߰���ư�� Ŭ�����ּ���.");
  			return;
  		}
    }

    if(confirm("��ǰ�� �ø��ðڽ��ϱ�?") == true){
		<% ''�������� api�� ��ȸ �� ���� ������ db���� �� ����idx�� �޾� ���� %>
		if(itemreg.safetyYn[0].checked) {
			$("#real_safetyidx").val(jsCallAPIsafety($("#real_safetynum").val(),"o",$("#real_safetydiv").val()));
		}

        itemreg.submit();
    }

}

function getByteLength(inputValue) {
     var byteLength = 0;
     for (var inx = 0; inx < inputValue.length; inx++) {
         var oneChar = escape(inputValue.charAt(inx));
         if ( oneChar.length == 1 ) {
             byteLength ++;
         } else if (oneChar.indexOf("%u") != -1) {
             byteLength += 2;
         } else if (oneChar.indexOf("%") != -1) {
             byteLength += oneChar.length/3;
         }
     }
     return byteLength;
}

// ============================================================================
	// ī�°� ���� �˾�
	function popCateSelect(iid){
		var popwin = window.open("/common/module/NewCategorySelect.asp?iid=" + iid, "popCateSel","width=700,height=400,scrollbars=yes,resizable=yes");
        popwin.focus();
	}

	// ����ī�װ� ���� �˾�
	function popDispCateSelect(){
		var designerid = document.all.itemreg.makerid.value;
		if(designerid == ""){
			alert("��ü�� �����ϼ���.");
			return;
		}
		
		var dCnt = $("input[name='isDefault'][value='y']").length;
		$.ajax({
			url: "/common/module/act_DispCategorySelect.asp?designerid="+designerid+"&isDft="+dCnt,
			cache: false,
			success: function(message) {
				$("#lyrDispCateAdd").empty().append(message).fadeIn();
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	// �˾����� ���� ī�װ� �߰�
	function addCateItem(lcd,lnm,mcd,mnm,scd,snm,div)
	{
		// ������ ���� �ߺ� ī�װ� ���� �˻� - �ö�� ��������� ������;
		if(tbl_Category.rows.length>0)	{
			if(tbl_Category.rows.length>1)	{
				for(l=0;l<document.all.cate_div.length;l++)	{
				    if (!((document.all.cate_large[l].value=="110")&&(document.all.cate_mid[l].value=="060"))){
    					if((document.all.cate_large[l].value==lcd)&&(document.all.cate_mid[l].value==mcd)) {
    						alert("���� �ߺз��� �̹� ������ ī�װ��� �ֽ��ϴ�.\n���� ī�װ��� �����ϰ� �ٽ� �������ּ���.");
    						return;
    					}
    				}
				}
			}
			else {
			    if (!((document.all.cate_large.value=="110")&&(document.all.cate_mid.value=="060"))){
    				if((document.all.cate_large.value==lcd)&&(document.all.cate_mid.value==mcd)) {
    					alert("���� �ߺз��� �̹� ������ ī�װ��� �ֽ��ϴ�.\n�ر��� ī�װ��� �����ϰ� �ٽ� �������ּ���.");
    					return;
    				}
    			}
			}
		}
		
		// ���߰�
		var oRow = tbl_Category.insertRow();
		oRow.onmouseover=function(){tbl_Category.clickedRowIndex=this.rowIndex};

		// ���߰� (����,ī�װ�,������ư)
		var oCell1 = oRow.insertCell();		
		var oCell2 = oRow.insertCell();
		var oCell3 = oRow.insertCell();

		if(div=="D") {
			oCell1.innerHTML = "<font color='darkred'><b>[�⺻]<b></font><input type='hidden' name='cate_div' value='D'>";
		} else {
			oCell1.innerHTML = "<font color='darkblue'>[�߰�]</font><input type='hidden' name='cate_div' value='A'>";
		}
		oCell2.innerHTML = lnm + " >> " + mnm + " >> " + snm
					+ "<input type='hidden' name='cate_large' value='" + lcd + "'>"
					+ "<input type='hidden' name='cate_mid' value='" + mcd + "'>"
					+ "<input type='hidden' name='cate_small' value='" + scd + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delCateItem()' align=absmiddle>";
	}

	// ���̾�� ����ī�װ� �߰�
	function addDispCateItem(dcd,cnm,div,dpt) {
		// ������ ���� �ߺ� ī�װ� ���� �˻�
		if(tbl_DispCate.rows.length>0)	{
			if(tbl_DispCate.rows.length>1)	{
				for(l=0;l<document.all.isDefault.length;l++)	{
				    if((document.all.catecode[l].value==dcd)) {
						alert("�̹� ������ ���� ī�װ��� �ֽ��ϴ�..");
						return;
					}
				}
			}
			else {
			    if((document.all.catecode.value==dcd)) {
					alert("�̹� ������ ���� ī�װ��� �ֽ��ϴ�..");
					return;
				}
			}
		}
		
		// ���߰�
		var oRow = tbl_DispCate.insertRow();
		oRow.onmouseover=function(){tbl_DispCate.clickedRowIndex=this.rowIndex};

		// ���߰� (����,ī�װ�,������ư)
		var oCell1 = oRow.insertCell();		
		var oCell2 = oRow.insertCell();
		var oCell3 = oRow.insertCell();

		if(div=="y") {
			oCell1.innerHTML = "<font color='darkred'><b>[�⺻]<b></font><input type='hidden' name='isDefault' value='y'>";
		} else {
			oCell1.innerHTML = "<font color='darkblue'>[�߰�]</font><input type='hidden' name='isDefault' value='n'>";
		}
		$(cnm).each(function(i){
			if(dpt>i) {
				if(i>0) oCell2.innerHTML += " >> ";
				oCell2.innerHTML += $(this).text();
			}
		});
		oCell2.innerHTML += "<input type='hidden' name='catecode' value='" + dcd + "'>";
		oCell2.innerHTML += "<input type='hidden' name='catedepth' value='" + dpt + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delDispCateItem()' align=absmiddle>";
		$("#lyrDispCateAdd").fadeOut();

		//��ǰ�Ӽ� ���
		printItemAttribute();
	}

	// ���� ī�װ� ����
	function delCateItem() {
		if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?"))
			tbl_Category.deleteRow(tbl_Category.clickedRowIndex);
	}

	// ���� ����ī�װ� ����
	function delDispCateItem() {
		if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

			//��ǰ�Ӽ� ���
			printItemAttribute();
		}
	}

function checkItemDiv(comp){
    var frm = comp.form;
    
    if (comp.name=="itemdiv"){
        if (frm.itemdiv[1].checked){
            frm.reqMsg.disabled=false;
        }else{
            //frm.reqMsg.checked=false;
            frm.reqMsg.disabled=true;
        }
    }
    
    //�ֹ����� ��ǰ�ΰ��.
    if (frm.itemdiv[1].checked){
        if (frm.reqMsg.checked){
            frm.itemdiv[1].value="06";
        }else{
            frm.itemdiv[1].value="16";
        }
    }

    //��Ż ��ǰ�ΰ��.
    if (frm.itemdiv[7].checked){
		frm.reserveItemTp[1].checked = true;
    }	
}

//ǰ�� ���� / ǰ�񳻿� ǥ��
function chgInfoDiv(v) {
	$("#itemInfoList").empty();

	if(v=="") {
		$("#itemInfoCont").hide();
	} else {
		$("#itemInfoCont").show();

		var str = $.ajax({
			type: "POST",
			url: "act_itemInfoDivForm.asp",
			data: "itemid=<%=itemid%>&ifdv="+v,
			dataType: "html",
			async: false
		}).responseText;
	
		if(str!="") {
			$("#itemInfoList").html(str);
		}
	}

	if(v=="35") {
		$("#lyItemSrc").show();
		$("#lyItemSize").show();
	} else {
		$("#lyItemSrc").hide();
		$("#lyItemSize").hide();
	}

	// ��������üũ. ���ȹ�
	jsSafetyCheck('','');
}

//�ܼ� ���� ������
function chgInfoChk(fm) {
	$(fm).parent().parent().find('[name="infoChk"]').val($(fm).val());
}

//���� ���� ������
function chgInfoSel(fm) {
	$(fm).parent().parent().find('[name="infoChk"]').val($(fm).val());
	$(fm).parent().parent().find('[name="infoCont"]').val($(fm).attr("msg"));

	if($(fm).val()=="Y") {
		$(fm).parent().parent().find('[name="infoCont"]').removeAttr("readonly");
		$(fm).parent().parent().find('[name="infoCont"]').removeClass("text_ro");
		$(fm).parent().parent().find('[name="infoCont"]').addClass("text");
	} else {
		$(fm).parent().parent().find('[name="infoCont"]').attr("readonly", true);
		$(fm).parent().parent().find('[name="infoCont"]').addClass("text_ro");
	}
}

//��ǰ���� ���� ������ ���� ǥ��
function jsSetArea(iValue){ 
	var i;
	for(i=0;i<=4;i++) { 
 		eval("document.all.dvArea"+i).style.display = "none";
	}
 	eval("document.all.dvArea"+iValue).style.display = "";
} 

function jsCallAPIsafety(certnum,isSave,safetydiv){
	var returnmsg = "";
	$.ajax({
		url: "/admin/itemmaster/safety_api_auth_proc.asp?itemid=<%=itemid%>&issave="+isSave+"&certnum="+certnum+"&safetydiv="+safetydiv+"&statusmode=real",
		cache: false,
		async: false,
		success: function(message)
		{
			returnmsg = message;
		}
	});
	return returnmsg;
}

//����ī�װ�(����������)�� ���� alert �޼���.
function jsAlertCatecodeSafety(){
	var auth_go_catecode = "";
	if(typeof itemreg.catecode != "undefined"){
		if(itemreg.catecode.length == undefined){
			auth_go_catecode = itemreg.catecode.value;
		}else{
			for(si=0; si<itemreg.catecode.length; si++){
				auth_go_catecode = auth_go_catecode + itemreg.catecode[si].value + ",";
			}
		}
		
		if(auth_go_catecode != ""){
			$("#auth_go_catecode").val(auth_go_catecode);
			
			var ccode = $("#auth_go_catecode").val();
			$.ajax({
					url: "/common/item/catecode_safety_info_ajax.asp?catecode="+ccode,
					cache: false,
					async: false,
					success: function(msgc)
					{
						if(msgc != ""){
							msgc = msgc.replace(/br/gi,"\n");
							alert(msgc);
						}
					}
			});
		}
	}else{
		alert("����ī�װ��� �������ּ���.");
	}
}

//�߰��� �������� ����Ʈ ���� ����
function jsSafetyDivListDel(listnum){
	var realvalue = $("#real_safetydiv").val();
	var jbSplit = $("#real_safetydiv").val().split(",");
	var jbSplitnum = $("#real_safetynum").val().split(",");
	var resultDiv = "";
	var resultNum = "";
	var del_safetynum = "";
	var del_safetydiv = "";
	
	for(var i in jbSplit){
		if(jbSplit[i] != listnum){
			resultDiv = resultDiv + jbSplit[i] + ",";
			resultNum = resultNum + jbSplitnum[i] + ",";
		}else{
			del_safetynum = jbSplitnum[i];
			del_safetydiv = jbSplit[i];
		}
	}
	
	if(resultDiv.substr(resultDiv.length-1, 1) == ","){
		resultDiv = resultDiv.substr(0, resultDiv.length-1);
		resultNum = resultNum.substr(0, resultNum.length-1);
	}
	$("#real_safetydiv").val(resultDiv);
	$("#real_safetynum").val(resultNum);
	
	$("#l"+listnum+"").remove();
	
	var tmp_num = $("#real_safetynum_delete").val();
	var tmp_div = $("#real_safetydiv_delete").val();
	if(tmp_num == ""){
		$("#real_safetynum_delete").val(del_safetynum);
		$("#real_safetydiv_delete").val(del_safetydiv);
	}else{
		$("#real_safetynum_delete").val(tmp_num + "," + del_safetynum);
		$("#real_safetydiv_delete").val(tmp_div + "," + del_safetydiv);
	}
}

// �귣��ID ����
function fnChangeBrandID() {
//
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>��ǰ �⺻���� ����</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<br><b>��ϵ� ��ǰ�� �⺻������ �����մϴ�.</b>
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>

<p>

<form name="itemreg" method="post" action="/admin/itemmaster/itemmodify_Process.asp" onsubmit="return false;" style="margin:0;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="ItemBasicInfo">
<input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
<input type="hidden" name="designerid" value="<%= oitem.FOneItem.Fmakerid %>">
<input type="hidden" name="orgprice" value="<%= oitem.FOneItem.Forgprice %>">
<input type="hidden" name="orgsuplycash" value="<%= oitem.FOneItem.Forgsuplycash %>">
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left">�⺻����</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�ڵ� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<%= oitem.FOneItem.Fitemid %>
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="�̸�����" class="button" onclick="window.open('http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FOneItem.Fitemid %>');">
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">���/�Ǹ��� :</td>
	<td width="35%" bgcolor="#FFFFFF">
		��ǰ����� : <%= oitem.FOneItem.FRegDate %>
		<%
			if oitem.FOneItem.FsellSTDate<>"" then
				Response.Write "<br />�ǸŽ����� : " & oitem.FOneItem.FsellSTDate
			elseif oitem.FOneItem.Fsellreservedate<>"" then
				Response.Write "<br />�Ǹſ����� : " & oitem.FOneItem.Fsellreservedate
			end if
		%>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="���� �귣��ID">�귣��ID :</td> 
	<td width="35%" bgcolor="#FFFFFF">
		<% 'NewDrawSelectBoxDesignerChangeMargin "makerid", oitem.FOneItem.Fmakerid, "marginData", "fnChangeBrandID" %>
		<% drawSelectBoxDesignerWithName "makerid", oitem.FOneItem.Fmakerid %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF" title="����Ʈ���� ǥ�õ� �귣��(������ ���� �귣�� �����)">ǥ�� �귣�� :</td>
	<td width="35%" bgcolor="#FFFFFF">
	<%
		drawSelectBoxDesignerWithName "frontMakerid", oitem.FOneItem.FfrontMakerid

		'ǥ�ú귣�� ���� ��ư
		response.Write "&nbsp;<input type=""button"" class=""button"" value=""����"" onClick=""this.form.frontMakerid.value='';"">"
	%>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�귣��� :</td>
	<td bgcolor="#FFFFFF" colspan="3" id="txtBrandName"><%=oitem.FOneItem.Fbrandname%></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemname" maxlength="64" size="60" class="text" id="[on,off,off,off][��ǰ��]" value="<%= Replace(oitem.FOneItem.Fitemname,"""","&quot;") %>">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ǰ�� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemnameEng" maxlength="64" size="60" class="text_ro" readonly id="[off,off,off,off][������ǰ��]" value="<%= Replace(oitem.FOneItem.FitemnameEng,"""","&quot;") %>">&nbsp;
		<input type="button" value="�ٱ��� ���� <%=chkIIF(oitem.FOneItem.FitemnameEng="" or isnull(oitem.FOneItem.FitemnameEng),"���","����")%>" class="button" onclick="popMultiLangEdit(<%= oitem.FOneItem.Fitemid %>)" />
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="���/���� ���� ���� ī�װ�" style="cursor:help;">���� ī�װ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td><%=getCategoryInfo(oitem.FOneItem.Fitemid)%></td>
			<td valign="bottom"><input type="button" value="�߰�" class="button" onClick="popCateSelect('<%=oitem.FOneItem.Fitemid%>')"></td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="����Ʈ�� ������ ī�װ�" style="cursor:help;">���� ī�װ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td><%=getDispCategory(oitem.FOneItem.Fitemid)%></td>
			<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<label><input type="radio" name="itemdiv" value="01" <%=chkIIF(oitem.FOneItem.Fitemdiv="01","checked","")%> onClick="this.form.requireMakeDay.value=0;document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">�Ϲݻ�ǰ</label>
		<br>
		<label><input type="radio" name="itemdiv" value="<%= oitem.FOneItem.Fitemdiv %>" <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","checked","")%> onClick="document.getElementById('lyRequre').style.display='block';checkItemDiv(this);">�ֹ� ���ۻ�ǰ</label>
		<input type="checkbox" name="reqMsg" value="10" <%=chkIIF(oitem.FOneItem.Fitemdiv="06","checked","")%> <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","","disabled")%> onClick="checkItemDiv(this);">�ֹ����� ���� �ʿ�<font color=red>(�ֹ��� �̴ϼȵ� ���۹����� �ʿ��Ѱ�� üũ)</font>
		<br>
		<label><input type="radio" name="itemdiv" value="08" <%=chkIIF(oitem.FOneItem.Fitemdiv="08","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">Ƽ�ϻ�ǰ</label>
		<label><input type="radio" name="itemdiv" value="09" <%=chkIIF(oitem.FOneItem.Fitemdiv="09","checked","")%> >Present��ǰ</label>
		<label><input type="radio" name="itemdiv" value="18" <%=chkIIF(oitem.FOneItem.Fitemdiv="18","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">�����ǰ</label>

		<!--% if oitem.FOneItem.Fitemdiv ="07" then %--> <!-- 2014������ �ܵ����� ��ǰ > reserveItemTp=1 / ����� ��������(ȸ���� ���� ����) -->
			<label><input type="radio" name="itemdiv" value="07" <%=chkIIF(oitem.FOneItem.Fitemdiv="07","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">�������ѻ�ǰ</label>
		<!--% end if %-->
		<% if oitem.FOneItem.Fitemdiv ="82" then %>
			<label><input type="radio" name="itemdiv" value="82" <%=chkIIF(oitem.FOneItem.Fitemdiv="82","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">���ϸ����� ��ǰ</label>
		<% end if %>

		<label><input type="radio" name="itemdiv" value="75" <%=chkIIF(oitem.FOneItem.Fitemdiv="75","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">���ⱸ����ǰ</label>

		<% If rentalItemFlag Then %>
			<label><input type="radio" name="itemdiv" value="30" <%=chkIIF(oitem.FOneItem.Fitemdiv="30","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">��Ż��ǰ<font color=red>(��Ż��ǰ�� �ݵ�� �ܵ�(����)���Ż�ǰ���� ����ϼž� �մϴ�.)</font></label>
		<% End If %>
		<label><input type="radio" name="itemdiv" value="23" <%=chkIIF(oitem.FOneItem.Fitemdiv="23","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">B2B��ǰ</label>
		<label><input type="radio" name="itemdiv" value="17" <%=chkIIF(oitem.FOneItem.Fitemdiv="17","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">�����������ǰ</label>
		<label><input type="radio" name="itemdiv" value="11" <%=chkIIF(oitem.FOneItem.Fitemdiv="11","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">��ǰ�ǻ�ǰ</label>
	</td>
	<td  bgcolor="#FFFFFF">
	    <div id="lyRequre" style="<%=chkIIF((oitem.FOneItem.Fitemdiv ="06") or (oitem.FOneItem.Fitemdiv ="16"),"","display:none;")%>padding-left:22px;">
		�������ۼҿ��� <input type="text" name="requireMakeDay" value="<%=oitem.FOneItem.FrequireMakeDay%>" size="2" class="text" id="[off,on,off,off][�������ۼҿ���]">��
		<font color="red">(��ǰ�߼��� ��ǰ���� �Ⱓ)</font>
		</div>
	</td>
</tr>
<!-- �����ߴ� 2017.10.17 ������(����ȣ ��ȹ) -->
<!--<tr>
	<td height="30" width="15%" bgcolor="#DDDDFF"> �귣�� �������� : </td>
	<td bgcolor="#FFFFFF" colspan="3">	 
	<span style="margin-right:10px;"><input type="checkbox" name="chkMsg" value="Y" <%if oitem.FOneItem.FaddMsg ="Y" then%>checked<%end if%>> �޽��� ÷��</span>
	<span  style="margin-right:10px;"><input type="checkbox" name="chkCarve"  value="Y" <%if oitem.FOneItem.FaddCarve ="Y" then%>checked<%end if%>> ���� ����</span>
	<span  style="margin-right:10px;"><input type="checkbox"  name="chkBox"  value="Y" <%if oitem.FOneItem.FaddBox ="Y" then%>checked<%end if%>>  �ڽ�����</span>
	<span style="margin-right:10px;"><input type="checkbox"  name="chkSet"  value="Y"<%if oitem.FOneItem.FaddSet ="Y" then%>checked<%end if%>>  ������Ʈ</span>
	<span  style="margin-right:10px;"><input type="checkbox"  name="chkCustom"  value="Y" <%if oitem.FOneItem.FaddCustom ="Y" then%>checked<%end if%>>  �ֹ����� </span>
	</td>
</tr>-->
<!---// ---------------------->
<!--% if (oitem.FOneItem.IsReserveOnlyItem) then %-->
<!-- ������ �ý����� only 2012/03/26 �߰�-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ܵ�(����)���� :</td>
	<td bgcolor="#FFFFFF" colspan="3"> 
	    <%if isNull(oitem.FOneItem.FreserveItemTp) then oitem.FOneItem.FreserveItemTp=0 %>
	    <label><input type="radio" name="reserveItemTp" value="0" <%=chkIIF(oitem.FOneItem.FreserveItemTp="0" And oitem.FOneItem.Fitemdiv<>"30","checked","")%>>�Ϲ�</label>
		<label><input type="radio" name="reserveItemTp" value="1" <%=chkIIF(oitem.FOneItem.FreserveItemTp="1" or oitem.FOneItem.Fitemdiv="30","checked","")%>>�ܵ�(����)���Ż�ǰ</label>
	</td>
</tr>
<!--% end if %-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ٹ����� ���� :</td>
	<td bgcolor="#FFFFFF" colspan="3"> 
		<label><input type="radio" name="tenOnlyYn" value="Y" <%=chkIIF(oitem.FOneItem.FtenOnlyYn="Y","checked","")%>>������ǰ</label>
		<label><input type="radio" name="tenOnlyYn" value="N" <%=chkIIF(oitem.FOneItem.FtenOnlyYn="N","checked","")%>>�Ϲݻ�ǰ</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">���� ���� ���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="adultType" value="0" <%=chkIIF(oitem.FOneItem.FadultType=0,"checked","")%>>��ü����</label>
		<label><input type="radio" name="adultType" value="1" <%=chkIIF(oitem.FOneItem.FadultType=1,"checked","")%>>�̼��� ��ȸ�Ұ�</label>
		<label><input type="radio" name="adultType" value="2" <%=chkIIF(oitem.FOneItem.FadultType=2,"checked","")%>>���Ž� ��������</label>
	</td>
</tr>
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ ���� ��ǰ :</td>
	<td width="85%" bgcolor="#FFFFFF" colspan="3"> 
		<label><input type="radio" name="availPayType" value="9" <%=chkIIF(oitem.FOneItem.FavailPayType="9","checked","")%>>������</label>
		<label><input type="radio" name="availPayType" value="8" <%=chkIIF(oitem.FOneItem.FavailPayType="8","checked","")%>>����Ʈ������</label>
		<label><input type="radio" name="availPayType" value="0" <%=chkIIF(oitem.FOneItem.FavailPayType="0","checked","")%>>�Ϲ�</label> 
	</td>
</tr>-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰī�� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="designercomment" size="60" maxlength="128" class="text" id="[off,off,off,off][��ǰī��]" value="<%= Replace(oitem.FOneItem.Fdesignercomment,"""","&quot;") %>">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemWeight" maxlength="12" size="8" class="text" id="[on,off,off,off][��ǰ����]" style="text-align:right" value="<%= oitem.FOneItem.FitemWeight %>">g &nbsp;(�׷������� �Է�, ex:1.5kg�� 1500) / �ؿܹ�۽� ��ۺ� ������ ���� ���̹Ƿ� ��Ȯ�� �Է�.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
	<td bgcolor="#FFFFFF" colspan="3"> 
		 <p> 
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="0" <%if isNull(oitem.FOneItem.Fsourcekind) or oitem.FOneItem.Fsourcekind="0" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ��ǰ ��</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="1" <%if oitem.FOneItem.Fsourcekind="1" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ����깰</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="2" <%if oitem.FOneItem.Fsourcekind="2" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ���깰</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="3" <%if oitem.FOneItem.Fsourcekind="3" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ��깰</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="4" <%if oitem.FOneItem.Fsourcekind="4" then%>checked<%end if%> onClick="jsSetArea(this.value);"> ����갡��ǰ</span>
	  </p>
	  <p><input type="text" name="sourcearea" maxlength="64" size="64" class="text" id="[on,off,off,off][������]"  value="<%= oitem.FOneItem.Fsourcearea %>"/></p>
	  <div id="dvArea0" style="display:<%if isNull(oitem.FOneItem.Fsourcekind) or oitem.FOneItem.Fsourcekind="0" then%>block<%else%>none<%end if%>;">
	  <p><strong>ex: �ѱ�, �߱�, �߱�OEM, �Ϻ� �� </strong></BR>
	   - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea1" style="display:<%if oitem.FOneItem.Fsourcekind ="1" then%>block<%else%>none<%end if%>;">
	  <p><strong>������ :</strong> ����, ������ �Ǵ� �á�����, �á�����(���ѹα�, �ѱ�X)  <span style="margin-right:10px;">ex. ��(����)</span></BR>
	   <strong>���Ի� :</strong> ������� ���Ա����� <span style="margin-right:10px;">ex. ����(�߱���)</span></BR>
	   - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea2" style="display:<%if oitem.FOneItem.Fsourcekind ="2" then%>block<%else%>none<%end if%>;">
	  <p><strong>������ :</strong> ����,������ �Ǵ� �����ػ�(��� ���깰�� �á����� ����)   <span style="margin-right:10px;">ex. ��ġ(����), ��¡��(�����ػ�)</span> </BR>
	  	<strong>����� :</strong> ����� �Ǵ� �����(�ؿ���)   <span style="margin-right:10px;">ex. ��ġ[�����(�뼭��)]</span> </BR>
	    <strong>���Ի� :</strong> ������� ���Ա����� <span style="margin-right:10px;">ex. ���(�߱���)</span></BR>
	   - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea3" style="display:<%if oitem.FOneItem.Fsourcekind ="3" then%>block<%else%>none<%end if%>;">
	  <p>�Ұ���� ��� ������ ����(�ѿ�/����/���ұ���) �� ������   <span style="margin-right:10px;">ex. ����(Ⱦ���� �ѿ�), ����(ȣ�ֻ�)</span></BR>
	  - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div>
	  <div id="dvArea4" style="display:<%if oitem.FOneItem.Fsourcekind ="4" then%>block<%else%>none<%end if%>;">
	  <p><strong>98%�̻� ���ᰡ �ִ� ���:</strong>  �Ѱ��� ���Ḹ ǥ�� ����    <span style="margin-right:10px;">ex. ����(�̱���)</span> </BR>
	  	<strong>���� ���Ḧ ����� ���:</strong> ȥ�պ����� ���� ������ 2�� ����   <span style="margin-right:10px;">ex. ������[�а���(�̱���),���尡��(������)]</span></BR>
	  - ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.</p>
	  </div> 
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="makername" maxlength="32" size="25" id="[on,off,off,off][������]" value="<%= oitem.FOneItem.Fmakername %>">&nbsp;(������ü��)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="Ű���� �˻����� ���� �߰� �ܾ��" style="cursor:help;">�˻�Ű���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="keywords" maxlength="250" size="50" class="text" id="[on,off,off,off][�˻�Ű����]" value="<%= oitem.FOneItem.Fkeywords %>">&nbsp;(�޸��α��� ex: Ŀ��,Ƽ����,����)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="��ǰ �� �Ӽ�" style="cursor:help;">��ǰ�Ӽ� :</td>
	<td id="lyrItemAttribAdd" bgcolor="#FFFFFF" colspan="3"></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ü��ǰ�ڵ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="upchemanagecode" class="text" id="[off,off,off,off][��ü��ǰ�ڵ�]" value="<%= oitem.FOneItem.Fupchemanagecode %>" size="20" maxlength="32">
		(��ü���� �����ϴ� �ڵ� �ִ� 32�� - ����/���ڸ� ����)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">ISBN :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		ISBN 13 <input type="text" name="isbn13" class="text" value="<%= oitem.FOneItem.Fisbn13 %>" size="13" maxlength="13">
		/ �ΰ���ȣ <input type="text" name="isbn_sub" class="text" value="<%= oitem.FOneItem.FisbnSub %>" size="5" maxlength="5"><br />
		ISBN 10 <input type="text" name="isbn10" class="text" value="<%= oitem.FOneItem.Fisbn10 %>" size="10" maxlength="10"> (Optional)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ǰ��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <input type="text" name="relateItems" value="<%=strItemRelation%>" size="52" class="text" id="[off,off,off,off][������ǰ]">
	    <br>(������ǰ�� �ִ� 6������ ��ϰ���, ��ǰ��ȣ�� �޸�(,)�� �����Ͽ� �Է�)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ ���� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<p style="text-align:right;"><input type="button" value="��ǰ �̹��� ����" class="bgBlue" onClick="window.open('http://www.10x10.co.kr/shopping/itemImageView.asp?itemid=<%=itemid%>');"></p>
		<div>
		<!--
		<label><input type="radio" name="usinghtml" value="N" <% if oitem.FOneItem.Fusinghtml = "N" then response.write "checked" %>>�Ϲ�TEXT</label>
		<label><input type="radio" name="usinghtml" value="H" <% if oitem.FOneItem.Fusinghtml = "H" then response.write "checked" %>>TEXT+HTML</label>
		<label><input type="radio" name="usinghtml" value="Y" <% if oitem.FOneItem.Fusinghtml = "Y" then response.write "checked" %>>HTML���</label>
		<br>
		-->
		<input type="hidden" name="usinghtml" value="Y" />
		<textarea name="itemcontent" rows="15" class="textarea" style="width:100%" id="[on,off,off,off][�����ۼ���]"><%= oitem.FOneItem.Fitemcontent %></textarea>
		<script>
		//
		window.onload = new function(){
			var itemContEditor = CKEDITOR.replace('itemcontent',{
				height : 400,
				// ���ε�� ���� ���
				//filebrowserBrowseUrl : '/browser/browse.asp',
				// ���� ���ε� ó�� ������
				filebrowserImageUploadUrl : '<%= ItemUploadUrl %>/linkweb/items/itemEditorContentUpload.asp?itemid=<%=itemid%>'
			});
			itemContEditor.on( 'change', function( evt ) {
			    // �Է��� �� textarea ���� ����
			    document.itemreg.itemcontent.value = evt.editor.getData();
			});
		}
		</script>
		</div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ ������ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="itemvideo" rows="5" class="textarea" cols="80" id="[off,off,off,off][�����۵�����]"><%=oitemvideo.FOneItem.FvideoFullUrl%></textarea>
	    <br>�� Youtube, Vimeo ������ ����(Youtube : �ҽ��ڵ尪 �Է�, Vimeo : �Ӻ����� �Է�)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ��� ���ǻ��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="ordercomment" rows="5" cols="80" class="textarea" id="[off,off,off,off][���ǻ���]"><%= oitem.FOneItem.Fordercomment %></textarea><br>
	<font color="red">Ư���� ��۱Ⱓ�̳� �ֹ��� Ȯ���ؾ߸� �ϴ� ����</font>�� �Է��Ͻø� ���Ҹ��̳� ȯ���� ���ϼ� �ֽ��ϴ�.
	</td>
</tr>
</table>

<!-- ǰ������� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">ǰ������� &nbsp;<font color=gray>��ǰ����������� ���� ���� ������ ���� �Ʒ� ������ ��Ȯ�� �Է����ֽñ� �ٶ��ϴ�.</font></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">ǰ���� :</td>
	<td bgcolor="#FFFFFF">
		<% DrawInfoDiv "infoDiv", oitem.FOneItem.FinfoDiv, " onchange='chgInfoDiv(this.value);'" %>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:<%=chkIIF(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="","none","")%>;">
	<td height="30" width="15%" bgcolor="#F8DDFF">ǰ�񳻿� :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList">
	<%
		if Not(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="") then
			Server.Execute("act_itemInfoDivForm.asp")
		end if
	%>
	</td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">��ǰ���������� ������ ���� �Ǿ��ִ��� ��Ȯ�� �Է¹ٶ��ϴ�. ����Ȯ�ϰų� �߸��� ���� �Է½�, �׿� ���� å���� ���� ���� �ֽ��ϴ�.</font></td>
</tr>
<tr align="left" id="lyItemSrc" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsource" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsource %>">&nbsp;(ex:�ö�ƽ,����,��,...)
	</td>
</tr>
<tr align="left" id="lyItemSize" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ������ :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsize" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsize %>">&nbsp;(ex:7.5x15(cm))
	</td>
</tr>
</table>
<!-- ������������ -->
<%
dim arrAuth, r, real_safetydiv, real_safetynum, safetyDivList
arrAuth = oitem.FAuthInfo
if isArray(arrAuth) THEN
	For r =0 To UBound(arrAuth,2)
		real_safetydiv = real_safetydiv & arrAuth(0,r)
		if r <> UBound(arrAuth,2) then real_safetydiv = real_safetydiv & "," end if
		
		real_safetynum = real_safetynum & arrAuth(1,r)
		if r <> UBound(arrAuth,2) then real_safetynum = real_safetynum & "," end if
		
		safetyDivList = safetyDivList & "<p class='tPad05' id='l"&arrAuth(0,r)&"'>"
		safetyDivList = safetyDivList & "- "&fnSafetyDivCodeName(arrAuth(0,r))&"("&CHKIIF(arrAuth(1,r)="x","������ȣ ����",arrAuth(1,r))&")"
		safetyDivList = safetyDivList & " <input type='button' value='����' class='btn3 btnIntb' onClick='jsSafetyDivListDel("&arrAuth(0,r)&");'>"
		safetyDivList = safetyDivList & "</p>"
	Next
end if
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">������������</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">
		����������� :
		<input type="button" value="�������� �ʼ� ǰ�� Ȯ��" onclick="jsSafetyPopup();" class="button" />
	</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
		<tr align="left" height="30">
			<td bgcolor="#FFFFFF">
				<label><input type="radio" name="safetyYn" value="Y" <%=chkIIF(oitem.FOneItem.FsafetyYn="Y","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> ���</label>
				<label><input type="radio" name="safetyYn" value="N" <%=chkIIF(oitem.FOneItem.FsafetyYn="N","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> ���ƴ�</label>
				<label><input type="radio" name="safetyYn" value="I" <%=chkIIF(oitem.FOneItem.FsafetyYn="I","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> ��ǰ���� ǥ��</label>
				<label><input type="radio" name="safetyYn" value="S" <%=chkIIF(oitem.FOneItem.FsafetyYn="S","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> ���������ؼ�</label>
				<input type="hidden" name="auth_go_catecode" id="auth_go_catecode" value="">
				<input type="hidden" name="real_safetydiv" id="real_safetydiv" value="<%=real_safetydiv%>">
				<input type="hidden" name="real_safetynum" id="real_safetynum" value="<%=real_safetynum%>">
				<input type="hidden" name="real_safetyidx" id="real_safetyidx" value="">
				<input type="hidden" name="real_safetynum_delete" id="real_safetynum_delete" value="">
				<input type="hidden" name="real_safetydiv_delete" id="real_safetydiv_delete" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<% drawSelectBoxSafetyDivCode "safetyDiv", "", oitem.FOneItem.FsafetyYn, "" %>

				������ȣ <input type="text" name="safetyNum" id="[off,off,off,off][�������� ������ȣ]" <%=chkIIF(oitem.FOneItem.FsafetyYn<>"Y","disabled","")%> size="35" maxlength="25" value="" /><%'=oitem.FOneItem.FsafetyNum%>
				<input type="button" id="safetybtn" value="��   ��" onclick="jsSafetyAuth();" class="button">
				<input type="hidden" name="issafetyauth" id="issafetyauth" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<div id="safetyDivList">
					<%=safetyDivList%>
				</div>
				<div id="safetyYnI" style="display:none;">
					<font color="blue">��ǰ ���� ǥ��(ǥ���� ��ǰ�ΰ�� ��ǰ �� �������� ������ȣ�� �𵨸�, KC ��ũ�� �� ǥ�����ּ���.)</font>
				</div>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td bgcolor="#FFFFFF" colspan=2>
		* ���������� �Է� �� �ϰų�, �߸��� ���������� �Է��� ��� �߰� <strong><font color='red'>��� �Ǹ����� �Ǵ� ����</font></strong> �˴ϴ�.<br>
		* <strong><font color='red'>���������ؼ�</font></strong> ����ϰ�� ������ȣ�� ������, KC��ũ�� ǥ������ �ʾƾ� �˴ϴ�.<br>
		* �Է��� ���������� ��ǰ�����������Ϳ��� ������ ������ �������� ��ȸ�Ǹ�, <strong><font color='red'>�������� ���� ������ ����� �Ұ�</font></strong>���մϴ�.<br>
		* �������� ���������� �Է��������� �ұ��ϰ� ����� �ȵɰ�쿡 "��ǰ���� ǥ��"�� ������ �����ϸ�, ��ǰ �� �������� �𵨸�� ǥ���� ��ǰ�ΰ�� ������ȣ,KC��ũ�� ǥ���ؾ� �մϴ�.<br>
		* ������������ ���� ���Ǵ� Ȩ������(<u><a href="http://safetykorea.kr" target="_blank">http://safetykorea.kr</a></u>)�� Ȯ���� �ֽñ� �ٶ��ϴ�.
	</td>
</tr>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="�����ϱ�" class="button" onClick="SubmitSave()">
          <input type="button" value="����ϱ�" class="button" onClick="self.close()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
</form>

<script type="text/javascript">
	itemreg.makerid.readOnly = true;
	itemreg.frontMakerid.readOnly = true;

	// ��������üũ. ���ȹ�
	jsSafetyCheck('<%= oitem.FOneItem.FsafetyYn %>','');
</script>

<% 
set oitem = Nothing
Set oitemvideo = Nothing

Sub SelectBoxDesignerItem(selectedId)
   dim query1,tmp_str
   %><select name="designer" onchange="TnDesignerNMargineAppl(this.value);">
     <option value='' <%if selectedId="" then response.write " selected"%>>-- ��ü���� --</option><%
   query1 = " select userid,socname_kor,defaultmargine from [db_user].[dbo].tbl_user_c order by userid"
'   query1 = query1 + " where isusing='Y' order by userid desc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")& "," & rsget("defaultmargine") & "' "&tmp_str&">" & rsget("userid") & "  [" & replace(db2html(rsget("socname_kor")),"'","") & "]" & "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->