<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<%

dim itemid, oitem , oitemvideo
dim makerid
Dim fingerson : fingerson = "on" '//��ǰ��ÿ� fingersflag

itemid = RequestCheckvar(request("itemid"),10)
makerid = RequestCheckvar(request("makerid"),32)
menupos = RequestCheckvar(request("menupos"),10)
if (itemid = "") then
    response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if


'==============================================================================
set oitem = new CItem

oitem.FRectItemID = itemid
oitem.GetOneItem

Set oitemvideo = New CItem
oitemvideo.FRectItemId = itemid
oitemvideo.FRectItemVideoGubun = "video1"
oitemvideo.GetItemContentsVideo

'==============================================================================
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

'==============================================================================
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
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script>
function UseTemplate() {
	window.open("/academy/comm/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
}

// ============================================================================
// ī�װ����
function editCategory(cdl,cdm,cds){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cds=" + cds ;

	popwin = window.open('/academy/comm/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
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
        alert('�ֹ����� ��ǰ�� ��ü����ΰ�츸 �����մϴ�.');
        itemreg.itemdiv[0].focus();
        return;
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

	//������������
    if (itemreg.safetyYn[0].checked){
	    if (!itemreg.safetyDiv.value){
	        alert('�������������� �������ּ���.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
	    if (!itemreg.safetyNum.value){
	        alert('����������ȣ�� �Է����ּ���.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
    }

    if(confirm("��ǰ�� �ø��ðڽ��ϱ�?") == true){
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
		var popwin = window.open("/academy/comm/NewCategorySelect.asp?iid=" + iid, "popCateSel","width=700,height=400,scrollbars=yes,resizable=yes");
        popwin.focus();
	}

	// ����ī�װ� ���� �˾�
	function popDispCateSelect(){
		var designerid = document.all.itemreg.designerid.value;
		if(designerid == ""){
			alert("��ü�� �����ϼ���.");
			return;
		}
		
		var dCnt = $("input[name='isDefault'][value='y']").length;
		$.ajax({
			url: "/academy/comm/act_DispCategorySelect.asp?designerid="+designerid+"&isDft="+dCnt,
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
		//printItemAttribute();
	}

	// ���� ī�װ� ����
	function delCateItem()
	{
		if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?"))
			tbl_Category.deleteRow(tbl_Category.clickedRowIndex);
	}
	
	// ���� ����ī�װ� ����
	function delDispCateItem() {
		if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

			//��ǰ�Ӽ� ���
			//printItemAttribute();
		}
	}
//-------------------------------------
function chgodr(v){
	if (v == 1){
		$("#customorder").css("display","none");
	}else{
		$("#customorder").css("display","");
	}
}

function chgodr2(v){
	if (v == 1){
		$("#subodr").css("display","none");
	}else{
		$("#subodr").css("display","");
	}
}

// ������������ ����
function chgSafetyYn(frm) {
	if(frm.safetyYn[0].checked) {
		frm.safetyDiv.disabled=false;
		frm.safetyNum.disabled=false;
	} else {
		frm.safetyDiv.disabled=true;
		frm.safetyNum.disabled=true;
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
			url: "/admin/itemmaster/act_itemInfoDivForm.asp",
			data: "itemid=<%=itemid%>&ifdv="+v+"&fingerson=on",
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

function checkItemDiv(comp){
    var frm = comp.form;
    
    if (comp.name=="itemdiv"){
        if (frm.itemdiv[1].checked){
            frm.reqMsg.disabled=false;
            frm.requireimgchk.disabled=false;
        }else{
            //frm.reqMsg.checked=false;
            frm.reqMsg.disabled=true;
            frm.requireimgchk.disabled=true;
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
}

function requireimg(){
	var frm = document.itemreg;
	if (frm.requireimgchk.checked){
		$("#rmemail").css("display","");
	}else{
		$("#rmemail").css("display","none");
	}
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
        <td align="left">
          <br>�⺻����
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<form name="itemreg" method="post" action="itemmodify_Process.asp" onsubmit="return false;" style="margin:0;">
<input type="hidden" name="mode" value="ItemBasicInfo">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="designerid" value="<%= oitem.FOneItem.Fmakerid %>">
<input type="hidden" name="orgprice" value="<%= oitem.FOneItem.Forgprice %>">
<input type="hidden" name="orgsuplycash" value="<%= oitem.FOneItem.Forgsuplycash %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�ڵ� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <%= itemid %>
  	  &nbsp;&nbsp;&nbsp;&nbsp;
  	  <input type="button" value="�̸�����" onclick="window.open('<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%= itemid %>');">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�귣��ID :</td>
  	<td bgcolor="#FFFFFF" colspan="3"><%= oitem.FOneItem.Fmakerid %></td>
  </tr>
  <!-- // ��Ƽ ī�װ� �߰������� ���� (��ü ��ǰ���������� �⺻ī�װ� ���������� ���) (2008.03.28;������ ����)
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">ī�װ� ���� :</td>
    <input type="hidden" name="cd1" value="<%= oitem.FOneItem.FCate_large %>">
    <input type="hidden" name="cd2" value="<%= oitem.FOneItem.FCate_mid %>">
    <input type="hidden" name="cd3" value="<%= oitem.FOneItem.FCate_small %>">
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="cd1_name" value="<%= oitem.FOneItem.FCate_large_name %>" id="[on,off,off,off][ī�װ�]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd2_name" value="<%= oitem.FOneItem.FCate_mid_name %>" id="[on,off,off,off][ī�װ�]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd3_name" value="<%= oitem.FOneItem.FCate_small_name %>" id="[on,off,off,off][ī�װ�]" size="20" readonly style="background-color:#E6E6E6">

      <input type="button" value="ī�װ� ����" onclick="editCategory(itemreg.cd1.value,itemreg.cd2.value,itemreg.cd3.value);">
  	</td>
  </tr>
  //-->
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">���� ī�װ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td><%=getCategoryInfo(itemid)%></td>
			<td valign="bottom"><input type="button" value="�߰�" onClick="popCateSelect('<%=itemid%>')"></td>
		</tr>
		</table>
  	</td>
  </tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="����Ʈ�� ������ ī�װ�" style="cursor:help;">���� ī�װ� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td><%=getDispCategory(itemid)%></td>
			<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="itemdiv" value="01" <% if oitem.FOneItem.Fitemdiv ="01" then  response.write "checked" %> onclick="checkItemDiv(this);chgodr2(1);">�Ϲݻ�ǰ
      <input type="radio" name="itemdiv" value="<%= oitem.FOneItem.Fitemdiv %>" <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","checked","")%> onclick="checkItemDiv(this);chgodr(2);">�ֹ����ۻ�ǰ
      <input type="checkbox" name="reqMsg" value="10" <%=chkIIF(oitem.FOneItem.Fitemdiv="06","checked","")%> <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","","disabled")%> onClick="checkItemDiv(this);">�ֹ����� ���� �ʿ�<font color=red>(�ֹ����� �޼����� �ʿ��� ���)</font>
	  <input type="checkbox" name="requireimgchk" value="Y" <%=chkIIF(oitem.FOneItem.Frequirechk="Y","checked","")%> onClick="requireimg();">�ֹ����� �̹��� �ʿ�
<!-- 	  <br> -->
<!--       <input type="radio" name="itemdiv" value="20" <% if oitem.FOneItem.Fitemdiv ="20" then  response.write "checked" %> onclick="checkItemDiv(this);chgodr(1);">�߰������ǰ -->
<!--       <font color="red">(��ǰ��Ͽ����� ����, �߰��ɼǿ����� ������)</font> -->
      <% if oitem.FOneItem.Fitemdiv ="07" then %>
      <input type="radio" name="itemdiv" value="07" <% if oitem.FOneItem.Fitemdiv ="07" then  response.write "checked" %> onclick="checkItemDiv(this);chgodr(1);">����(����)���Ż�ǰ
      <% end if %>
      
      <% if oitem.FOneItem.Fitemdiv ="82" then %>
      <input type="radio" name="itemdiv" value="82" <% if oitem.FOneItem.Fitemdiv ="82" then  response.write "checked" %> onclick="checkItemDiv(this);chgodr(1);">���ϸ����� ��ǰ
      <% end if %>
  	</td>
  </tr>
    <!-- �ֹ� ���� �̸��� -->
  <tr id="rmemail" style="display:<%=chkiif(oitem.FOneItem.Frequirechk="Y","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ����� �̸��� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="requireMakeEmail" value="<%=oitem.FOneItem.FrequireEmail%>" size="50" maxlength="100"> (ex)�۰����� ���� �ּ�)
  	</td>
  </tr>
  <!-- �ֹ� ���� �̸��� -->
  <tr id="customorder" style="display:<%=chkiif(oitem.FOneItem.Fitemdiv="06" Or oitem.FOneItem.Fitemdiv="16","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ����� �߰��ɼ�</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="cstodr" value="1" onclick="chgodr2(1)" <%=chkiif(oitem.FOneItem.Fcstodr="1","checked","")%>>��ù߼�
      <input type="radio" name="cstodr" value="2" onclick="chgodr2(2)" <%=chkiif(oitem.FOneItem.Fcstodr="2","checked","")%>>������ �߼�<br>
	  <div id="subodr" style="display:<%=chkiif(oitem.FOneItem.Fcstodr="2","block","none")%>;">
		������ �߼� �Ⱓ <input type="text" name="requireMakeDay" value="<%=oitem.FOneItem.FrequireMakeDay%>" size="3" maxlength="2">��<br>
		&lt--Ư�̻����� �Է� ���ּ���--&gt;<br><textarea name="requirecontents" rows="5" cols="80"><%=oitem.FOneItem.Frequirecontents%></textarea>
	  </div>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemname" maxlength="64" size="60" id="[on,off,off,off][��ǰ��]" value="<%= oitem.FOneItem.Fitemname %>">&nbsp;
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemsource" maxlength="64" size="50" id="[on,off,off,off][��ǰ����]" value="<%= oitem.FOneItem.Fitemsource %>">&nbsp;(ex:�ö�ƽ,����,��,...)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ������ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemsize" maxlength="64" size="50" id="[on,off,off,off][��ǰ������]" value="<%= oitem.FOneItem.Fitemsize %>">&nbsp;(ex:7.5x15(cm))
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemWeight" maxlength="12" size="4" id="[on,off,off,off][��ǰ����]" value="<%= oitem.FOneItem.FitemWeight %>">g&nbsp;(���Դ� g������ �Է�)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="sourcearea" maxlength="64" size="25" id="[on,off,off,off][������]" value="<%= oitem.FOneItem.Fsourcearea %>">&nbsp;(ex:�ѱ�,�߱�,�߱�OEM,�Ϻ�...)
      <br>( ������ ǥ�� ������ ��Ŭ������ ���� ū ���� �� �ϳ��Դϴ�. ��Ȯ�� �Է��� �ּ���.)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">������ :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="makername" maxlength="32" size="25" id="[on,off,off,off][������]" value="<%= oitem.FOneItem.Fmakername %>">&nbsp;(������ü��)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�˻�Ű���� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="keywords" maxlength="50" size="50" id="[on,off,off,off][�˻�Ű����]" value="<%= oitem.FOneItem.Fkeywords %>">&nbsp;(�޸��α��� ex: Ŀ��,Ƽ����,����)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ü��ǰ�ڵ� :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	    <input type="text" name="upchemanagecode" id="[off,off,off,off][��ü��ǰ�ڵ�]" value="<%= oitem.FOneItem.Fupchemanagecode %>" size="20" maxlength="32">
  	    (��ü���� �����ϴ� �ڵ� �ִ� 32�� - ����/���ڸ� ����)
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ ���� :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="radio" name="usinghtml" value="N" <% if oitem.FOneItem.Fusinghtml = "N" then response.write "checked" %>>�Ϲ�TEXT -->
<!--       <input type="radio" name="usinghtml" value="H" <% if oitem.FOneItem.Fusinghtml = "H" then response.write "checked" %>>TEXT+HTML -->
<!--       <input type="radio" name="usinghtml" value="Y" <% if oitem.FOneItem.Fusinghtml = "Y" then response.write "checked" %>>HTML��� -->
<!--       <br> -->
<!--       <textarea name="itemcontent" rows="10" cols="80" id="[on,off,off,off][�����ۼ���]"><%= oitem.FOneItem.Fitemcontent %></textarea> -->
<!--   	</td> -->
<!--   </tr> -->
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">�ֹ��� ���ǻ��� :<br/>(��۾ȳ�)</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="ordercomment" rows="5" cols="80" id="[off,off,off,off][���ǻ���]"><%= oitem.FOneItem.Fordercomment %></textarea><br>
      <font color="red">Ư���� ��۱Ⱓ�̳� �ֹ��� Ȯ���ؾ߸� �ϴ� ����</font>�� �Է��Ͻø� ���Ҹ��̳� ȯ���� ���ϼ� �ֽ��ϴ�.
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">��ȯ / ȯ�� ��å</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="refundpolicy" rows="5" cols="80" id="[off,off,off,off][ȯ����å]"><%=oitem.FOneItem.Frefundpolicy%></textarea><br>
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">��ü�ڸ�Ʈ :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="text" name="designercomment" size="50" maxlength="40" id="[off,off,off,off][��ü�ڸ�Ʈ]" value="<%= oitem.FOneItem.Fdesignercomment %>"><br> -->
<!--       ��ǰ������ ���丮�� ��̳� �̾߱⸦ �����ּ���... -->
<!--   	</td> -->
<!--   </tr> -->
  <tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">������ ������ :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="itemvideo" rows="5" class="textarea" cols="90" id="[off,off,off,off][�����۵�����]"><%= db2html(oitemvideo.FOneItem.FvideoFullUrl) %></textarea>
		<br>�� Youtube, Vimeo ������ ����(Youtube : �ҽ��ڵ尪 �Է�, Vimeo : �Ӻ����� �Է�)
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
		<select name="infoDiv" class="select" onchange="chgInfoDiv(this.value)">
		<option value="">::��ǰǰ��::</option>
		<option value="01">�Ƿ�</option>
		<option value="02">����/�Ź�</option>
		<option value="03">����</option>
		<option value="04">�м���ȭ(����/��Ʈ/�׼�����)</option>
		<option value="05">ħ����/Ŀư</option>
		<option value="06">����(ħ��/����/��ũ��/DIY��ǰ)</option>
<!-- 		<option value="07">������(TV��)</option> -->
<!-- 		<option value="08">������ ������ǰ(�����/��Ź��/�ı⼼ô��/���ڷ�����)</option> -->
<!-- 		<option value="09">��������(������/��ǳ��)</option> -->
<!-- 		<option value="10">�繫����(��ǻ��/��Ʈ��/������)</option> -->
<!-- 		<option value="11">���б��(������ī�޶�/ķ�ڴ�)</option> -->
<!-- 		<option value="12">��������(MP3/���ڻ��� ��)</option> -->
<!-- 		<option value="14">������̼�</option> -->
		<option value="15">�ڵ�����ǰ(�ڵ�����ǰ/��Ÿ �ڵ�����ǰ)</option>
<!-- 		<option value="16">�Ƿ���</option> -->
		<option value="17">�ֹ��ǰ</option>
		<option value="18">ȭ��ǰ</option>
		<option value="19">�ͱݼ�/����/�ð��</option>
		<option value="20">��ǰ(����깰)</option>
		<option value="21">������ǰ</option>
		<option value="22">�ǰ���ɽ�ǰ/ü��������ǰ</option>
		<option value="23">�����ƿ�ǰ</option>
		<option value="24">�Ǳ�</option>
		<option value="25">��������ǰ</option>
		<option value="26">����</option>
<!-- 		<option value="27">ȣ��/��ǿ���</option> -->
<!-- 		<option value="28">�����ǰ</option> -->
<!-- 		<option value="29">�װ���</option> -->
		<option value="35">��Ÿ</option>
		</select>
		<script type="text/javascript">
		document.itemreg.infoDiv.value="<%=oitem.FOneItem.FinfoDiv%>";
		chgInfoDiv(<%=oitem.FOneItem.FinfoDiv%>);
		</script>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:<%=chkIIF(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="","none","")%>;">
	<td height="30" width="15%" bgcolor="#F8DDFF">ǰ�񳻿� :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList">
	<%
		if Not(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="") Then
			Server.Execute("/admin/itemmaster/act_itemInfoDivForm.asp")
		end if
	%>
	</td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">��ǰ���������� ������ ���� �Ǿ��ִ��� ��Ȯ�� �Է¹ٶ��ϴ�. ����Ȯ�ϰų� �߸��� ���� �Է½�, �׿� ���� å���� ���� ���� �ֽ��ϴ�.</font></td>
</tr>
<!-- <tr align="left" id="lyItemSrc" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;"> -->
<!-- 	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ���� :</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" name="itemsource" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsource %>">&nbsp;(ex:�ö�ƽ,����,��,...) -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr align="left" id="lyItemSize" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;"> -->
<!-- 	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ������ :</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" name="itemsize" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsize %>">&nbsp;(ex:7.5x15(cm)) -->
<!-- 	</td> -->
<!-- </tr> -->
</table>
<!-- ������������ -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">������������</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">����������� :</td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="safetyYn" value="Y" <%=chkIIF(oitem.FOneItem.FsafetyYn="Y","checked","")%> onclick="chgSafetyYn(document.itemreg)"> ���</label>
		<label><input type="radio" name="safetyYn" value="N" <%=chkIIF(oitem.FOneItem.FsafetyYn<>"Y","checked","")%> onclick="chgSafetyYn(document.itemreg)"> ���ƴ�</label> /
		<select name="safetyDiv" <%=chkIIF(oitem.FOneItem.FsafetyYn<>"Y","disabled","")%> class="select">
		<option value="">::������������::</option>
		<option value="10" <%=chkIIF(oitem.FOneItem.FsafetyDiv="10","selected","")%>>������������(KC��ũ)</option>
		<option value="20" <%=chkIIF(oitem.FOneItem.FsafetyDiv="20","selected","")%>>�����ǰ ��������</option>
		<option value="30" <%=chkIIF(oitem.FOneItem.FsafetyDiv="30","selected","")%>>KPS �������� ǥ��</option>
		<option value="40" <%=chkIIF(oitem.FOneItem.FsafetyDiv="40","selected","")%>>KPS �������� Ȯ�� ǥ��</option>
		<option value="50" <%=chkIIF(oitem.FOneItem.FsafetyDiv="50","selected","")%>>KPS ��� ��ȣ���� ǥ��</option>
		</select>
		������ȣ <input type="text" name="safetyNum" <%=chkIIF(oitem.FOneItem.FsafetyYn<>"Y","disabled","")%> size="35" maxlength="25" class="text" value="<%=oitem.FOneItem.FsafetyNum%>" />
		
		<font color="darkred">���ƿ�ǰ�̳� �����ǰ�� ��� �ʼ� �Է�</font>
	</td>
</tr>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="�����ϱ�" onClick="SubmitSave()">
          <input type="button" value="����ϱ�" onClick="self.close()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
</form>
<!-- ǥ �ϴܹ� ��-->
<% 
set oitem = Nothing
Set oitemvideo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->