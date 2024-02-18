<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/deal/deal_reg.asp
' Description :  �� �̺�Ʈ ���
' History : 2017.08.23 ������
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/newdealManageCls.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 560   'KB
Dim idx, itemsort
idx = requestCheckVar(Request("idx"),10)
itemsort  	= requestCheckvar(request("itemsort"),32)
If idx="" Then
	Dim oDealMax
	set oDealMax = New ClsDeal
	oDealMax.fnGetMAXDealMasterNum
	idx=oDealMax.FMasterIDX
	Set oDealMax=Nothing
Response.redirect "/admin/itemmaster/deal/new_deal_reg.asp?idx="&idx
Response.End
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
<!--
	function TnViewDivSelect(viewdiv){
		if(viewdiv==1){
			$("#datearea").css("display","none");
		}else{
			$("#datearea").css("display","");
		}
	}

	function TnSearchObjOpenWin(){
		//var winpop = window.open('/admin/itemmaster/deal/pop_deal_additemlist.asp?idx=<%=idx%>&stype=w','winpop','width=1024,height=768,scrollbars=yes,resizable=yes');
		var winpop = window.open('/admin/itemmaster/deal/dealitem_regist.asp?idx=<%=idx%>&stype=w','winpop','width=1024,height=768,scrollbars=yes,resizable=yes');
	}

	function SubmitSave(frm){
		if(frm.itemname.value=="")
		{
			alert("��ǰ���� �Է����ּ���.");
			frm.itemname.focus();
			return false;
		}
		else if(frm.itemname.value.length>50)
		{
			alert("50�� �̳��� ��ǰ���� �Է����ּ���.");
			frm.itemname.focus();
			return false;
		}
		else if(!frm.viewdiv[0].checked && !frm.viewdiv[1].checked)
		{
			alert("���� �Ⱓ�� �������ּ���.");
			return false;
		}
		else if(frm.viewdiv[1].checked && (frm.startdate.value=="" || frm.enddate.value==""))
		{
			alert("���� �Ⱓ�� �������ּ���.");
			return false;
		}
		else if(!frm.isusing[0].checked && !frm.isusing[1].checked)
		{
			alert("��� ���θ� �������ּ���.");
			return false;
		}
		else if(frm.itemid.value=="")
		{
			alert("��ǥ��ǰ�� �������ּ���.");
			frm.itemid.focus();
			return false;
		}
		else if(frm.mastersellcash.value=="")
		{
			alert("��ǥ ������ �Է����ּ���.");
			frm.mastersellcash.focus();
			return false;
		}
		else if(frm.masterdiscountrate.value=="")
		{
			alert("��ǥ �������� �Է����ּ���.");
			frm.masterdiscountrate.focus();
			return false;
		}
		else if($("#tbl_DispCate tr").length<1)
		{
			alert("���� ī�װ��� �߰����ּ���.");
			frm.catecode.focus();
			return false;
		}
		else if(frm.keywords.value=="")
		{
			alert("�˻� Ű���带 �Է��� �ּ���.");
			frm.keywords.focus();
			return false;
		}
		else
		{
			if(confirm("�Է��Ͻ� ������ ����ǰ�� ����Ͻðڽ��ϱ�?"))
			{
				frm.action="dodealinfo_process.asp";
				frm.submit();
			}
		}
	}

    function TnDealSaveAPICall(itemid){
		document.frm.target="FrameCKP";
		document.frm.tempitemid.value=itemid;
		document.frm.action="<%= ItemUploadUrl %>/linkweb/items/deal_itemregisterTempWithImage_process.asp";
		frm.submit();
    }


	function TnMasterItemSelect(itemid){
		if(document.frm.itemname.value=="")
		{
			document.frm.itemname.value=$("#itemcode option:selected").text();
		}
		$("#selectitem").val(itemid);
		$.ajax({
			url: "selectdealitemkeywords.asp?itemid="+itemid,
			cache: false,
			async: false,
			success: function(message) {
				//alert(message);
				if(message!="") {
					$('#keywords').val(message);
				}
			}
		});
	}

	// �⺻���� ����
	function editItemBasicInfo(itemid) {
		var param = "itemid=" + itemid + "&menupos=<%= menupos %>";
		popwin = window.open('/admin/itemmaster/pop_ItemBasicInfo.asp?' + param ,'editItemBasic','width=1100,height=600,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	// �⺻���� ����
	function fnSaleInfo() {
		popwin = window.open('/admin/shopmaster/sale/saleList.asp?menupos=290' ,'saleinfo','width=1100,height=600,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	function onlyNumerSet(text){
		if(window.event.keyCode < 48 || window.event.keyCode > 57) {
			return false;
		}
	}

	function fnPaste() {
		var regex = /\D/ig;
		if (regex.test(window.clipboardData.getData("text"))) {
			return false;
		} else {
			return true;
		}
	}

	function jsSetImg(sName, sSpan){ 
		var winImg;
		winImg = window.open('pop_deal_uploadimg.asp?yr=<%=Year(now())%>&sName='+sName+'&sSpan='+sSpan+'&wid=900&hei=1600','popImg','width=370,height=150');
		winImg.focus();
	}

	function fnItemSelectboxLoad(){
		$.ajax({
			type: "POST",
			url: "ajaxDealItemSelectboxLoad.asp",
			data: "idx=<%=idx%>",
			dataType: "JSON",
			cache: false,
			success: function(data){
				$("#itemcode").attr("disabled",false);
				$('#itemcode').children('option:not(:first)').remove();
				$.each(data.option, function(i, record) {
					$("#itemcode").append($("<option></option>").attr("value",record.optionValue).text(record.optionName));
				});
				$("#itemcode").val($("#selectitem").val()).prop("selected", true);
				$("#mastersellcash").val(data.minPrice);
				$("#masterdiscountrate").val(data.salePer);
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
	}

	function jsAddGroup(){
		var wingroup;
		wingroup = window.open('pop_dealitem_group.asp?idx=<%=idx%>','popGroup','width=500,height=350');
		wingroup.focus();
	}

	function fnLoadItems(){
		$.ajax({
			type: "POST",
			url: "doDealItemInfo.asp",
			data: "mode=load&idx=<%=idx%>",
			cache: false,
			success: function(data) {
				if(data.response=="ok"){
					$("#itemButton").val("�� ��ǰ ���� (" + data.itemCount + "��)");
					$("#groupButton").val("�� �׷� ���� (" + data.groupCount + "��)");
				}else{
					alert("������ ó���� ������ �߻��Ͽ����ϴ�.");
				}
			},
			error: function(err) {
				console.log(err.responseText);
			}
		});
	}

	$(function(){
		fnLoadItems();
		fnItemSelectboxLoad();
	});
//-->
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<form name="frm" method="post" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="auser" value="<%=session("ssBctId")%>">
<input type="hidden" name="catecnt" id="catecnt">
<input type="hidden" name="mode" value="reg">
<input type="hidden" name="tempitemid">
<input type="hidden" name="sortarr">
<input type="hidden" name="sitemarr">
<input type="hidden" id="selectitem">
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>�� �⺻ ����</B></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="itemname" id="itemname" size="80" maxlength="120" value="" class="text">
		</td>
	</tr>
	<tr>
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">���� �Ⱓ<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			 <input type="radio" name="viewdiv" id="viewdiv" value="1" onClick="TnViewDivSelect(1)" checked>��õ� <input type="radio" name="viewdiv" id="viewdiv" value="2" onClick="TnViewDivSelect(2)">�Ⱓ��
			 <span id="datearea" style="display:none">
				<input id="startdate" name="startdate" value="<%=Date()%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> <input type="text" name="shour" size="2" class="text" value="00" maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');">:<input type="text" name="sminute" size="2" class="text" value="00" maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');"> ~
				<input id="enddate" name="enddate" value="<%=DateAdd("D",14,Date())%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" /> <input type="text" name="ehour" size="2" class="text" value="23" maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');">:<input type="text" name="eminute" size="2" class="text" value="59" maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');">
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "startdate", trigger    : "iSD_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "enddate", trigger    : "iED_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
			 </span>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��뿩��<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="radio" name="isusing" id="isusing" value="Y" checked>��� <input type="radio" name="isusing" id="isusing" value="N">��� ����
		</td>
	</tr>
</table>
<p>&nbsp;</p>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>���� ��ǰ ����</B></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width="100">��ǰ���<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="button" class="button" id="itemButton" value="�� ��ǰ ���� (��)" onclick="TnSearchObjOpenWin();">&nbsp;
			<input type="button" class="button" id="groupButton" value="�� �׷� ���� (��)" onclick="jsAddGroup();">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ ��ǰ<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<select name="itemid"  id="itemcode" disabled  onChange="TnMasterItemSelect(this.value);">
				<option value="" selected>��ǰ�� �߰��� �ּ���.</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ ����<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="mastersellcash" id="mastersellcash" size="10" class="text" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ ������<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="masterdiscountrate" id="masterdiscountrate" value="" size="10" class="text" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">���� ī�װ�<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<table class="a">
			<tr>
				<td id="lyrDispList"><table class="a" id="tbl_DispCate"></table></td>
				<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
			</tr>
			</table>
			<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
		</td>
	</tr>
	<tr align="left">
	<td height="30" width="15%" bgcolor="<%= adminColor("tabletop") %>">���� ���� ���� </td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="adultType" value="0" checked>��ü����</label>
		<label><input type="radio" name="adultType" value="1" >���Žü�������</label>
		<label><input type="radio" name="adultType" value="2" >�̼��� ��ȸ �Ұ�</label>
	</td>
</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�˻� Ű����<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="keywords" id="keywords" size="80" maxlength="250" value="" class="text"> (�޸��α��� ex: Ŀ��,Ƽ����,����)
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����ī��</td>
		<td bgcolor="#FFFFFF">
			<textarea name="mainTitle" rows="4" cols="80"></textarea>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����ī��</td>
		<td bgcolor="#FFFFFF">
			<textarea name="subTitle" rows="4" cols="80"></textarea>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">PC ��� �̹���</td>
		<td bgcolor="#FFFFFF">
			<input class="button" type="button" value="�̹��� �ҷ�����" onClick="jsSetImg('dealcontents','spandealcontents');"/>
			(����,800X1600, Max 800KB,jpg,gif)
			<div id="spandealcontents"></div>
			<input type="hidden" name="addimggubun" value="1">
			<input type="hidden" name="addimgdel" value="">
			<input type="hidden" name="dealcontents" value="">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����� ��� �̹���</td>
		<td bgcolor="#FFFFFF">
			<input class="button" type="button" value="�̹��� �ҷ�����" onClick="jsSetImg('mobiledealcontents','spanmobiledealcontents');"/>
			(����,800X1600, Max 800KB,jpg,gif)
			<div id="spanmobiledealcontents"></div>
			<input type="hidden" name="mobileaddimggubun" value="2">
			<input type="hidden" name="mobileaddimgdel" value="">
			<input type="hidden" name="mobiledealcontents" value="">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">���</td>
		<td bgcolor="#FFFFFF">
			<textarea name="work_notice" rows="18" class="textarea" style="width:99%" id="[on,off,off,off][��ǰ����]"></textarea>
		</td>
	</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" height="25">
        <td valign="bottom" align="center">
			<input type="button" value="���" class="button" onClick="SubmitSave(this.form)">
			<input type="button" value="���" class="button" onClick="fnCancel()">
        </td>
    </tr>
</table>
</form>
<% if (application("Svr_Info")	= "Dev") then %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="100"></iframe>
<% else %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="100"></iframe>
<% end if %>
<script type="text/javascript">
<!--
	function fnCancel(){
		if(confirm("�Է��Ͻ� ������ �������� �ʰ� ����Ͻðڽ��ϱ�?")){
			location.href="/admin/itemmaster/deal/index.asp";
		}
	}
	//��ǰ �ִ� ������ ����
	function fnGetMaxSalevalue() {
		var idx = document.frm.idx.value;
		$.ajax({
			type: "POST",
			url: "ajaxGetDealMaxItemSalePer.asp",
			data: "idx="+idx,
			cache: false,
			success: function(message) {
				var splitmessage = message.split("|")
				if(message) {
					document.frm.masterdiscountrate.value=splitmessage[0];
					document.frm.discountitemid.value=splitmessage[1];
				} else {
					alert("��ǰ�� ���ų� �������� ��ǰ�� �����ϴ�.");
				}
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
	}

	//��ǰ ������ ����
	function fnGetMinPricevalue() {
		var idx = document.frm.idx.value;
		$.ajax({
			type: "POST",
			url: "ajaxGetDealMinItemPrice.asp",
			data: "idx="+idx,
			cache: false,
			success: function(message) {
				var splitmessage = message.split("|")
				if(message) {
					document.frm.mastersellcash.value=splitmessage[0];
					document.frm.saleitemid.value=splitmessage[1];
				} else {
					alert("��ǰ�� �����ϴ�.");
				}
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
	}

	// ����ī�װ� ���� �˾�
	function popDispCateSelect(){
		var dCnt = $("input[name='isDefault'][value='y']").length;
		$.ajax({
			url: "/common/module/act_DispCategorySelect.asp?isDft="+dCnt,
			cache: false,
			success: function(message) {
				$("#lyrDispCateAdd").empty().append(message).fadeIn();
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	// ���̾�� ����ī�װ� �߰�
	function addDispCateItem(dcd,cnm,div,dpt) {
		// ������ ���� �ߺ� ī�װ� ���� �˻�
		if(tbl_DispCate.rows.length>=2)	{
			alert("���� ī�װ��� �ִ� 2������ �Է°����մϴ�.");
			return false;
		}
		else
		{
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
		$("#catecnt").val($("#catecnt").val()+1);
		//��ǰ�Ӽ� ���
		printItemAttribute();
	}

	// ���� ����ī�װ� ����
	function delDispCateItem() {
		if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);
			$("#catecnt").val($("#catecnt").val()-1);
			//��ǰ�Ӽ� ���
			printItemAttribute();
		}
	}

	function printItemAttribute() {
		var arrDispCd="";
		$("input[name='catecode']").each(function(i){
			if(i>0) arrDispCd += ",";
			arrDispCd += $(this).val();
		});
		$.ajax({
			url: "/common/module/act_ItemAttribSelect.asp?itemid=0&arrDispCate="+arrDispCd,
			cache: false,
			success: function(message) {
				$("#lyrItemAttribAdd").empty().append(message);
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	function CheckImage(img, filesize, imagewidth, imageheight, extname, fsize)
	{
		var ext;
		var filename;

		filename = img.value;
		if (img.value == "") { return false; }

		if (CheckExtension(filename, extname) != true) {
			alert("�̹���ȭ���� ������ ȭ�ϸ� ����ϼ���.[" + extname + "]");
			ClearImage(img,fsize,imagewidth,imageheight);
			return false;
		}

		return true;
	}

	function ClearImage2(img,fsize,wd,ht) {
		img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg,gif', "+ fsize +");\" class='text' size='"+ fsize +"'>";
	}
//-->
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->