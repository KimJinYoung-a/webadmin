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
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 560   'KB
Dim idx, k, j, itemid
idx = requestCheckVar(Request("idx"),10)
If idx="" Then
Response.write "<script>alert('�� ������ �����ϴ�.');history.back();</script>"
Response.End
End If
Dim oDeal, oitem, oitemimg, arrIMG
set oDeal = new CDealView
oDeal.FRectMasterIDX = idx
oDeal.GetDealView

Dim oDealitem, arrList, iTotCnt, intLoop
set oDealitem = new CDealItem
oDealitem.FRectMasterIDX = idx
arrList = oDealitem.fnGetDealEventItem	
iTotCnt = oDealitem.FTotCnt	'��ü ������  ��
Set oDealitem=Nothing
itemid=oDeal.Fdealitemid
set oitem = new CItem
oitem.FRectItemID = oDeal.Fdealitemid
oitem.GetOneItem
Dim vArr
set oitemimg = new CItemAddImage
oitemimg.FRectItemID = oDeal.Fdealitemid
vArr = oitemimg.GetAddImageListIMGTYPE1

Function FormatDatePart(div,vdate)
	If div = "h" Then
		FormatDatePart=DatePart("h",vdate)
		If FormatDatePart<10 Then FormatDatePart="0"&FormatDatePart
	Else
		FormatDatePart=DatePart("n",vdate)
		If FormatDatePart<10 Then FormatDatePart="0"&FormatDatePart
	End If
End Function
%>
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
		var winpop = window.open('/admin/itemmaster/deal/pop_deal_addItems.asp?idx=<%=idx%>&stype=w','winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
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
		else if(!frm.sellyn[0].checked && !frm.sellyn[1].checked && !frm.sellyn[2].checked)
		{
			alert("�Ǹ� ���θ� �������ּ���.");
			return false;
		}
		else if(frm.itemid.value=="" && frm.isusing[0].checked)
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
		else
		{
			if(confirm("�Է��Ͻ� ������ ����ǰ�� �����Ͻðڽ��ϱ�?"))
			{
				frm.target="FrameCKP";
				frm.action="<%= ItemUploadUrl %>/linkweb/items/deal_itemeditWithImage_process.asp";
				frm.submit();
			}
		}		
	}

	function ClearImage2(img,fsize,wd,ht) {
		img.outerHTML = "<input type='file' name='" + img.name + "' onchange=\"CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg,gif', "+ fsize +");\" class='text' size='"+ fsize +"'>";
		$("#divaddimgname").hide();
		document.frm.addimgdel.value = "del";
	}

	function TnMasterItemSelect(itemid){
		if(document.frm.itemname.value=="")
		{
			document.frm.itemname.value=$("#itemcode option:selected").text();
		}
		$.ajax({
			url: "selectdealitemkeywords.asp?itemid="+itemid,
			cache: false,
			async: false,
			success: function(message) {
				//alert(message);
				if(message!="") {
					$('#keywords').val(message);
				} else {
					alert("���� �� ������ �����ϴ�.");
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
	function fnCancel(){
		if(confirm("�Է��Ͻ� ������ �������� �ʰ� ����Ͻðڽ��ϱ�?")){
			location.href="/admin/itemmaster/deal/index.asp";
		}
	}
	function editItemImage(itemid) {
		var param = "itemid=" + itemid;

		//if(makerid =="ithinkso"){
			//popwin = window.open('/common/pop_itemimage_ithinkso.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
		//}else{
			popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage','width=1000,height=900,scrollbars=yes,resizable=yes');
		//}
		popwin.focus();
	}	
//-->
</script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<form name="frm" method="post" onsubmit="return false;" style="margin:0px;"  enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="realitemid" value="<%=oDeal.Fdealitemid%>">
<input type="hidden" name="masteritemid" value="<%=oDeal.Fmasteritemcode%>">
<input type="hidden" name="auser" value="<%=session("ssBctId")%>">
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>�� �⺻ ����</B></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="itemname" id="itemname" size="80" maxlength="120" value="<%=oitem.FOneItem.Fitemname%>" class="text">
		</td>
	</tr>
	<tr>
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">���� �Ⱓ<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			 <input type="radio" name="viewdiv" id="viewdiv" value="1" onClick="TnViewDivSelect(1)"<% If oDeal.Fviewdiv="1" Then Response.write " checked" %>>��õ� <input type="radio" name="viewdiv" id="viewdiv" value="2" onClick="TnViewDivSelect(2)"<% If oDeal.Fviewdiv="2" Then Response.write " checked" %>>�Ⱓ��
			 <span id="datearea" style="display:<% If oDeal.Fviewdiv<>"2" Then Response.write "none" %>">
				<input id="startdate" name="startdate" value="<%=FormatDate(oDeal.Fstartdate,"0000-00-00")%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> <input type="text" name="shour" size="2" class="text" value="<% =FormatDatePart("h",oDeal.Fstartdate)%>"  maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');">:<input type="text" name="sminute" size="2" class="text" value="<%= FormatDatePart("n",oDeal.Fstartdate)%>"  maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');"> ~
				<input id="enddate" name="enddate" value="<%=FormatDate(oDeal.Fenddate,"0000-00-00")%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" /> <input type="text" name="ehour" size="2" class="text"  value="<%=FormatDatePart("h",oDeal.Fenddate)%>"  maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');">:<input type="text" name="eminute" size="2" class="text" value="<%=FormatDatePart("n",oDeal.Fenddate)%>"  maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');">
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
			<input type="radio" name="isusing" id="isusing" value="Y"<% If oDeal.Fisusing="Y" Then Response.write " checked" %>>��� <input type="radio" name="isusing" id="isusing" value="N"<% If oDeal.Fisusing="N" Then Response.write " checked" %>>��� ����
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ǹſ���<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="radio" name="sellyn" id="sellyn" value="Y"<% If oDeal.Fsellyn="Y" Then Response.write " checked" %>>��� <input type="radio" name="sellyn" id="sellyn" value="S"<% If oDeal.Fsellyn="S" Then Response.write " checked" %>>�Ͻ� ǰ�� <input type="radio" name="sellyn" id="sellyn" value="N"<% If oDeal.Fsellyn="N" Then Response.write " checked" %>>�Ǹ� ����
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
			<input type="button" class="button" style="width:105;" value="�˻�" onclick="TnSearchObjOpenWin('Just1Day_list.asp');">&nbsp;<b style="color:red">*</b>�� ��ǰ�� �˻��Ͽ� �߰��� �ּ���.
			<div id="divForm">
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
			<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
				<td>����</td>
				<td>��ǰ�ڵ�</td>
				<td>��ǰ��</td>
				<td>�ǸŰ�</td>
				<td>���԰�</td>
				<td>������</td>
			</tr>
			<% If isArray(arrList) Then %>
			<% For intLoop = 0 To UBound(arrList,2) %>
			<tr bgcolor="#FFFFFF" align="center">
				<td><%=arrList(0,intLoop)%></td>
				<td><a href="javascript:editItemBasicInfo('<%=arrList(1,intLoop)%>')"><%=arrList(1,intLoop)%></a></td>
				<td><%=arrList(2,intLoop)%></td>
				<td>
					<%
						Response.Write FormatNumber(arrList(5,intLoop),0)
						'���ΰ�
						if arrList(9,intLoop)="Y" then
							Response.Write "<br><font color=#F08050>(��)" & FormatNumber(arrList(7,intLoop),0) & "</font>"
						end if
						'������
						if arrList(10,intLoop)="Y" then
							Select Case arrList(11,intLoop)
								Case "1"
									Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(4,intLoop)*((100-arrList(12,intLoop))/100),0) & "</font>"
								Case "2"
									Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(4,intLoop)-arrList(12,intLoop),0) & "</font>"
							end Select
						end if
					%>
				</td>
				<td>
					<%
						Response.Write FormatNumber(arrList(6,intLoop),0)
						'���ΰ�
						if arrList(9,intLoop)="Y" then
							Response.Write "<br><font color=#F08050>" & FormatNumber(arrList(8,intLoop),0) & "</font>"
						end if
						'������
						if arrList(10,intLoop)="Y" then
							if arrList(12,intLoop)="1" or arrList(12,intLoop)="2" then
								if arrList(13,intLoop)=0 or isNull(arrList(13,intLoop)) then
									Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(6,intLoop),0) & "</font>"
								else
									Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(13,intLoop),0) & "</font>"
								end if
							end if
						end if
					%>
				</td>
				<td>
					<a href="javascript:fnSaleInfo();"><%if arrList(9,intLoop)="Y" then%>
					<font color="#F08050"><%=CLng(((arrList(5,intLoop)-arrList(7,intLoop))/arrList(5,intLoop))*100)%>%</font>		
					<%end if%>
					<%if arrList(10,intLoop)="Y" then 
					if arrList(12,intLoop)="1" or arrList(12,intLoop)="2" then
						if arrList(13,intLoop)=0 or isNull(arrList(13,intLoop)) then
							 Response.Write "<br><font color=#5080F0>" & FormatNumber( arrList(6,intLoop),0) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(12,intLoop),0) 
							 if arrList(12,intLoop)="1" then 
							 Response.Write "%"
							else
							 Response.Write "��"
							end if
							 Response.Write "</font>"
						end if
					end if
					end if%></a>
				</td>
			</tr>
			<% Next %>
			<% End If %>
			</table>
			</div>
			<div id="divFrm3"></div>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ ��ǰ<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<% If isArray(arrList) Then %>
			<select name="itemid" id="itemcode" onChange="TnMasterItemSelect(this.value);">
				<option value="" selected>��ǰ�� ������ �ּ���.</option>
				<% For intLoop = 0 To UBound(arrList,2) %>
				<option value="<%=arrList(1,intLoop)%>"<% If arrList(1,intLoop) = oDeal.Fmasteritemcode Then  Response.write " selected"%>><%=arrList(2,intLoop)%></option>
				<% Next %>
			</select>
			<% Else %>
			<select name="itemid"  id="itemcode" disabled  onChange="TnMasterItemSelect();">
				<option value="" selected>��ǰ�� �߰��� �ּ���.</option>
			</select>
			<% End If %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����ǰ �̹�������</td>
		<td bgcolor="#FFFFFF">
			<input type="button" value="�̹�������" class="button" onClick="editItemImage('<%= itemid %>')">
		</td>
	</tr>	
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ ����, ����<br>������ �ڵ�</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="saleitemid" size="10" class="text" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');">
			<input type="text" name="discountitemid" size="10" class="text" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');">
			(����� ������ �ڵ带 �Է��ص� ����,���������� ���� �� �� �ֽ��ϴ�.)
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ ����<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="mastersellcash" size="10" class="text" value="<%=oDeal.Fmastersellcash%>" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');">&nbsp;<input type="button" value="��������" class="button" onClick="fnGetMinPricevalue()" id="saleper1" name="saleper1" style="display:<% If Not isArray(arrList) Then %>none<% End If %>"><!-- &nbsp;<input type="checkbox" name="pricesdash" value="Y"<% If oDeal.Fpricesdash ="Y" Then Response.write " checked" %>>"~"���� ���� (��: 19,900��~) -->
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��ǥ ������<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="masterdiscountrate" id="masterdiscountrate" value="<%=oDeal.Fmasterdiscountrate%>" size="10" class="text" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\��-����-�Ӱ�-�R]/g, '');">&nbsp;<input type="button" value="��������" class="button" onClick="fnGetMaxSalevalue()" id="saleper2" name="saleper2" style="display:<% If Not isArray(arrList) Then %>none<% End If %>"><!-- &nbsp;<input type="checkbox" name="sailsdash" value="Y"<% If oDeal.Fsailsdash ="Y" Then Response.write " checked" %>>"~"���� ���� (��: ~77%) -->
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">���� ī�װ�<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<table class="a">
			<tr>
				<td id="lyrDispList"><%=getDispCategory(oDeal.Fdealitemid)%></td>
				<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
			</tr>
			</table>
			<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
		</td>
	</tr>
	<tr align="left">
	<td height="30" width="15%" bgcolor="<%= adminColor("tabletop") %>">���� ���� ���� </td>
	<td bgcolor="#FFFFFF" colspan="2">
		<label><input type="radio" name="adultType" value="0" <%=chkIIF(oitem.FOneItem.FadultType=0,"checked","")%>>��ü����</label>
		<label><input type="radio" name="adultType" value="1" <%=chkIIF(oitem.FOneItem.FadultType=1,"checked","")%>>�̼��� ��ȸ �Ұ�</label>
		<label><input type="radio" name="adultType" value="2" <%=chkIIF(oitem.FOneItem.FadultType=2,"checked","")%>>���Žü�������</label>
	</td>
</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�˻� Ű����<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="keywords" id="keywords" size="80" maxlength="250" value="<%=oitem.FOneItem.Fkeywords%>" class="text"> (�޸��α��� ex: Ŀ��,Ƽ����,����)
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��� �̹���</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="addimgname" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, 'jpg,gif',40);" class="text" size="40">
			<input type="button" value="#1 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname,40, 800, 1600)"> (����,800X1600, Max 800KB,jpg,gif)<br>
			<input type="hidden" name="addimggubun" value="1">
			<input type="hidden" name="addimgdel" value="">
			<%
			If isArray(vArr) Then
					If vArr(3,UBound(vArr,2)) > 0 Then
					For k = 1 To vArr(3,UBound(vArr,2))
		    			If CStr(vArr(3,j)) = CStr(k) AND (vArr(4,j) <> "" and isNull(vArr(4,j)) = False) Then
							Response.Write "<div id=""divaddimgname"" style=""display:block;""><img src=""" & webImgUrl & "/item/contentsimage/" & GetImageSubFolderByItemid(vArr(1,j)) & "/" & vArr(4,j) & """ height=""250""></div>"
							Exit For
		    			End If
					Next
					End If
			End If
			%>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">���</td>
		<td bgcolor="#FFFFFF">
			<textarea name="itemcontent" rows="18" class="textarea" style="width:99%" id="[on,off,off,off][��ǰ����]"><%=oDeal.Fwork_notice%></textarea>
		</td>
	</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" height="25">
        <td valign="bottom" align="center">
			<input type="button" value="����" class="button" onClick="SubmitSave(this.form)">
			<input type="button" value="���" class="button" onClick="fnCancel()">
        </td>
    </tr>
</table>
</form>
<% if (application("Svr_Info")	= "Dev") then %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="100"></iframe>
<% else %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>
<script type="text/javascript">
<!--
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

		//��ǰ�Ӽ� ���
		printItemAttribute();
	}

	// ���� ����ī�װ� ����
	function delDispCateItem() {
		if(confirm("������ ī�װ��� �����Ͻðڽ��ϱ�?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

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
//-->
</script>
<%
Set oitem = Nothing
Set oDeal = Nothing
Set oitemimg = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->