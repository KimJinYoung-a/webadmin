<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : Ŭ����� ���� ����
'	History		: 2016.01.14 ���¿� ����
'               : 2022.02.11 ���� �ٿ�ޱ�; ������
'#############################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/ClearanceSale/ClearanceSaleCls.asp"-->

<%
Dim i, idx
Dim FResultCount, iTotCnt, iCurrentpage
Dim itemid, rectitemid, itemname, makerid, usingyn, sellyn, limityn, catecode, sailyn, itemcouponyn
dim iSalePercent

	idx = request("idx")
	iCurrentpage = NullFillWith(requestCheckVar(Request("IC"),10),1)
	itemid      = requestCheckvar(request("itemid"),255)
	rectitemid  = requestCheckvar(request("rectitemid"),255)
	itemname    = requestCheckvar(request("itemname"),64)
	makerid     = requestCheckvar(request("makerid"),32)
	sellyn      = requestCheckvar(request("sellyn"),10)
	usingyn     = requestCheckvar(request("usingyn"),10)
	limityn     = requestCheckvar(request("limityn"),10)
	catecode    = requestCheckvar(request("catecode"),10)
	sailyn      = requestCheckvar(request("sailyn"),10)
	itemcouponyn = requestCheckvar(request("itemcouponyn"),10)

if iCurrentpage="" then iCurrentpage=1

if rectitemid<>"" then
	dim iA ,arrTemp,arrrectitemid
  rectitemid = replace(rectitemid,chr(13),"")
	arrTemp = Split(rectitemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrrectitemid = arrrectitemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrrectitemid)>0 then
		rectitemid = left(arrrectitemid,len(arrrectitemid)-1)
	else
		if Not(isNumeric(rectitemid)) then
			rectitemid = ""
		end if
	end if
end if

dim oclear
set oclear = new CClaearanceitem
	oclear.FPageSize = 20
	oclear.FRectItemid		= rectitemid
	oclear.FRectSellYN		= sellyn
	oclear.FRectIsusing		= usingyn
	oclear.FRectMakerid		= makerid
	oclear.FRectLimityn		= limityn
	oclear.FRectCatecode		= catecode
	oclear.FRectSaleYN		= sailyn
	oclear.FRectItemcouponYN	= itemcouponyn
	oclear.FRectitemname	= itemname
	oclear.FCurrPage = iCurrentpage
	oclear.fnGetclaearanceitemList
iTotCnt = oclear.FTotalCount
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function itemwrite(){
	if(frmitem.itemid.value==""){
		alert('��ǰ�ڵ带 �Է��� �ּ���.');
		frmitem.itemid.focus();
		return;
	}
    var ff = document.frmitem; 
    var itemid = "itemid"; 
    var cnt = document.getElementsByName(itemid); 
    var totalCnt = 0;
    var replacechksame = document.frmitem.itemid.value.replace(/\r\n/g, ",");
		replacechksame = replacechksame.replace(/\s/g,''); // ���� ���� 
    var chksame = replacechksame.split(","); 
    for(var j=0; j < chksame.length; j++) { 
        var tmp = chksame[j];
        tmp.replace(/s/gi, "");
        for(var k=j+1; k <= chksame.length; k++) { 
            if (tmp == chksame[k]) { 
                alert('�ߺ� �� ���� �ֽ��ϴ�. Ȯ�� �� �ּ���\n�ߺ� ��ǰ�ڵ� : '+chksame[k]); 
                chkfocus(); 
                return; 
            } 
        } 
    } 

	frmitem.submit();
}

//��뿩�� ����
function jsSortIsusing() {
//	alert(document.fitem.isusing.length);
//	return

	var frm;
	var sValue, isusing;
	frm = document.fitem;
	sValue = "";
	isusing = "";
	dispcate1 = "";
	chkSel	= 0;

//alert(frm.chkI.value);
//return;

	if (frm.chkI.length > 1){
		for (var i=0;i<frm.isusing.length;i++){
			if(frm.chkI[i].checked) chkSel++;

			if (frm.isusing[i].value ==''){
				alert('��뿩�θ� �����ϼ���.');
				frm.isusing[i].focus();
				return;
			}

			if (frm.dispcate1[i].value ==''){
				alert('ī�װ��� �����ϼ���.');
				frm.dispcate1[i].focus();
				return;
			}

			//IDX
			if (sValue==""){
				sValue = frm.chkI[i].value;
			}else{
				sValue =sValue+","+frm.chkI[i].value;
			}

			// ��뿩��
			if (isusing==""){
				isusing = frm.isusing[i].value;
			}else{
				isusing =isusing+","+frm.isusing[i].value;
			}

			// ī�װ�
			if (dispcate1==""){
				dispcate1 = frm.dispcate1[i].value;
			}else{
				dispcate1 =dispcate1+","+frm.dispcate1[i].value;
			}
		}
	}else{
		if (frm.isusing.value ==''){
			alert('��뿩�θ� �����ϼ���.');
			frm.isusing.focus();
			return;
		}
		if(frm.chkI.checked) chkSel++;
		if(frm.chkI.checked){
			sValue = frm.chkI.value;
			isusing =  frm.isusing.value;
			dispcate1 = frm.dispcate1.value;
		}
	}
	if(chkSel<=0) {
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}

	document.frmSortIsusing.isusingarr.value = isusing;
	document.frmSortIsusing.dispcate1arr.value = dispcate1;
	document.frmSortIsusing.idxarr.value = sValue;
	document.frmSortIsusing.mode.value = "sortisusingedit";
	document.frmSortIsusing.submit();
}

//���� ��ǰ ����
function jsDelete() {
	var frm;
	var sValue, isusing;
	frm = document.fitem;
	sValue = "";
	chkSel	= 0;

	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(frm.chkI[i].checked) {
				chkSel++;

				//IDX
				if (sValue==""){
					sValue = frm.chkI[i].value;
				}else{
					sValue =sValue+","+frm.chkI[i].value;
				}
			}
		}
	}else{
		if(frm.chkI.checked) chkSel++;
		if(frm.chkI.checked){
			sValue = frm.chkI.value;
		}
	}
	if(chkSel<=0) {
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}

	if(confirm("������ " + chkSel + "�� ��ǰ�� �����Ͻðڽ��ϱ�?\n\n�� ���� ���� ������ �Ұ����ϸ�, ���� �� �� �ֽ��ϴ�.")) {
		document.frmSortIsusing.idxarr.value = sValue;
		document.frmSortIsusing.mode.value = "itemDelete";
		document.frmSortIsusing.submit();
	}
}

function showimage(img){
	var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
}

//üũ�ڽ� ��ü ����
var ichk;
ichk = 1;
function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.type=="checkbox")) {
			e.checked = blnChk ;
		}
	}
}

//��뿩�� ��ü ����
function jsIsusingChg(selv) {
    var frm, blnChk;
	frm = document.fitem;
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.isusing.length;i++){
			frm.isusing[i].value=selv;
		}
	}else{
		frm.isusing.value=selv;
	}
}

// �˻�
function jsSearch(p){
	frm.iC.value = p;
	document.frm.target="_self";
	document.frm.action="index.asp";
	document.frm.submit();
}

//��ǰ����Ʈ �ٿ�
function jsItemDown(){
  document.frm.page.value = $('#selDCnt').val();
	document.frm.target="hidifr";
	document.frm.action="index_excel.asp";
	document.frm.submit();
}
</script>

 <% '�˻�-------------------------------------------------------------------------------------------- %>
<table width="65%" height="120" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="GET">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<input type="hidden" name="iC" value="1">
	<tr align="left" bgcolor="<%= adminColor("topbar") %>" >
		<td align="center" width="50" bgcolor="<%= adminColor("gray") %>"><b>�˻�<br>����</b></td>
		<td>
			<table border="0" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td colspan="2" style="white-space:nowrap;">
						�귣��: <%	drawSelectBoxDesignerWithName "makerid", makerid %>&nbsp;&nbsp;
						��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32">
					</td>
					<td rowspan="3" align="center" style="white-space:nowrap;padding-left:5px;">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						��ǰ�ڵ� : <textarea rows="7" cols="15" name="rectitemid" id="rectitemid"><%=replace(rectitemid,",",chr(10))%></textarea>
					</td>
				</tr>
				<tr>
					<td>
						<span style="white-space:nowrap;">�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %></span>&nbsp;
				     	<span style="white-space:nowrap;">����:<% drawSelectBoxLimitYN "limityn", limityn %></span>&nbsp;
				     	<span style="white-space:nowrap;">���:<% drawSelectBoxUsingYN "usingyn", usingyn %></span>&nbsp;&nbsp;&nbsp;
						
						<select name="catecode" class="select">
							<option value="">��ü ī�װ�</option>
							<option value="999" style="background:red;color:white;" <% if catecode="999" then response.write "selected" %>>ī�װ� ����</option>
							<option value="101" <% if catecode="101" then response.write "selected" %>>�����ι���</option>
							<option value="102" <% if catecode="102" then response.write "selected" %>>������/�ڵ���</option>
							<option value="103" <% if catecode="103" then response.write "selected" %>>ķ��/Ʈ����</option>
							<option value="104" <% if catecode="104" then response.write "selected" %>>����</option>
							<option value="121" <% if catecode="121" then response.write "selected" %>>����/����</option>
							<option value="122" <% if catecode="122" then response.write "selected" %>>����/�ö��</option>
							<option value="120" <% if catecode="120" then response.write "selected" %>>�к긯/����</option>
							<option value="112" <% if catecode="112" then response.write "selected" %>>Űģ</option>
							<option value="119" <% if catecode="119" then response.write "selected" %>>Ǫ��</option>
							<option value="117" <% if catecode="117" then response.write "selected" %>>�м��Ƿ�</option>
							<option value="116" <% if catecode="116" then response.write "selected" %>>����/����/�־�</option>
							<option value="118" <% if catecode="118" then response.write "selected" %>>��Ƽ</option>
							<option value="115" <% if catecode="115" then response.write "selected" %>>���̺�/Ű��</option>
							<option value="110" <% if catecode="110" then response.write "selected" %>>Cat&Dog</option>
						</select>
				    </td>
				    <td>
				    </td>
				</tr>
				<tr>
					<td>
						<span style="white-space:nowrap;">���� <% drawSelectBoxSailYN "sailyn", sailyn %></span>&nbsp;
						<span style="white-space:nowrap;">���� <% drawSelectBoxCouponYN "itemcouponyn", itemcouponyn %></span>
					</td>
					<td></td>
				</tr>
			</table>
		</td>

		<td align="center" rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="jsSearch(1)">
		</td>
	</tr>
    </form>
</table>
<% '�˻���-------------------------------------------------------------------------------------------- %>
<form name="frmSortIsusing" method="post" action="/admin/clearancesale/itemRegProc.asp" style="margin:0px;">
	<input type="hidden" name="isusingarr" value="">
	<input type="hidden" name="dispcate1arr" value="">
	<input type="hidden" name="idxarr" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="mode" value="sortisusingedit">
</form>
<% '��ǰ�Է�------------------------------------------------------------------------------------ %>
<form name="frmitem" method="post" action="itemRegProc.asp">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="mode" value="iteminsert">
<table width="35%"  height="121" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="left" bgcolor="<%= adminColor("topbar") %>" >
		<td align="center" width="50" bgcolor="<%= adminColor("gray") %>"><b>��ǰ<br>�Է�</b></td>
		<td align="center"><textarea rows="7" cols="15" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea></td>
		<td>
			<font color="red"><strong>�� ���ٿ� ��ǰ�ڵ� 1����<br>�� �ѹ��� �ִ�10������ ��� ���� �մϴ�<br>�� ��ϵ� ��ǰ�� ����ī�װ��� �����˴ϴ�<br>�� ���� �������3~10�� �ҿ�˴ϴ�<br></strong></font>
			<input type="button" name="newBT" class="button" value="��ǰ ���" onclick="itemwrite(); return false;">
		</td>
	</tr>
</table>
</form>
<% '��ǰ�Է� ��---------------------------------------------------------------------------- %>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="right">
		<%dim   imax, imin%>
		<select name="selDCnt" id="selDCnt" class="select" style="height:25px;vertical-align:top;">
			<%for i =1 To Int(oclear.FTotalCount/5000)+1
					imin = ((i-1)*5000)+1
					if i <  Int(oclear.FTotalCount/5000)+1 then
					imax = i*5000
					else
					imax = oclear.FTotalCount
					end if
			%>
			<option value="<%=i%>"><%=imin%>~<%=imax%></option>
			<%Next%>
		</select>
		<input type="button" class="button" value="��ǰ�ٿ�ε�(����)" onclick="jsItemDown();">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<% '����Ʈ--------------------------------------------------------------------------------------------- %>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=adminColor("tablebg")%>">
<form name="fitem" method="post" style="margin:0px;">
	<tr align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
		<td  colspan="14" align="left"><b>Total : <%= iTotCnt %></b></td>
	</tr>
	
	<tr align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
		<td width="60" rowspan="2"><strong>��ǰ �ڵ�</strong></td>
		<td width="50" rowspan="2"><strong>�̹���</strong></td>
		<td rowspan="2"><strong>�귣��</strong></td>
		<td rowspan="2"><strong>��ǰ��</strong></td>
		<td rowspan="2">��౸��</td>
		<td rowspan="2">���λ���</td>
		<td rowspan="2">�ǸŰ�</td>
		<td rowspan="2">���԰�</td>
		<td rowspan="2">������</td> 
		<td rowspan="2">������</td> 
		<td rowspan="2"><strong>�Ǹſ���</strong></td>
		<td rowspan="2"><strong>��������</strong></td>
		<td width="300" colspan="2"><strong>��뿩�� & ī�װ�</strong></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20" ><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
		<td colspan="2"  width="200">
			<!--
			<select name="selisusing" onchange="jsIsusingChg(this.value)" class="select">
				<option value="">����</option>
				<option value="N">N</option>
				<option value="Y">Y</option>
			</select>

			<input class="button" type="button" id="btnEditSel" value="����" onClick="jsSortIsusing();" /> &nbsp;
			-->
			<input class="button_auth" type="button" id="btnDelSel" value="���û�ǰ ����" onClick="jsDelete();" />
			<br><font color="red">�� ���� �������3~10�� �ҿ�˴ϴ�</font>
		</td>
	</tr>
	
	<% if oclear.FResultCount > 0 then %>
		<% for i = 0 to oclear.FResultCount - 1 %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF'; height="30"> 

			<%''��ǰ�ڵ�%>
			<td><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oclear.FItemList(i).Fitemid %>" target="_blank" ><%= oclear.FItemList(i).Fitemid %></td>

			<%''�̸����� �̹���%>
			<td><img src="<%= db2html(oclear.FItemList(i).Flistimage) %>" width="50" height="50" style="cursor:pointer" onclick="showimage('<%= db2html(oclear.FItemList(i).Fbasicimage) %>');"></td>

			<%''�귣��� %>
			<td><%= oclear.FItemList(i).Fmakerid %></td>
			
			<%''��ǰ�� %>
			<td><%= oclear.FItemList(i).Fitemname %></td>

			<%''��౸�� %>
			<td><%=fnColor(oclear.FItemList(i).FmwDiv,"mw") %></td>
			<%''���λ��� %>
			<td><%=fnColor(oclear.FItemList(i).Fsaleyn,"yn") %></td>
			<%''�ǸŰ� %>
			<td><%=FormatNumber(oclear.FItemList(i).ForgPrice,0)%>
					<% 		'���ΰ�(������=(�Һ��ڰ�-���ΰ�)/�Һ��ڰ�*100) 
					if oclear.FItemList(i).Fsaleyn ="Y" then %>
					<br><font color=#F08050>(<%=CLng((oclear.FItemList(i).ForgPrice-oclear.FItemList(i).FsellCash)/oclear.FItemList(i).ForgPrice*100) %>%��)<%=FormatNumber(oclear.FItemList(i).FsellCash,0)%></font>
					<% end if %>
					<%'������ 
					if oclear.FItemList(i).FitemcouponYn="Y" then
					 
						Select Case oclear.FItemList(i).FitemcouponType
							Case "1" '% ����
					%>
						<br><font color=#5080F0>(��)<%=FormatNumber(oclear.FItemList(i).FsellCash-(CLng(oclear.FItemList(i).FsellCash*oclear.FItemList(i).FitemcouponValue/100)),0)%></font>  
					<%
							Case "2" '�� ����
					%>		
						<br><font color=#5080F0>(��)<%=FormatNumber(oclear.FItemList(i).FsellCash-oclear.FItemList(i).FitemcouponValue,0)%></font>
					<%			
						end Select
					end if
					
					'������
					iSalePercent = (1-(clng(oclear.FItemList(i).FsellCash)/clng(oclear.FItemList(i).ForgPrice)))*100
					%> 
			</td>
			<%''���԰� %>
			<td><%=FormatNumber(oclear.FItemList(i).ForgSuplyCash,0)%>
				<% '���ΰ�
					if oclear.FItemList(i).Fsaleyn ="Y" then
				%>		
					 <br><font color=#F08050><%=FormatNumber(oclear.FItemList(i).FsailSuplyCash,0) %></font> 
				<%
					end if
					'������
					if  oclear.FItemList(i).FitemcouponYn="Y" then
						if oclear.FItemList(i).FitemcouponType="1" or oclear.FItemList(i).FitemcouponType="2" then
							if  oclear.FItemList(i).FitemcouponBuyPrice=0 or isNull(oclear.FItemList(i).FitemcouponBuyPrice) then
								Response.Write "<br><font color=#5080F0>" & FormatNumber(oclear.FItemList(i).FbuyCash,0) & "</font>"
							else
								Response.Write "<br><font color=#5080F0>" & FormatNumber(oclear.FItemList(i).FitemcouponBuyPrice,0) & "</font>"
							end if
						end if
					end if
			%>
			</td>
			<%''������ %>
			<td>
				<%=fnPercent(oclear.FItemList(i).ForgSuplyCash,oclear.FItemList(i).ForgPrice,1)%>
				<%
					'���ΰ�
					if oclear.FItemList(i).Fsaleyn ="Y"  then
						Response.Write "<br><font color=#F08050>" & fnPercent(oclear.FItemList(i).FsailSuplyCash,oclear.FItemList(i).FsailPrice,1) & "</font>"
					end if
					'������
					if oclear.FItemList(i).FitemcouponYn="Y" then
						Select Case  oclear.FItemList(i).FitemcouponType
							Case "1"
								if oclear.FItemList(i).FitemcouponBuyPrice=0 or isNull(oclear.FItemList(i).FitemcouponBuyPrice) then
									Response.Write "<br><font color=#5080F0>" & fnPercent(oclear.FItemList(i).FbuyCash,oclear.FItemList(i).FsellCash-(CLng(oclear.FItemList(i).FitemcouponValue*oclear.FItemList(i).FsellCash/100)),1) & "</font>"
								else
									Response.Write "<br><font color=#5080F0>" & fnPercent(oclear.FItemList(i).FitemcouponBuyPrice,oclear.FItemList(i).FsellCash-(CLng(oclear.FItemList(i).FitemcouponValue*oclear.FItemList(i).FsellCash/100)),1) & "</font>"
								end if
							Case "2"
								if oclear.FItemList(i).FitemcouponBuyPrice=0 or isNull(oclear.FItemList(i).FitemcouponBuyPrice) then
									Response.Write "<br><font color=#5080F0>" & fnPercent(oclear.FItemList(i).FbuyCash,oclear.FItemList(i).FsellCash,1) & "</font>"
								else
									Response.Write "<br><font color=#5080F0>" & fnPercent(oclear.FItemList(i).FitemcouponBuyPrice,oclear.FItemList(i).FsellCash,1) & "</font>"
								end if
						end Select 
				end if
			%>
			</td> 
			<%''������ %>
			<td style="<%=chkIIF(iSalePercent>=50,"color:#EE0000;font-weight:bold;","")%>"><%=formatnumber(iSalePercent,0)%> %</td>

			<%''�Ǹſ��� %>
			<td><%= oclear.FItemList(i).Fsellyn %></td> <% '�Ǹſ���%>
			
			<%''�������� %>
			<td><%= oclear.FItemList(i).Flimityn %></td> <% '�������� %>
			
			<%''��뿩�� %>
			<td><input type="checkbox" name="chkI" onClick="AnCheckClick(this);"  value="<%= oclear.FItemlist(i).Fidx %>"></td>
			<td>
				<input type="hidden" value="<%= oclear.FItemList(i).FIsusing %>" name="orgisusing">
				<input type="hidden" name="limitynchk" value="<%= oclear.FItemList(i).Fidx %>">
				<input type="hidden" name="dispcate1" value="<%= oclear.FItemList(i).Fdispcate1 %>">
				<table border='0' cellspacing="0" cellpadding="3" class="a">
				<tr>
					<td>
						<% ''drawSelectBoxUsingYN "isusing", oclear.FItemList(i).FIsusing %>
						<%=chkIIF(oclear.FItemList(i).FIsusing="Y","�����","������")%> /
					</td>
					<td>
					<%= oclear.FItemList(i).FdispCateName %>
					<%
						if Not(oclear.FItemList(i).FdispCateNameReal="" or isNull(oclear.FItemList(i).FdispCateNameReal)) then
							Response.Write "<br /><span style=""color:#999;"">(" & oclear.FItemList(i).FdispCateNameReal & ")</span>"
						end if
					%>
					</td>
				</tr>
				</table>
			</td>
		</tr>

		<% next %>
		<% '����¡-------------------------------------------- %>
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="14" align="center">
		       	<% if oclear.HasPreScroll then %>
					<span class="list_link"><a href="javascript:jsSearch('<%= oclear.StartScrollPage-1 %>')">[pre]</a></span>
				<% else %>
				[pre]
				<% end if %>
					<% for i = 0 + oclear.StartScrollPage to oclear.StartScrollPage + oclear.FScrollCount - 1 %>
						<% if (i > oclear.FTotalpage) then Exit for %>
						<% if CStr(i) = CStr(iCurrentpage) then %>
							<span class="page_link"><font  size="3" color="red"><b><%= i %></b></font></span>&nbsp;
						<% else %>
							<a href="javascript:jsSearch('<%= i %>')" class="list_link"><font color="#000000" size="3"><%= i %></font></a>&nbsp;
						<% end if %>
					<% next %>
				<% if oclear.HasNextScroll then %>
					<span class="list_link"><a href="javascript:jsSearch('<%= i %>')">[next]</a></span>
				<% else %>
				[next]
				<% end if %>
			</td>
		</tr>
		<% '����¡��------------------------------------ %>
	<% else %>	
		<tr>
			<td colspan=14 align="center">
				���̻� �����ϴ�.
			</td>
		</tr>
	<% end if %>
</form>
</table>
<iframe id="hidifr" src="" width="0" height="0" frameborder="0"></iframe>
<% ''����Ʈ ��------------------------------------------------%>
<%
set oclear = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->