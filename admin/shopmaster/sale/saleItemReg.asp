<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� ��ǰ ����
' History : 2008.04.08 ������ ����
'           2013.06.21 ������ / ������ ǥ�� �� ��� �߰�
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/itemsalecls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim sCode, clsSale,clsSaleItem
Dim sTitle,isRate, isMargin, isStatus,eCode, egCode, dSDay, dEDay, isUsing, dOpenDay,isMValue, smargin
Dim acURL
Dim iTotCnt, arrList,intLoop
Dim iPageSize, iCurrpage ,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
dim makerid, sailyn,invalidmargin, sRectItemidArr 

sCode = requestCheckVar(Request("sC"),10)
makerid =  requestCheckVar(Request("makerid"),32)
sailyn	=  requestCheckVar(Request("sailyn"),1)
invalidmargin=  requestCheckVar(Request("invalidmargin"),1)
sRectItemidArr=  requestCheckVar(Request("sRectItemidArr"),400)

acURL =Server.HTMLEncode("/admin/shopmaster/sale/saleitemProc.asp?sC="&sCode)

if sRectItemidArr<>"" then
	dim iA ,arrTemp,arrItemid
	sRectItemidArr = replace(sRectItemidArr,",",chr(10)) 
	sRectItemidArr = replace(sRectItemidArr,chr(13),"") 
	arrTemp = Split(sRectItemidArr,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp) 
		if trim(arrTemp(iA))<>"" then 
			'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.05;������)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	sRectItemidArr = left(arrItemid,len(arrItemid)-1)
end if

'�������¿� ���� ���԰� ����-------------------------------------------------------
Function fnSetSaleSupplyPrice(ByVal MarginType, ByVal MarginValue, ByVal orgPrice, ByVal orgSupplyPrice, ByVal salePrice)
	Dim orgMRate
	if orgPrice <>0 then '�� ������
		orgMRate = 100-fix(orgSupplyPrice/orgPrice*10000)/100
	end if

	SELECT CASE MarginType
		Case 1	'���ϸ���
			fnSetSaleSupplyPrice = salePrice- fix(salePrice*(orgMRate/100))
		Case 2	'��ü�δ�
			fnSetSaleSupplyPrice = salePrice-(orgPrice-orgSupplyPrice)
		Case 3	'�ݹݺδ�
			fnSetSaleSupplyPrice = orgSupplyPrice- fix((orgPrice-salePrice)/2)
		Case 4	'10x10�δ�
			fnSetSaleSupplyPrice = orgSupplyPrice
		Case 5	'��������
			fnSetSaleSupplyPrice = salePrice - fix(salePrice*(MarginValue/100))
	END SELECT
End Function
'-----------------------------------------------------------------------------------
 
'���� �⺻����
set clsSale = new CSale
clsSale.FSCode  = sCode
clsSale.fnGetSaleConts

sTitle 		= clsSale.FSName
isRate 		= clsSale.FSRate
isMargin 	= clsSale.FSMargin
eCode 		= clsSale.FECode
egCode		= clsSale.FEGroupCode
dSDay 		= clsSale.FSDate
dEDay 		= clsSale.FEDate
isStatus 	= clsSale.FSStatus
isUsing     = clsSale.FSUsing
dOpenDay	= clsSale.FOpenDate
isMValue	= clsSale.FSMarginValue
set clsSale = nothing

iCurrpage = Request("iC")	'���� ������ ��ȣ
IF iCurrpage = "" THEN	iCurrpage = 1
iPageSize = 20		'�� �������� �������� ���� ��
iPerCnt = 10		'�������� ������ ����

'���� ��ǰ����
set clsSaleItem = new CSaleItem
clsSaleItem.FCPage = iCurrpage
clsSaleItem.FPSize = iPageSize
clsSaleItem.FSCode = sCode
clsSaleItem.FRectMakerid = makerid
clsSaleItem.FRectsailyn = sailyn
clsSaleItem.FRectinvalidmargin =invalidmargin
clsSaleItem.FRectItemidArr = sRectItemidArr 
arrList = clsSaleItem.fnGetSaleItemList
iTotCnt = clsSaleItem.FTotCnt	'��ü ������  ��

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��

'���Ⱓ�� ��ǰ���� ���� ����
Dim arrItemCoupon, iclp
arrItemCoupon = clsSaleItem.fnGetCouponListBySaleInfo

'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
Dim arrsalemargin, arrsalestatus
arrsalemargin = fnSetCommonCodeArr("salemargin",False)
arrsalestatus= fnSetCommonCodeArr("salestatus",False)
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
<!--
// ������ �̵�
function jsGoPage(iP){
	location.href="saleItemReg.asp?menupos=<%=menupos%>&sC=<%=sCode%>&iC="+iP;
}

// ����ǰ �߰� �˾�
function addnewItem(eC,egC){
		var popwin;
		if ( eC > 0 ){
			popwin = window.open("/admin/eventmanage/common/pop_eventitem_addinfo.asp?acURL=<%=acURL%>&eC="+eC+"&egC="+egC, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		}else{
			popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?acURL=<%=acURL%>&PR=S", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		}
		popwin.focus();
}

//��ü ����
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}


function CkDisPrice(){
	CkDisOrOrg(true);
}

function CkOrgPrice(){
	CkDisOrOrg(false);
}

//���� ���ΰ� ����
function CkDisOrOrg(isDisc){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if(isDisc==true){
					frm.iDSPrice.value = frm.saleprice.value;
					frm.iDBPrice.value = frm.salesupplyprice.value;
					frm.iDSMargin.value= frm.salemargin.value;
					frm.saleItemStatus.value = 7;
				}else{
					frm.iDSPrice.value = frm.orgPrice.value;
					frm.iDBPrice.value = frm.orgSupplyPrice.value;
					frm.iDSMargin.value= frm.orgMarginValue.value;
					frm.saleItemStatus.value = 9;
				}
			}
			reCALbyPrice(frm.itemid.value);
		}
	}
}

//���û�ǰ ����
function saveArr(){
	var frm;
	var pass = false;
	var ovPer = 0;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	frmarr.itemid.value = "";
	frmarr.sailyn.value = "";
	frmarr.iDSPrice.value ="";
	frmarr.iDBPrice.value ="";


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				//check Not AvaliValue
				if (!IsDigit(frm.iDSPrice.value)){
					alert('���ڸ� �����մϴ�.');
					frm.iDSPrice.focus();
					return;
				}

				if (frm.iDSPrice.value<1){
					alert('�ݾ��� ��Ȯ�� �Է��ϼ���.');
					frm.iDSPrice.focus();
					return;
				}

				if (!IsDigit(frm.iDBPrice.value)){
					alert('���ڸ� �����մϴ�.');
					frm.iDBPrice.focus();
					return;
				}

				if (frm.iDBPrice.value<1){
					alert('�ݾ��� ��Ȯ�� �Է��ϼ���.');
					frm.iDBPrice.focus();
					return;
				}

				if(Math.round((frm.orgPrice.value-frm.iDSPrice.value)/frm.orgPrice.value*100)>=50) {
					ovPer++;
				}

				frmarr.itemid.value = frmarr.itemid.value + frm.itemid.value + ","
				//if (frm.sailyn[0].checked){
					//frmarr.sailyn.value = frmarr.sailyn.value + "Y" + ","
				//}else{
					//frmarr.sailyn.value = frmarr.sailyn.value + "N" + ","
				//}
				frmarr.iDSPrice.value = frmarr.iDSPrice.value + frm.iDSPrice.value + ","
				frmarr.iDBPrice.value = frmarr.iDBPrice.value + frm.iDBPrice.value + ","
				frmarr.saleItemStatus.value = frmarr.saleItemStatus.value + frm.saleItemStatus.value+","

			}
		}
	}

	if(ovPer>0) {
		if(!confirm('!!!\n\n\n���� ��ǰ�߿� �������� �ſ� ���� ��ǰ(50%+)�� �ֽ��ϴ�!\n\n�Է��Ͻ� ������ �½��ϱ�?\n\n')) {
			return;
		}
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frmarr.submit();
	}

}

function delArr(){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	frmdel.itemid.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				frmdel.itemid.value = frmdel.itemid.value + frm.itemid.value + ","
			}
		}
	}

	var ret = confirm('�����Ͻðڽ��ϱ�?');

	if (ret){
		frmdel.submit();
	}

}

// ������ ����
function reCALbyPrice(fid) {
	var frm = document["frmBuyPrc_" + fid];
	if(frm.iDSPrice.value>0) {
		frm.iDSMargin.value = Math.round(((frm.iDSPrice.value-frm.iDBPrice.value)/frm.iDSPrice.value)*100);
	} else {
		frm.iDSMargin.value = 0;
	}

	//������ ǥ��
	var iorgPrice = frm.orgPrice.value;
	var isailprice = frm.iDSPrice.value;
	var isalePercent = Math.round((iorgPrice-isailprice)/iorgPrice*100);

	if(isalePercent>=50) {
		document.getElementById("lyrSpct"+fid).style.color="#EE0000";
		document.getElementById("lyrSpct"+fid).style.fontWeight="bold";
	} else {
		document.getElementById("lyrSpct"+fid).style.color="#000000";
		document.getElementById("lyrSpct"+fid).style.fontWeight="normal";
	}
	document.getElementById("lyrSpct"+fid).innerHTML = isalePercent + "%";

}

// ���԰� ����
function reCALbyMargin(fid) {
	var frm = document["frmBuyPrc_" + fid];
	if(frm.iDSMargin.value>0) {
		frm.iDBPrice.value = Math.round(frm.iDSPrice.value*(1-(frm.iDSMargin.value/100)));
	} else {
		frm.iDBPrice.value = frm.iDSPrice.value;
	}
}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" class="a">
<tr> 
	<td width="100%">
		<table  border="0"  width="100%" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">�����ڵ�</td>
			<td bgcolor="#FFFFFF" ><%=sCode%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>"  width="100">���θ�</td>
			<td bgcolor="#FFFFFF"  ><%=sTitle%></td>
		</tr>
		<tr>	
			<td align="center" bgcolor="<%= adminColor("tabletop") %>"   >�̺�Ʈ�ڵ�(�׷�)</td>
			<td bgcolor="#FFFFFF"  ><%If eCode > 0 THEN%><%=eCode%><%If egCode > 0 THEN%>(<%=egCode%>)<%END IF%><%END IF%>&nbsp;</td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>" >����</td>
			<td bgcolor="#FFFFFF" ><%=fnGetCommCodeArrDesc(arrsalestatus,isStatus)%></td>
		</tr>
		<tr>	
			<td align="center" bgcolor="<%= adminColor("tabletop") %>"  >�Ⱓ</td>
			<td bgcolor="#FFFFFF" colspan="3"><%=dSDay%> ~ <%=dEDay%></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<form name="frmSearch" method="get" action="">
			<input type=hidden name=menupos value="<%=menupos%>">
			<input type=hidden name=sC value="<%=sCode%>">
			<input type=hidden name=iC value="<%=iCurrpage%>">
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<tr>
				<td width="100" bgcolor="#EEEEEE" align="center">�˻�����</td>
				<td bgcolor="#ffffff">
					<table   border="0"  cellpadding="3" cellspacing="1" class="a">
					<tr>
						<td width="300"> �귣��: 
					   	<% drawSelectBoxDesignerWithName "makerid",makerid %> 
						</td> 
						<td>��ǰ�ڵ�:</td>
						<td rowspan="2" bgcolor="#FFFFFF"><textarea name="sRectItemidArr" rows="3" cols="10"><%=replace(sRectItemidArr,",",chr(10))%></textarea> </td>  
					</tr> 	
					<tr>
						<td colspan="3"  bgcolor="#FFFFFF">
					    	<input type="checkbox" name="sailyn" value="Y" <% if sailyn="Y" then response.write "checked" %> >�������� ��ǰ �˻�
				            &nbsp;<input type="checkbox" name="invalidmargin" value="Y" <% if invalidmargin="Y" then response.write "checked" %> >��������(or ������) ��ǰ �˻�
				       	</td> 
					</tr> 
				</table>
				</td>
				<td  width="120" bgcolor="#EEEEEE" align="center">
					 <input type="button" class="button" value="��ϵ� ��ǰ �˻�" onclick="document.frmSearch.submit();">
				</td> 
			</tr>
		</table>
		</form>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border=0>
		<form name=frmdummi>
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<tr height="40" valign="bottom">
			<td align="left"><input type=button value="���û�ǰ����" onClick="saveArr()" class="button">
			<!--<input type=button value="���û�ǰ����" onClick="delArr()" class="button">		-->
			</td>
			<td align="right">
			������: <font color="blue"><%=isRate%>%</font>, ��������: <%=fnGetCommCodeArrDesc(arrsalemargin,isMargin)%><%IF isMargin = 5 THEN%>,&nbsp;���θ�����: <font color="blue"><%=isMValue%>%</font> <%END IF%>
			<input type="button" value="��������" onClick="CkDisPrice();" class="button">
			<input type="button" value="��������" onClick="CkOrgPrice();" class="button">
			&nbsp;&nbsp;
			<input type="button" value="����ǰ �߰�" onclick="addnewItem(<%=eCode%>,<%=egCode%>);" class="button">
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2">
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td colspan="17" align="left">�˻���� : <b><%=iTotCnt%></b>&nbsp;&nbsp;������ : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td><input type="checkbox" name="ck_all" onclick="SelectCk(this)"></td>
				<td align="center">��ǰID</td>
				<td align="center" >�̹���</td>
				<td align="center">�귣��</td>
				<td align="center">��ǰ��</td>
				<td align="center">���<br>����</td>
				<td align="center">���λ���</td>
				<td align="center">����<br>�ǸŰ�</td>
				<td align="center">����<br>���԰�</td>
				<td align="center">����<br>������</td>

				<td align="center">��<br>�ǸŰ�</td>
				<td align="center">��<br>���԰�</td>
				<td align="center">��<br>������</td>

				<td align="center">������</td>
				<td align="center">����<br>�ǸŰ�</td>
				<td align="center">����<br>���԰�</td>
				<td align="center">����<br>������</td>
		</tr>
		<%	Dim mSPrice, mSBPrice, iSaleMargin, iOrgMargin, iSalePercent
			Dim cpSP, cpSB, cpSM, strCpDesc, strCpList
			iSaleMargin=0
			iOrgMargin = 0
		IF isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
			mSPrice  =arrList(13,intLoop) - (arrList(13,intLoop)*(isRate/100))
			mSBPrice = fnSetSaleSupplyPrice(isMargin,isMValue,arrList(13,intLoop),arrList(14,intLoop),mSPrice)
			if mSPrice<>0 then iSaleMargin =  100-fix(mSBPrice/mSPrice*10000)/100
			 if arrList(13,intLoop)<>0 then iOrgMargin= 100-fix(arrList(14,intLoop)/arrList(13,intLoop)*10000)/100
			 iSalePercent = ((arrList(13,intLoop)-arrList(2,intLoop))/arrList(13,intLoop))*100

			cpSP=0: cpSB=0: cpSM=0: strCpDesc="": strCpList=""
			if isArray(arrItemCoupon) then

				for icLp=0 to ubound(arrItemCoupon,2)
					if cStr(arrItemCoupon(4,icLp))=cStr(arrList(1,intLoop)) then
						'��ǰ�����ǸŰ�
						Select Case arrItemCoupon(1,icLp)
							Case "1"
								cpSP = mSPrice- CLng(arrItemCoupon(2,icLp)*mSPrice/100)
							Case "2"
								cpSP = mSPrice- arrItemCoupon(2,icLp)
							Case Else
								cpSP = mSPrice
						End Select
						'��ǰ�������԰�
						cpSB = arrItemCoupon(5,icLp)
						'��ǰ��������
						if cpSB>0 then cpSM = formatNumber(100-fix(cpSB/cpSP*10000)/100,0)

						strCpList = strCpList & "<tr align='center' onclick=""window.open('/admin/shopmaster/itemcouponlist.asp?menupos=786&research=on&iSerachType=1&sSearchTxt=" & arrItemCoupon(0,icLp) & "')"">" &_
								"<td>[" & arrItemCoupon(0,icLp) & "]</td>" &_
								"<td>" & arrItemCoupon(3,icLp) & "</td>" &_
								"<td>" & FormatNumber(cpSP,0) & "��</td>" &_
								"<td>" & FormatNumber(cpSB,0) & "��</td>" &_
								"<td " & chkIIF(cpSM<=5,"style='color:#ee0000;'","") & ">" & FormatNumber(cpSM,0) & "%</td>" &_
								"<td>" & left(arrItemCoupon(6,icLp),10) & "</td>" &_
								"<td>" & left(arrItemCoupon(7,icLp),10) & "</td>" &_
								"</tr>"
					end if
				next

				if strCpList<>"" then
					strCpDesc = "<div><font color=darkgreen style='cursor:pointer;' onmouseover=""$(this).find('div').show()"" onmouseout=""$(this).find('div').hide()"">��ǰ���� ��" &_
							"<div style='display:none;position:absolute;border:1px solid #C0C0C0;padding:5px;background-color:#FFFFFF;margin:-10px -20px;'>" &_
							"<table width='600' border='0' cellpadding='3' cellspacing='1' class='a'>" &_
							"<tr><td colspan='7' align='left'><strong>���αⰣ�� ����Ǵ� ����</strong></td></tr>" &_
							"<tr align='center' bgcolor='#F0F0F0'>" &_
							"<td colspan='2'>������</td>" &_
							"<td>�������ΰ�</td>" &_
							"<td>�������԰�</td>" &_
							"<td>�������θ���</td>" &_
							"<td>������</td>" &_
							"<td>������</td>" &_
							"</tr>" &_
							strCpList &_
							"</table>" &_
							"</div></font></div>"
				end if

			end if
			%>
			<form name="frmBuyPrc_<%=arrList(1,intLoop)%>" >
			<input type=hidden name="itemid" value="<%=arrList(1,intLoop)%>">
			<input type=hidden name="saleprice" value="<%=mSPrice%>">
			<input type=hidden name="salesupplyprice" value="<%=mSBPrice%>">
			<input type=hidden name="salemargin" value="<%=iSaleMargin%>">
			<input type=hidden name="orgPrice" value="<%=arrList(13,intLoop)%>">
			<input type=hidden name="orgSupplyPrice" value="<%=arrList(14,intLoop)%>">
			<input type=hidden name="orgMarginValue" value="<%=iOrgMargin%>">
			<input type=hidden name="saleItemStatus" value="<%=arrList(4,intLoop)%>">
		 <tr align="center" bgcolor=<%IF cint(arrList(4,intLoop)) = 8 THEN%>"#B3B3B3"<%ELSE%>"#FFFFFF"<%END IF%>>
			    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
			    <td><%=arrList(1,intLoop)%></td>
			    <td><%IF arrList(9,intLoop) <> "" THEN%><img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(arrList(1,intLoop))%>/<%=arrList(9,intLoop)%>"><%END IF%></td>
			    <td><%=db2html(arrList(7,intLoop))%></td>
			    <td align="left">&nbsp;<%=db2html(arrList(8,intLoop))%></td>
			    <td><%= fnColor(arrList(17,intLoop),"mw") %></td>
			    <td>
			    	<%= fnColor(arrList(10,intLoop),"yn") %>&nbsp;<%IF arrList(4,intLoop) = 6 THEN%><font color="blue"><%END IF%><%=fnGetCommCodeArrDesc(arrsalestatus,arrList(4,intLoop))%>
			    	<%=chkIIF(strCpDesc>"",strCpDesc,"")%>
			    </td>

			    <td><%=formatnumber(arrList(11,intLoop),0)%></td>
			    <td><%=formatnumber(arrList(12,intLoop),0)%></td>
			    <td><% if arrList(11,intLoop)<>0 then %>
					<%= 100-fix(arrList(12,intLoop)/arrList(11,intLoop)*10000)/100 %>%
					<% end if %>  
				</td>


			    <td><%=formatnumber(arrList(13,intLoop),0)%></td>
			    <td><%=formatnumber(arrList(14,intLoop),0)%></td>
			    <td><%=iOrgMargin%>%</td>

				<td id="lyrSpct<%=arrList(1,intLoop)%>" style="<%=chkIIF(iSalePercent>=50,"color:#EE0000;font-weight:bold;","")%>"><%=formatnumber(iSalePercent,0)%>%</td>
			<%IF cint(arrList(4,intLoop)) = 8 or  cint(arrList(4,intLoop)) = 9 THEN%>
				<td><input type="text" name="iDSPrice" size="6" maxlength="9" value="0" style="text-align:right;" onkeyup="reCALbyPrice('<%=arrList(1,intLoop)%>')"><br><%=arrList(2,intLoop)%></td>
			    <td><input type="text" name="iDBPrice" size="6" maxlength="9" value="0" style="text-align:right;" onkeyup="reCALbyPrice('<%=arrList(1,intLoop)%>')"><br><%=arrList(3,intLoop)%></td>
			    <td><input type="text" name="iDSMargin" value="0" style="text-align:right;" size="4" onkeyup="reCALbyMargin('<%=arrList(1,intLoop)%>')">%<br><%  if arrList(2,intLoop)<>0 then smargin= 100-fix(arrList(3,intLoop)/arrList(2,intLoop)*10000)/100 	%></td>
			<%ELSE%>
			    <td><input type="text" name="iDSPrice" size="6" maxlength="9" value="<%=arrList(2,intLoop)%>" style="text-align:right;" onkeyup="reCALbyPrice('<%=arrList(1,intLoop)%>')"></td>
			    <td><input type="text" name="iDBPrice" size="6" maxlength="9" value="<%=arrList(3,intLoop)%>" style="text-align:right;" onkeyup="reCALbyPrice('<%=arrList(1,intLoop)%>')"></td>
			    <td><%  if arrList(2,intLoop)<>0 then smargin= 100-fix(arrList(3,intLoop)/arrList(2,intLoop)*10000)/100 	%> 
					<input type="text" name="iDSMargin" value="<%=smargin%>" style=text-align:right;" size="4" onkeyup="reCALbyMargin('<%=arrList(1,intLoop)%>')">%
			    </td>
			<%END IF%>
		</tr>
		</form>
		<%	next %>
		<% else %>
		<tr>
			<td colspan="17" bgcolor="#ffffff" align="center">��ϵ� ������ �����ϴ�.</td>
		</tr>
		<%
		END IF%>
		<!-- ����¡ó�� -->
			<%
			iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1

			If (iCurrpage mod iPerCnt) = 0 Then
				iEndPage = iCurrpage
			Else
				iEndPage = iStartPage + (iPerCnt-1)
			End If
			%>
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			    <tr valign="bottom" height="25">
			        <td valign="bottom" align="center">
			         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
					<% else %>[pre]<% end if %>
			        <%
						for ix = iStartPage  to iEndPage
							if (ix > iTotalPage) then Exit for
							if Cint(ix) = Cint(iCurrpage) then
					%>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
					<%		else %>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
					<%
							end if
						next
					%>
			    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
					<% else %>[next]<% end if %>
			        </td>
			        <td  width="50" align="right"><a href="saleList.asp?menupos=<%=menupos%>"><img src="/images/icon_list.gif" border="0"></a></td>
			    </tr>
		</table>
	</td>
</tr>
</table>
<form name="frmarr" method="post" action="saleItemPRoc.asp">
<input type="hidden" name="mode" value="U">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="sC" value="<%=sCode%>">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="sailyn" value="">
<input type="hidden" name="iDSPrice" value="">
<input type="hidden" name="iDBPrice" value="">
<input type="hidden" name="saleItemStatus" value="">
<input type="hidden" name="saleStatus" value="<%=isStatus%>">
</form>
<form name="frmdel" method="post" action="saleItemPRoc.asp">
<input type="hidden" name="mode" value="D">
<input type="hidden" name="sC" value="<%=sCode%>">
<input type="hidden" name="itemid" value="">
</form>
<%
set clsSaleItem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->