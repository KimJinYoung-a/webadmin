<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� ��ǰ ����
' History : 2010.09.28 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/sale/salecls.asp"-->
<%
Dim sCode, clsSale,clsSaleItem ,acURL ,iTotCnt, arrList,intLoop ,iPageSize, iCurrpage ,iDelCnt
Dim sTitle,isRate, isMargin, isStatus,eCode, egCode, dSDay, dEDay, isUsing, dOpenDay,isMValue, smargin
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
	sCode = requestCheckVar(Request("sC"),10)
	acURL =Server.HTMLEncode("/academy/sale/saleitemProc.asp?sC="&sCode)

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
		Case 4	'�ΰŽ��δ�
			fnSetSaleSupplyPrice = orgSupplyPrice
		Case 5	'��������
			fnSetSaleSupplyPrice = salePrice - fix(salePrice*(MarginValue/100))
	END SELECT	
End Function
'-----------------------------------------------------------------------------------
If sCode<> "" Then
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

	iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
	IF iCurrpage = "" THEN	iCurrpage = 1			
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	
'���� ��ǰ����	
set clsSaleItem = new CSaleItem
	clsSaleItem.FCPage = iCurrpage
	clsSaleItem.FPSize = iPageSize	
	clsSaleItem.FSCode = sCode	
	arrList = clsSaleItem.fnGetSaleItemList
	iTotCnt = clsSaleItem.FTotCnt	'��ü ������  ��
	 
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ �� 

'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
Dim arrsalemargin, arrsalestatus
arrsalemargin = fnSetCommonCodeArr("salemargin",False)
arrsalestatus= fnSetCommonCodeArr("salestatus",False)	
%>

<script language="javascript">

// ������ �̵�
function jsGoPage(iP){
	location.href="saleItemReg.asp?menupos=<%=menupos%>&sC=<%=sCode%>&iC="+iP;		
}

// ����ǰ �߰� �˾�
function addnewItem(eC,egC){
	var popwin;
	if ( eC > 0 ){
		popwin = window.open("/academy/event/common/pop_eventitem_addinfo.asp?acURL=<%=acURL%>&eC="+eC+"&egC="+egC, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
	}else{
		popwin = window.open("/academy/itemmaster/pop_itemAddInfo.asp?acURL=<%=acURL%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
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
		}
	}
}

//���û�ǰ ����
function saveArr(){
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
}

// ���԰� ����
function reCALbyMargin(fid) {
	var frm = document["frmBuyPrc_" + fid];
	if(frm.iDSMargin.value>0) {
		alert(frm.originprice.value + "/" + frm.iDSPrice.value );
		frm.iDBPrice.value = Math.round(frm.iDSPrice.value*(1-(frm.iDSMargin.value/100)));
		frm.iDSSaleRate.value = Math.round(((frm.originprice.value-frm.iDSPrice.value)/frm.originprice.value)*100);
	} else {
		frm.iDBPrice.value = frm.iDSPrice.value;
		frm.iDSSaleRate.value = Math.round(((frm.originprice.value-frm.iDSPrice.value)/frm.originprice.value)*100);
	}
}

</script>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" class="a">
<tr>
	<td width="100%">
		<table  border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="80">�����ڵ�</td>
			<td bgcolor="#FFFFFF" width="60"><%=sCode%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>"  width="80">���θ�</td>
			<td bgcolor="#FFFFFF"  width="150"><%=sTitle%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">�̺�Ʈ�ڵ�(�׷�)</td>
			<td bgcolor="#FFFFFF"  width="80"><%If eCode > 0 THEN%><%=eCode%><%If egCode > 0 THEN%>(<%=egCode%>)<%END IF%><%END IF%>&nbsp;</td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="80">����</td>
			<td bgcolor="#FFFFFF"  width="60"><%=fnGetCommCodeArrDesc(arrsalestatus,isStatus)%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="80">�Ⱓ</td>
			<td bgcolor="#FFFFFF" width="200"><%=dSDay%> ~ <%=dEDay%></td>
		</tr>		
		</table>	
	</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border=0>
		<form name=frmdummi>
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<tr height="40" valign="bottom">
			<td align="left"><input type=button value="���û�ǰ����" onClick="saveArr()" class="button">
			<input type=button value="���û�ǰ����" onClick="delArr()" class="button">
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
			<td colspan="16" align="left">�˻���� : <b><%=iTotCnt%></b>&nbsp;&nbsp;������ : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td><input type="checkbox" name="ck_all" onclick="SelectCk(this)"></td>    				    				    	
				<td align="center">��ǰID</td>
				<td align="center" >�̹���</td>				
				<td align="center">�귣��</td>
				<td align="center">��ǰ��</td>
				<td align="center">���<br>����</td>
				<td align="center">����<br>�����ǸŰ�</td>
				<td align="center">������԰�<br>������԰�</td>
				<td align="center">��������<br>���縶����</td>
			                      

			                      
				<td align="center">����<br>�ǸŰ�</td>
				<td align="center">����<br>���԰�</td>
				<td align="center">������</td>
				<td align="center">����<br>������</td>				    			
		</tr>	
		<% Dim mSPrice, mSBPrice, iSaleMargin, iOrgMargin
			iSaleMargin=0
			iOrgMargin = 0
		IF isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)	
			mSPrice  =arrList(13,intLoop) - (arrList(13,intLoop)*(isRate/100))	
			mSBPrice = fnSetSaleSupplyPrice(isMargin,isMValue,arrList(13,intLoop),arrList(14,intLoop),mSPrice)	
			if mSPrice<>0 then iSaleMargin =  100-fix(mSBPrice/mSPrice*10000)/100
			 if arrList(13,intLoop)<>0 then iOrgMargin= 100-fix(arrList(14,intLoop)/arrList(13,intLoop)*10000)/100
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
			<input type=hidden name="originprice" value="<%=arrList(13,intLoop)%>">
		 <tr align="center" bgcolor=<%IF cint(arrList(4,intLoop)) = 8 THEN%>"#B3B3B3"<%ELSE%>"#FFFFFF"<%END IF%>>    
			    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>    				    	
			    <td><%=arrList(1,intLoop)%></td>		    				    	
			    <td><%IF arrList(9,intLoop) <> "" THEN%><img src="<%=imgFingers%>/diyItem/webimage/small/<%=GetImageSubFolderByItemid(arrList(1,intLoop))%>/<%=arrList(9,intLoop)%>"><%END IF%></td>	
			    <td><%=db2html(arrList(7,intLoop))%></td>
			    <td align="left">&nbsp;<%=db2html(arrList(8,intLoop))%></td>			    
			    <td><%= fnColor(arrList(17,intLoop),"mw") %></td>
			    
			    <td><%=formatnumber(arrList(13,intLoop),0)%><% If arrList(10,intLoop)="Y" Then %><br><font color=#F08050><%=formatnumber(arrList(11,intLoop),0)%></font><% End If %></td>
			    <td><%=formatnumber(arrList(14,intLoop),0)%><% If arrList(10,intLoop)="Y" Then %><br><font color=#F08050><%=formatnumber(arrList(12,intLoop),0)%></font><% End If %></td>
			    <td><% if arrList(11,intLoop)<>0 then %>
					<%= 100-fix(arrList(14,intLoop)/arrList(13,intLoop)*10000)/100 %>%<% If arrList(10,intLoop)="Y" Then %><br><font color=#F08050><%= 100-fix(arrList(12,intLoop)/arrList(11,intLoop)*10000)/100 %>%</font><% End If %>
					<% end if %>					
				</td>

			
			<%IF cint(arrList(4,intLoop)) = 8 or  cint(arrList(4,intLoop)) = 9 THEN%>
				<td><input type="text" name="iDSPrice" size="6" maxlength="9" value="0" style="text-align:right;" onkeyup="reCALbyMargin('<%=arrList(1,intLoop)%>')"></td>
			    <td><input type="text" name="iDBPrice" size="6" maxlength="9" value="0" style="text-align:right;"></td>
			    <td><input type="text" name="iDSSaleRate" value="0" style="text-align:right;" size="4">%</td>
				<td><input type="text" name="iDSMargin" value="0" style="text-align:right;" size="4">%</td>
			<%ELSE%>
			    <td><input type="text" name="iDSPrice" size="6" maxlength="9" value="<%=arrList(2,intLoop)%>" style="text-align:right;" onkeyup="reCALbyMargin('<%=arrList(1,intLoop)%>')"></td>
			    <td><input type="text" name="iDBPrice" size="6" maxlength="9" value="<%=arrList(3,intLoop)%>" style="text-align:right;"></td>
			    <td><input type="text" name="iDSSaleRate" value="<%=round((arrList(13,intLoop)-arrList(2,intLoop))/arrList(13,intLoop)*100,1)%>" style="text-align:right;" size="4">%</td>
				<td><%  if arrList(2,intLoop)<>0 then smargin= 100-fix(arrList(3,intLoop)/arrList(2,intLoop)*10000)/100 	%>				
					<input type="text" name="iDSMargin" value="<%=smargin%>" style="text-align:right;" size="4">%
			    </td>
			<%END IF%>
		</tr>	
		</form>    
		<%	next
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

<form name=frmarr method=post action="saleItemPRoc.asp">
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
<form name=frmdel method=post action="saleItemPRoc.asp">
	<input type="hidden" name="mode" value="D">
	<input type="hidden" name="sC" value="<%=sCode%>">
	<input type="hidden" name="itemid" value="">
</form>
<%END IF%>
<%
set clsSaleItem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->