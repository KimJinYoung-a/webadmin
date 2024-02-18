<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  ���� ����
' History : 2010.12.01 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/sale/sale_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->

<%
Dim iPageSize, iCurrpage ,iDelCnt , iTotCnt ,clsSale, arrList, intLoop , eCode , shopid
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt , strParm
Dim iSerachType,sSearchTxt,sBrand,  sDate,sSdate,sEdate,isStatus
	eCode     		= requestCheckVar(Request("eC"),10)			'�̺�Ʈ �ڵ�
	iSerachType    = requestCheckVar(Request("selType"),4)		'�˻�����
	sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'�˻���
	sBrand     	= requestCheckVar(Request("ebrand"),32)		'�귣��
	sDate     		= requestCheckVar(Request("selDate"),1)		'�˻��� ����
	sSdate     	= requestCheckVar(Request("iSD"),10)		'������
	sEdate     	= requestCheckVar(Request("iED"),10)		'������
	isStatus		= requestCheckVar(Request("salestatus"),4)	'���� ����
	arrList = ""
	shopid		= requestCheckVar(Request("shopid"),32)		'����
	iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ

	'�˻��κ��� ��ȣ�� �޾ƾߵȴٸ� ���ڸ� ����
 	if iSerachType="1" or iSerachType="2" then
 		sSearchTxt = getNumeric(sSearchTxt)
 	end if

	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 50
	iPerCnt = 10

	IF Cstr(eCode) = "0" THEN eCode = ""
	IF (eCode <> "" AND sSearchTxt = "") THEN
		iSerachType = 2
		sSearchTxt = eCode
	END IF

    strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&ebrand="&sBrand&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&sstatus="&isStatus
	set clsSale = new CSale
		clsSale.FECode = eCode
		clsSale.FSearchType = iSerachType
 		clsSale.FSearchTxt  = sSearchTxt
 		clsSale.FBrand		= sBrand
 		clsSale.FDateType   = sDate
 		clsSale.FSDate		= sSdate
 		clsSale.FEDate		= sEdate
 		clsSale.FSStatus	= isStatus
	 	clsSale.FCPage 		= iCurrpage
	 	clsSale.FPSize 		= iPageSize
	 	clsSale.frectshopid = 	shopid
		arrList = clsSale.fnGetSaleList	'�����͸�� ��������

 		iTotCnt = clsSale.FTotCnt	'��ü ������  ��
 	set clsSale = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��

	Dim arrsalemargin, arrsalestatus , arrsaleshopmargin
	'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
	arrsalemargin = fnSetCommonCodeArr_off("salemargin",False)
	arrsaleshopmargin = fnSetCommonCodeArr_off("shopsalemargin",False)
	arrsalestatus= fnSetCommonCodeArr_off("salestatus",False)
%>

<script language="javascript">

	//�޷�
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//����
	function jsMod(scode){
		location.href = "saleReg.asp?sC="+scode+"&menupos=<%=menupos%>&<%=strParm%>";
	}

	//����¡ó��
		function jsGoPage(iP){
		document.frmSearch.iC.value = iP;
		document.frmSearch.submit();
	}

	//�̵�
	function jsGoURL(type,ival){
		if(type=="e"){
			location.href = "/admin/offshop/event_off/event_modify.asp?evt_code="+ival;
		}else if(type=="i"){
			location.href = "saleItemReg.asp?sC="+ival+"&menupos=<%=menupos%>";
		}
	}

	//���� �ٷ� ����ó��
 	function jsSetRealSale(sCode, chkState){
 		if(chkState !=1){
 			alert("�������̰� ���糯¥�� ���� �Ⱓ���϶��� �ǽð� ó�� �����մϴ�.");
 			return;
 		}

 		if(confirm("��ϵ� ����ǰ�� ���� ����� �������� �Ǽ����� ����Ǹ�,\n\n�������� ��� ������� �Ͻǰ�� �ٷ� �ݿ� �˴ϴ�.\n\nó���Ͻðڽ��ϱ�?")){
 			document.frmReal.sC.value = sCode;
 			document.frmReal.submit();
 		}
 	}

	//���� ��ü �ǽð� ����
 	function jsSetRealSaleall(){
 		if(confirm("[�����ڸ��] ���� ��ü �ǽð� ����\nó���Ͻðڽ��ϱ�?")){
			var pop_realall = window.open('/admin/offshop/sale/saleproc.asp?menupos=<%=menupos%>&sM=realall','pop_realall','width=600,height=400,scrollbars=yes,resizable=yes');
			pop_realall.focus();
 		}
 	}

	//���� ���� ����
	function copyshop(upfrm, onlySameMargin) {
		if (upfrm.copyshopid.value == ''){
			alert('���� ���� ������ �������ּ���');
			return;
		}

		if(confirm("�����Ͻ� ���γ����� �����ϰ� ���� ������忡 ���� ���� ������ ���� ���� �Ͻðڽ��ϱ�?") == true) {
			upfrm.sC.value = '';

			if (!CheckSelected()){
					alert('���þ������� �����ϴ�.\n������ ��� ������ ������ �ּ���.');
					return;
				}
				var frm;
				var tmp = 0;
					for (var i=0;i<document.forms.length;i++){
						frm = document.forms[i];
						if (frm.name.substr(0,9)=="frmBuyPrc") {
							if (frm.cksel.checked){
								upfrm.sC.value = upfrm.sC.value + frm.sale_code.value;
								tmp = tmp + 1;
							}
						}
					}

				if (tmp != '1'){
					alert('���� ������ �Ѱ����� ���� �ϽǼ� �ֽ��ϴ�');
					return;
				}

			if (onlySameMargin != undefined) {
				upfrm.sOnlySameMargin.value = onlySameMargin;
			}
			upfrm.sM.value = 'copyshop';
			upfrm.action='saleProc.asp';
			upfrm.submit();
		}
	}

</script>

<!---- �˻� ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmReal" method="post" action="saleItemProc.asp?<%=strParm%>">
<input type="hidden" name="sC">
<input type="hidden" name="mode" value="P">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<form name="frmSearch" method="get"  action="saleList.asp" onSubmit="return jsSearch(this,'E');">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="iC">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan=2>�˻�<br>����</td>
	<td align="left">
		<select name="selType">
			<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>�����ڵ�</option>
			<option value="2" <%IF Cstr(iSerachType) = "2" THEN%>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
			<option value="3" <%IF Cstr(iSerachType) = "3" THEN%>selected<%END IF%>>���θ�</option>
		</select>
		<input type="text" name="sTxt" value="<%=sSearchTxt%>" size="30" maxlength="30">
		&nbsp;&nbsp;
		* �Ⱓ:
		<select name="selDate">
		<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>������ ����</option>
		<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>������ ����</option>
		</select>
		<input type="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:hand;">
		~ <input type="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:hand;">
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan=2>
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frmSearch.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ����:
		<% sbGetOptCommonCodeArr_off "salestatus", isStatus, True, False,"onChange='javascript:document.frmSearch.submit();'"%>
		&nbsp;&nbsp;
		* ���� : <% drawSelectBoxOffShopdiv_off "shopid",shopid , "1,3,11" ,"","" %>
	</td>
</tr>

</form>
</table>
<!---- /�˻� ---->

<Br>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
<form name="copyfrm" method="post">
<input type="hidden" name="sC">
<input type="hidden" name="sM">
<input type="hidden" name="sOnlySameMargin">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr height="40" valign="bottom">
    <td align="left">
		<font color="red">[�ʵ�]</font> �Ϸ翡 �ѹ� ����5�ÿ� ���°� ���¿�û�� ��� �������� �ڵ�����Ǹ�, ���»����ε� ��¥�� ������� �ڵ� ���� �˴ϴ�
		<Br>���忡�� �����̳�, ����� <font color="red">��ùݿ�</font>�� ���Ͻô°��, <font color="red">�ݵ�� �ǽð�����</font> ��ư�� ��������.
		<!--<br>&nbsp;&nbsp;���� �������� <font color="red">�������� ��ǰ</font>�̶��, ���ο� ��ǰ ����� �Ұ��� �մϴ�.-->
    </td>
    <td align="right">
    	* ����������� : <% drawSelectBoxOffShopdiv_off "copyshopid","" , "1,3" ,"","" %>
    	<input type="button" value="�����ڵ� ������ ��������" class="button" onclick="copyshop(copyfrm);">
		&nbsp;&nbsp;
		<input type="button" value="��ǰ�� ������ ��������(���ϸ��� ��ǰ��)" class="button" onclick="copyshop(copyfrm, 'Y');">
    	&nbsp;&nbsp;
    	<input type="button" value="�űԵ��" class="button" onclick="javascript:location.href='saleReg.asp?menupos=<%=menupos%>&eC=<%=eCode%>';" >
    </td>
</tr>
</form>

</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="20">
		�˻���� : <b><%=iTotCnt%></b>&nbsp;&nbsp;������ : <b><%=iCurrpage%> / <%=iTotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>����<br>�ڵ�</td>
	<!--<td>�̺�Ʈ�ڵ�</br>(�׷��ڵ�)</td>-->
	<td>���θ�</td>
	<td>����</td>
	<td>���Ը���<br>������޸���</td>
	<td>������<br>������</td>
	<td>��ǰ��������ð�</td>
	<td>����</td>
	<td>������</td>
	<td>����<br>����Ʈ</td>
	<td>
		���
    	<% if C_ADMIN_AUTH then %>
    		<input type="button" value="��ü�ǽð�����" class="button" onclick="jsSetRealSaleall();">
    	<% end if %>
	</td>
</tr>
<% Dim chkState
IF isArray(arrList) THEN
	For intLoop = 0 To UBound(arrList,2)
	chkState = 0
	'����: ����, �����û )�Ⱓ: �����ϱ��� �Ⱓ��
	if (arrList(8,intLoop) = 6 or arrList(8,intLoop) = 7 or arrList(8,intLoop) = 9) and datediff("d",arrList(6,intLoop),date()) >=0 and datediff("d",arrList(7,intLoop),date()) <=0 then
		chkState = 1
	end if
%>
<form action="" name="frmBuyPrc<%=intLoop%>" method="get">
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
	<td align="center" width=25>
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
	</td>
	<td width=70>
		<%=arrList(0,intLoop)%><input type="hidden" name="sale_code" value="<%= arrList(0,intLoop) %>">
	</td>
	<!--<td>-->
		<%' IF arrList(4,intLoop) > 0 THEN%>
			<!--<a href="javascript:jsGoURL('e',<%=arrList(4,intLoop)%>)" title="�̺�Ʈ ��������">-->
			<%'= arrList(4,intLoop) %></a>
			<%' IF arrList(5,intLoop) > 0 THEN %>
				<!--<br>(<%'= arrList(5,intLoop) %>)-->
			<%' END IF %>
		<%' END IF %>
	<!--</td>-->
	<td align="left">
		<%=db2html(arrList(1,intLoop))%>
	</td>
	<td width=140>
		<%=arrList(20,intLoop)%><Br><%=arrList(17,intLoop)%>
	</td>
	<td>
		<%=fnGetCommCodeArrDesc_off(arrsalemargin,arrList(3,intLoop))%>
		<br><%=fnGetCommCodeArrDesc_off(arrsaleshopmargin,arrList(18,intLoop))%>
	</td>
	<td width=80>
		<%=arrList(6,intLoop)%><br><%=arrList(7,intLoop)%>
	</td>
	<td width=170>
		<% if arrList(15,intLoop) <> "" or not isnull(arrList(15,intLoop)) then %>
			����:<%=arrList(15,intLoop)%>
		<% end if %>
		<% if arrList(16,intLoop) <> "" or not isnull(arrList(16,intLoop)) then %>
			<br>����:<%=arrList(16,intLoop)%>
		<% end if %>
	</td>
	<td width=60>
		<%
		'/����
		IF arrList(8,intLoop) = 6 THEN
		%>
			<font color="blue"><%=fnGetCommCodeArrDesc_off(arrsalestatus,arrList(8,intLoop))%></font>
		<%
		'/����
		elseIF arrList(8,intLoop) = 8 THEN
		%>
			<font color="gray"><%=fnGetCommCodeArrDesc_off(arrsalestatus,arrList(8,intLoop))%></font>
		<%
		'/���¿�û , �����û
		elseIF arrList(8,intLoop) = 7 or arrList(8,intLoop) = 9 THEN
		%>
			<font color="red"><%=fnGetCommCodeArrDesc_off(arrsalestatus,arrList(8,intLoop))%></font>
		<% else %>
			<%=fnGetCommCodeArrDesc_off(arrsalestatus,arrList(8,intLoop))%>
		<% end if %>
	</td>
	<td width=50>
		<%=arrList(2,intLoop)%> %
	</td>
	<td width=50>
		<%=arrList(19,intLoop)%> %
	</td>
	<td width=280>
		<input type="button" value="����" onclick="jsMod(<%=arrList(0,intLoop)%>);" class="button">
		<input type="button" value="��ǰ(<%=arrList(13,intLoop)%>)" class="button" onClick="javascript:jsGoURL('i',<%=arrList(0,intLoop)%>)">
		<%IF chkState = 1 THEN%>
			<input type="button" value="�ǽð�����" class="button" onClick="jsSetRealSale(<%=arrList(0,intLoop)%>,<%=chkState%>);">
		<%END IF%>
	</td>
</tr>
</form>
<% Next %>
<%
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1

If (iCurrpage mod iPerCnt) = 0 Then
	iEndPage = iCurrpage
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
<tr align="center" bgcolor="#FFFFFF" >
    <td valign="bottom" align="center" colspan=20>
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
</tr>
<% ELSE %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">��ϵ� ������ �����ϴ�.</td>
</tr>
<%END IF%>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
