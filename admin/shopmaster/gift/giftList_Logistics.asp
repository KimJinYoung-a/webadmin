<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����ǰ ����
' History : 2008.04.01 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/itemgiftcls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
'Call fnSetEventCommonCode '�����ڵ� ���ø����̼� ������ ����

 Dim eCode
 Dim clsGift, arrList, intLoop
 Dim iTotCnt
 Dim iPageSize, iCurrpage ,iDelCnt
 Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
 Dim iSerachType,sSearchTxt,sBrand,  sDate,sSdate,sEdate,igStatus,sgDelivery
 Dim strParm

 eCode     		= requestCheckVar(Request("eC"),10)			'�̺�Ʈ �ڵ�
 iSerachType    = requestCheckVar(Request("selType"),4)		'�˻�����
 sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'�˻���
 sBrand     	= requestCheckVar(Request("ebrand"),32)		'�귣��
 sDate     		= requestCheckVar(Request("selDate"),1)		'�˻��� ����
 sSdate     	= requestCheckVar(Request("iSD"),10)		'������
 sEdate     	= requestCheckVar(Request("iED"),10)		'������
 igStatus		= requestCheckVar(Request("giftstatus"),4)	'����ǰ ����
 sgDelivery		= requestCheckVar(Request("selDelivery"),1)	'�������

 if igStatus="" then igStatus="6" end if
 if sgDelivery="" then sgDelivery="N" end if

 iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ

	IF iCurrpage = "" THEN	iCurrpage = 1
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����

	IF Cstr(eCode) = "0" THEN eCode = ""

	IF (eCode <> "" AND sSearchTxt = "") THEN
		iSerachType = "2"
		sSearchTxt = eCode
	ELSEIF 	(iSerachType="2" AND sSearchTxt <> "") THEN
		eCode = sSearchTxt
	END IF

    strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&ebrand="&sBrand&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&giftstatus="&igStatus
	set clsGift = new CGift
		clsGift.FECode = eCode
		clsGift.FSearchType = iSerachType
 		clsGift.FSearchTxt  = sSearchTxt
 		clsGift.FBrand		= sBrand
 		clsGift.FDateType   = sDate
 		clsGift.FSDate		= sSdate
 		clsGift.FEDate		= sEdate
 		clsGift.FGStatus	= igStatus
 		clsGift.FGDelivery	= sgDelivery

	 	clsGift.FCPage 		= iCurrpage
	 	clsGift.FPSize 		= iPageSize

		arrList = clsGift.fnGetGiftList	'�����͸�� ��������
 		iTotCnt = clsGift.FTotCnt	'��ü ������  ��
 	set clsGift = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��

	'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
	Dim  arrgiftscope, arrgifttype,arrgiftstatus
	arrgiftscope 	= fnSetCommonCodeArr("giftscope",False)
	arrgifttype 	= fnSetCommonCodeArr("gifttype",False)
	arrgiftstatus 	= fnSetCommonCodeArr("giftstatus",False)

%>
<script language="JavaScript" src="/js/ttpbarcode.js"></script>
<script language="javascript">
<!--
	//�޷�
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//����
	function jsMod(gcode){
		location.href = "giftMod.asp?gC="+gcode+"&menupos=<%=menupos%>&<%=strParm%>";
	}

	//����¡ó��
		function jsGoPage(iP){
		document.frmSearch.iC.value = iP;
		document.frmSearch.submit();
	}

	//�̵�
	function jsGoURL(type,ival){
		if(type=="e"){
			location.href = "/admin/eventmanage/event/event_modify.asp?eC="+ival;
		}
	}

	//��ǰ������ �������̵�
	function jsItem(giftscope,gCode, eCode){
		//�̺�Ʈ��ϻ�ǰ, ���û�ǰ�ϋ� ��ǰ view, �׿� �������̵�
		if(giftscope == 2 || giftscope == 4 ){
			location.href = "/admin/eventmanage/event/eventitem_regist.asp?eC="+eCode+"&menupos=870";
		}else if(giftscope==5){
			location.href = "giftItemReg.asp?gC="+gCode+"&menupos=<%=menupos%>&<%=strParm%>";
		}
	}

//-->

function DrawReceiptPrintobj_TEC(elementid,printname){
        var objstring = "";
        var e;
        objstring = '<OBJECT name="' + elementid + '" ';
        objstring = objstring + ' classid="clsid:E76C9051-A8C4-458E-9F60-3C14DB9EECF9" ';
        objstring = objstring + ' codebase="http://billyman/Tec_dol.cab#version=1,5,0,0" ';
        objstring = objstring + ' width=0 ';
        objstring = objstring + ' height=0 ';
        objstring = objstring + ' align=center ';
        objstring = objstring + ' hspace=0 ';
        objstring = objstring + ' vspace=0 ';
        objstring = objstring + ' > ';
        objstring = objstring + ' <PARAM Name="PrinterName" Value="' + printname + '"> ';
        objstring = objstring + ' </OBJECT>';

        document.write(objstring);
}

/*
function eventindexprint(ievt_code, ievt_name01, ievt_name02, ievt_startdate, ievt_enddate, ievt_gift_code, ievt_gift_kind, ievt_gift01, ievt_gift02){
	var X = 1;
	var Y = 1;
	var F = 1;

	// TEC_DO3 : 452
	if (TEC_DO3.IsDriver == 1){
           X = 1.05;
           Y = 1.05;
           F = 1.2;

			TEC_DO3.SetPaper(900,600);
			TEC_DO3.OffsetX = 20;
			TEC_DO3.OffsetY = 20;
			TEC_DO3.PrinterOpen();



			TEC_DO3.PrintText(500*X, 30*Y, "Arial Bold", 100*F, 0, 0, ievt_code);

			TEC_DO3.PrintText(50*X, 50*Y, "HY�߰��", 30*F, 0, 0, "[������]");
			TEC_DO3.PrintText(250*X, 50*Y, "HY�߰��", 30*F, 1, 0, ievt_startdate);

			TEC_DO3.PrintText(50*X, 100*Y, "HY�߰��", 30*F, 0, 0, "[������]");
			TEC_DO3.PrintText(250*X, 100*Y, "HY�߰��", 30*F, 1, 0, ievt_enddate);

			TEC_DO3.PrintText(50*X, 150*Y, "HY�߰��", 30*F, 0, 0, "[�̺�Ʈ��]");
			TEC_DO3.PrintText(250*X, 150*Y, "HY�߰��", 30*F, 1, 0, ievt_name01);
			TEC_DO3.PrintText(250*X, 200*Y, "HY�߰��", 30*F, 1, 0, ievt_name02);

			TEC_DO3.PrintText(50*X, 250*Y, "HY�߰��", 30*F, 0, 0, "----------------------------------------");



			TEC_DO3.PrintText(50*X, 300*Y, "HY�߰��", 30*F, 0, 0, "[����ǰ]");

			TEC_DO3.PrintText(50*X, 330*Y, "Arial Bold", 100*F, 1, 0, ievt_gift_code);
			TEC_DO3.PrintText(300*X, 350*Y, "HY�߰��", 30*F, 1, 0, ievt_gift_kind);	<!-- ���ٿ� �ѱ� 24�ڱ��� ������ �Ʒ��� -->
			TEC_DO3.PrintText(300*X, 400*Y, "HY�߰��", 30*F, 1, 0, ievt_gift01);
			TEC_DO3.PrintText(300*X, 450*Y, "HY�߰��", 30*F, 1, 0, ievt_gift02);

			TEC_DO3.PrinterClose();


    }else window.status = "TEC B-452 ����̹��� ��ġ�� �ּ���"
}

DrawReceiptPrintobj_TEC("TEC_DO3","TEC B-452");
*/

function eventIndexBarcodePrint(eventCode, eventName01, eventName02, eventStartdate, eventEnddate, eventGiftCode, eventGiftKind, eventGift01, eventGift02) {
	// /js/barcode.js ����
	if (initTTPprinter("TTP-243_80x50", "T", "N", "                         www.10x10.co.kr                         ", "Y", "��", "Y", 3, 0) != true) {
		alert('�����Ͱ� ��ġ���� �ʾҰų� �ùٸ� �����͸��� �ƴմϴ�.[4]');
		return;
	}

	printTTPOneIndexBarcodeForEventItem(eventCode, eventName01, eventName02, eventStartdate, eventEnddate, eventGiftCode, eventGiftKind, eventGift01, eventGift02, 1);
}
</script>
<!---- �˻� ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmSearch" method="get"  action="giftList_Logistics.asp" onSubmit="return jsSearch(this,'E');">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="iC">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<select class="select" name="selType">
				<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>����ǰ�ڵ�</option>
				<option value="2" <%IF Cstr(iSerachType) = "2" THEN%>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
			</select>
			<input type="text" class="text" name="sTxt" value="<%=sSearchTxt%>" size="10" maxlength="10">
			&nbsp;
			�귣��:<% drawSelectBoxDesignerwithName "ebrand", sBrand %>
			<br>
			�Ⱓ:
			<select class="select" name="selDate">
				<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>������ ����</option>
				<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>������ ����</option>
			</select>
			<input type="text" class="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:hand;">
			~ <input type="text" class="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:hand;">
			&nbsp;
			����:<%sbGetOptCommonCodeArr "giftstatus", igStatus, True,False,"onChange='javascript:document.frmSearch.submit();'"%>
			&nbsp;
			���:
			<select class="select" name="selDelivery" onChange="javascript:document.frmSearch.submit();">
				<option value="">��ü</option>
				<option value="Y" <%IF sgDelivery="Y" THEN%>selected<%END IF%>>��ü</option>
				<option value="N" <%IF sgDelivery="N" THEN%>selected<%END IF%>>�ٹ�����</option>
			</select>
		</td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frmSearch.submit();">
		</td>
	</tr>
</table>
<!---- /�˻� ---->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="17">�˻���� : <b><%=iTotCnt%></b>&nbsp;&nbsp;������ : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50">����ǰ�ڵ�</td>
    	<td width="50">�̺�Ʈ<br>�ڵ�</br>(�׷�)</td>
    	<td>�̺�Ʈ��</td>
        <td>��ǰ�ڵ�</td>
    	<td>�귣��</td>
    	<td>�������</td>
    	<td>��������</td>
    	<td>�̻�</td>
    	<td>�̸�</td>
    	<td>����</td>
    	<td>����</td>
    	<td>������</td>
    	<td>������</td>
    	<td>����</td>
    	<td>����</td>
    	<td>���</td>
    	<td width="40">�ε���<br>���</td>
    </tr>
    <%IF isArray(arrList) THEN
    	For intLoop = 0 To UBound(arrList,2)
    %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%=arrList(0,intLoop)%></a></td>
    	<td nowrap><%IF arrList(3,intLoop) > 0 THEN%><a href="javascript:jsGoURL('e',<%=arrList(3,intLoop)%>)" title="�̺�Ʈ ��������"><%=arrList(3,intLoop)%></a><%IF arrList(4,intLoop) > 0 THEN%><br>(<%=arrList(4,intLoop)%>)<%END IF%><%END IF%></td>
    	<td align="left"><%=db2html(arrList(1,intLoop))%></td>
        <td align="center">
            <% if Not IsNull(arrList(27,intLoop)) and Not IsNull(arrList(28,intLoop)) and Not IsNull(arrList(29,intLoop)) then %>
            <%= BF_MakeTenBarcode(arrList(27,intLoop), arrList(28,intLoop), arrList(29,intLoop)) %>
            <% end if %>
        </td>
    	<td><%=db2html(arrList(5,intLoop))%></td>
    	<td> <%IF (arrList(2,intLoop) = 2 or arrList(2,intLoop) = 4 or arrList(2,intLoop) = 5) then %>
    		<a href="javascript:jsItem(<%=arrList(2,intLoop)%>,<%=arrList(0,intLoop)%>,<%=arrList(3,intLoop)%>)" title="��ϻ�ǰ ����"><%=fnGetCommCodeArrDesc(arrgiftscope,arrList(2,intLoop))%><br>(<%=arrList(20,intLoop)%>)</a>
    		<%else%>
    		<%=fnGetCommCodeArrDesc(arrgiftscope,arrList(2,intLoop))%>
    		<%end if%>
    		</td>
    	<td><%=fnGetCommCodeArrDesc(arrgifttype,arrList(6,intLoop))%></td>
    	<td nowrap><%=formatnumber(arrList(7,intLoop),0)%></td>
    	<td nowrap><%=formatnumber(arrList(8,intLoop),0)%></td>
    	<td nowrap><%=arrList(11,intLoop)%></td>
    	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%IF arrList(9,intLoop) > 0 THEN%>[<%=arrList(9,intLoop)%>]<%=arrList(19,intLoop)%><%END IF%></a></td>
    	<td nowrap><%=arrList(13,intLoop)%></td>
    	<td nowrap><%=arrList(14,intLoop)%></td>
    	<td nowrap><%=fnGetCommCodeArrDesc(arrgiftstatus,arrList(15,intLoop))%></td>
    	<td nowrap><%IF arrList(12,intLoop) > 0 THEN%><%=arrList(12,intLoop)%><%END IF%></td>
    	<td nowrap><%IF arrList(21,intLoop)="Y" THEN%><font color="#F08050">��ü</font><%ELSE%><font color="#5080F0">�ٹ�����</font><%END IF%></td>
    	<td>
    	    <!--
    	    <input type="button" value="���" class="button" onClick="eventindexprint('<%=arrList(3,intLoop)%>', '<%= left(db2html(arrList(1,intLoop)),20) %>','<%= mid(db2html(arrList(1,intLoop)),21) %>','<%=arrList(13,intLoop)%>','<%=arrList(14,intLoop)%>','<%=arrList(0,intLoop)%>','[<%=arrList(9,intLoop)%>]','<%= left(arrList(19,intLoop),20) %>','<%= mid(arrList(19,intLoop),21) %>')">
    	    -->
    	    <!-- eventindexprint('0', '�����ι���_RECYCLE LETTER','ING ver.3 ���� �� ��Ʈ����','2014-02-19','2014-03-16','13688','[17466]','CLASS NOTE ver.8 ����(','���󷣴�)') -->
    	    <input type="button" class="button" value="���" onClick="eventIndexBarcodePrint('<%=arrList(3,intLoop)%>', '<%= left(db2html(arrList(1,intLoop)),23) %>','<%= mid(db2html(arrList(1,intLoop)),24) %>','<%=arrList(13,intLoop)%>','<%=arrList(14,intLoop)%>','<%=arrList(0,intLoop)%>','[<%=arrList(9,intLoop)%>]','<%= left(arrList(19,intLoop),23) %>','<%= mid(arrList(19,intLoop),24) %>')">

    	</td>

    </tr>
	<% Next
	ELSE
	%>
	<tr>
		<td colspan="17" align="center" bgcolor="#FFFFFF">��ϵ� ������ �����ϴ�.</td>
	</tr>
	<%END IF%>
</table>
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
    </tr>
    </form>
</table>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
