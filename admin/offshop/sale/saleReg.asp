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
Dim sMode ,eCode, cEvent ,sTitle, dSDay, dEDay, sBrand,eState , sale_shopmargin
Dim sCode, clsSale,isRate, isMargin,sale_shopmarginvalue, isStatus, egCode, isUsing, dOpenDay,isMValue,dCloseDay
Dim intgroup , strParm , shopid , point_rate, shopname
Dim iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,sStatus
	eCode     = requestCheckVar(Request("eC"),10)
	sCode     = requestCheckVar(Request("sC"),10)
	isRate = 0
	isUsing = true
	sMode  = "I"
	isStatus =0

IF sCode <> "" THEN
	set clsSale = new CSale
	sMode = "U"
	clsSale.FSCode  = sCode
	clsSale.fnGetSaleConts

	sTitle 		= clsSale.FSName
	isRate 		= clsSale.FSRate
	point_rate = clsSale.fpoint_rate
	isMargin 	= clsSale.FSMargin
	eCode 		= clsSale.FECode
	egCode		= clsSale.FEGroupCode
	dSDay 		= clsSale.FSDate
	dEDay 		= clsSale.FEDate
	isStatus 	= clsSale.FSStatus
	isUsing     = clsSale.FSUsing
	dOpenDay	= clsSale.FOpenDate
	isMValue	= clsSale.FSMarginValue
	sale_shopmargin = clsSale.fsale_shopmargin
	sale_shopmarginvalue	= clsSale.fsale_shopmarginvalue
	dCloseDay 	= clsSale.FCloseDate
	shopid = clsSale.Fshopid
	shopname = getoffshopname(shopid)

	'-�˻�----------------------------------------
	 iSerachType    = requestCheckVar(Request("selType"),4)		'�˻�����
	 sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'�˻���
	 sDate     		= requestCheckVar(Request("selDate"),1)		'�˻��� ����
	 sSdate     	= requestCheckVar(Request("iSD"),10)		'������
	 sEdate     	= requestCheckVar(Request("iED"),10)		'������
	 sStatus		= requestCheckVar(Request("salestatus"),4)	' ����
	 iCurrpage		= requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ

	 strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&sstatus="&sStatus
	'---------------------------------------------
	set clsSale = nothing
END IF

IF eCode = "0" THEN eCode = ""

IF eCode <> "" THEN		'�̺�Ʈ ���� �ϰ��
	IF sCode = "" THEN
		set cEvent = new cevent_list
			cEvent.Frectevt_code = eCode
			cEvent.fnGetEventConts

			sTitle 	= cEvent.foneitem.fevt_name
			dSDay	= cEvent.foneitem.fevt_startdate
			dEDay	= cEvent.foneitem.fevt_enddate
			isStatus  = cEvent.foneitem.fevt_state
			dOpenDay = cEvent.foneitem.FOpenDate
			shopid = cEvent.foneitem.Fshopid
		set cEvent = nothing
	END IF
END IF

IF dSDay ="" THEN dSDay = date()
IF isStatus < 6 THEN isStatus = 0
if point_rate = "" then point_rate = "0"

'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
Dim  arrsalestatus
	arrsalestatus = fnSetCommonCodeArr_off("salestatus",False)
%>

<script language="javascript">

	//-- jsPopCal : �޷� �˾� --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	function jsSubmitSale(){
		var frm = document.frmReg;

		if(!frm.sSN.value){
			alert("������ �Է��� �ּ���");
			return false;
		}

		if(!frm.sSD.value ){
		  	alert("�������� �Է����ּ���");
		  	frm.sSD.focus();
		  	return false;
	  	}

		if(!frm.sED.value ){
		  	alert("�������� �Է����ּ���");
		  	frm.sED.focus();
		  	return false;
	  	}

	  	if(frm.sED.value){
		  	if(frm.sSD.value > frm.sED.value){
			  	alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
			  	frm.sED.focus();
			  	return false;
		  	}
		}

		if(!frm.shopid.value){
			alert("������ �������ּ���");
			return false;
		}

		if(frm.shopid.value.substring(0,3) == "ith"){
			alert("��ϺҰ� �����Դϴ�.");
			return false;
		}

		if(typeof(frm.chkstatus)=="object"){
			if(frm.chkstatus.checked) {
				frm.salestatus.value = frm.chkstatus.value;
			}
		}

		var nowDate = "<%=date()%>";
	   if(frm.salestatus.value==7){
	 	if(frm.sOD.value !=""){
	 		nowDate = '<%IF dOpenDay <> ""THEN%><%=FormatDate(dOpenDay,"0000-00-00")%><%END IF%>';
		}

		if(frm.sSD.value < nowDate){
			alert("�������� �����Ϻ���  ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ���");
		  	frm.sSD.focus();
		  	return false;
		 }
	  }

	  	if(!frm.iSR.value){
			alert("�������� �Է��� �ּ���");
			frm.iSR.focus();
			return false;
		}
		if (!IsDouble(frm.iSR.value)){
			alert('�������� ���ڸ� �����մϴ�.');
			frm.iSR.focus();
			return false;
		}

		if(confirm('�����Ͻðڽ��ϱ�?')){
			return true;
		}else{
			return false;
		}
	}

	function jsChSetValue(iVal,itype){

		if (itype == 'salemargin'){
			if(iVal ==5){
				document.all.divM.style.display = "";
			}else{
				document.all.divM.style.display = "none";
			}
		}else if (itype == 'shopsalemargin'){
			if(iVal ==5){
				document.all.divsM.style.display = "";
			}else{
				document.all.divsM.style.display = "none";
			}
		}
	}

</script>

�� ��������<br>
���ϸ���: �ǸŰ� ��� ���� ������ ����<br>
��ü�δ�: ���ǸŰ��� �����ݾ׸�ŭ �����ǸŰ����� ����<br>
�ݹݺδ�: ���αݾ��� 1/2�ݾ��� �����ް����� ����<br>
�ٹ����ٺδ�: �����ް��� �����ǸŰ��ް��� ����<br>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1"  >
<form name="frmReg" method="post" action="saleProc.asp?<%=strParm%>" onSubmit="return jsSubmitSale();">
<input type="hidden" name="sM" value="<%=sMode%>">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="sC" value="<%=sCode%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="sSU" value="1">
<tr>
	<td>
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<%IF eCode <> "" THEN%>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">�̺�Ʈ�ڵ�(�׷�)</td>
			<td bgcolor="#FFFFFF" colspan="3"><%=eCode%>
			 <input type="hidden" name="selG" value="0">
			</td>
		</tr>
		<%END IF%>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> ����</td>
			<td bgcolor="#FFFFFF"><%IF eCode <> "" THEN %><%=sTitle%><input type="hidden" name="sSN" value="<%=sTitle%>"><%ELSE%><input type="text" name="sSN" size="30" maxlength="64" value="<%=sTitle%>"><%END IF%></td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center"> �Ⱓ</td>
			<td bgcolor="#FFFFFF">
				������ :
				<%IF eCode <> "" THEN %>
					<input type="hidden" name="sSD" value="<%=dSDay%>">
					<%=dSDay%> ~
				<%ELSE%>
					<input type="text" name="sSD" size="10" onClick="jsPopCal('sSD');" style="cursor:hand;" value="<%=dSDay%>"> ~
				<%END IF%>
				������ :
				<%IF eCode <> "" THEN %>
					<input type="hidden" name="sED" value="<%=dEDay%>">
					<%=dEDay%>
				<%ELSE%>
					<input type="text" name="sED"  size="10" onClick="jsPopCal('sED');" style="cursor:hand;" value="<%=dEDay%>">
				<%END IF%>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"> ������</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iSR" size="3" maxlength="3" value="<%=isRate%>" style="text-align:right;">%</td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center">����</td>
			<td bgcolor="#FFFFFF">
				<% if eCode <> "" then %>
					- �̺�Ʈ�� ���� ������ ��� �̺�Ʈ�� ��ϴ�� �ش� ����� ���ο� ��ϴ�� �ش��� ������ ���ƾ� �մϴ�<br>
				<% end if %>
				<% if sCode <> "" then %>
					<%= shopname %>(<%= shopid %>)
					<input type="hidden" name="shopid" value="<%= shopid %>">
				<% else %>
					<% drawSelectBoxOffShopdiv_off "shopid",shopid , "1,3,11" ,"","" %>
				<% end if %>

			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"> ����</td>
			<td bgcolor="#FFFFFF" >
				<input type="hidden" name="sOD" value="<%=dOpenDay%>">
				<input type="hidden" name="salestatus" value="<%=isStatus%>">
				<%=fnGetCommCodeArrDesc_off(arrsalestatus,isStatus)%>
				<%if eCode = "" then%>
					<%IF isStatus =0 then '��ϴ�� %>
						<input type="checkbox" name="chkstatus" value="7">���¿�û
						<Br>�� ���μ����� ������ �ݵ�� <font color="red">���¿�û</font>�� üũ �ϼž�, ������ �ڵ� ����ó�� �˴ϴ�.
						<Br><font color="red">�ٷ� ����</font>�ÿ��� ���¿�û�� üũ �ϼž�, ����Ʈ�� <font color="red">�ǽð�����</font> ��ư�� Ȱ��ȭ �˴ϴ�.
					<%elseif isStatus = 6 or isStatus = 7 then '���� %>
						<input type="checkbox" name="chkstatus" value="9">�����û
						<Br>�� ���»����ε� <font color="red">��¥�� ����</font>��� �����û üũ�� ���� �ʾƵ�, ������ �ڵ� ���� �˴ϴ�.
						<br><font color="red">���� ����</font>�ÿ��� �����û�� üũ�ϼž�, ����Ʈ�� <font color="red">�ǽð�����</font> ��ư�� Ȱ��ȭ �˴ϴ�.
					<%elseif isStatus = 8 then %>
						<div style="padding-top:5px;">������: <%=dCloseDay%></div>
					<%end if%>
				<%end if%>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" align="center">���Ը���</td>
			<td bgcolor="#FFFFFF" >
				<%sbGetOptCommonCodeArr_off "salemargin", isMargin, False,True," onchange='jsChSetValue(this.value,""salemargin"");'"%>
				<span id="divM" style="display:<%IF isMargin<> 5 THEN %>none<%END IF%>;">���θ���<input type="text" size="4" name="isMV" maxlength="10" value="<%=isMValue%>" style="text-align:right;">%</span>
				<br><br>
				��å�� <font color="red">����</font>��ǰ�� <font color="red">�ٹ����ٺδ�</font> ���� ����ϼž� �մϴ�.(���԰� ���� �Ұ�)
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">�����޸���</td>
			<td bgcolor="#FFFFFF">
				<%sbGetOptCommonCodeArr_off "shopsalemargin", sale_shopmargin, False,True," onchange='jsChSetValue(this.value,""shopsalemargin"");'"%>
				<span id="divsM" style="display:<%IF isMargin<> 5 THEN %>none<%END IF%>;">���θ���<input type="text" size="4" name="sale_shopmarginvalue" maxlength="10" value="<%=sale_shopmarginvalue%>" style="text-align:right;">%</span>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����Ʈ����</td>
			<td bgcolor="#FFFFFF">
				<!--�������
				<input type="text" size="3" name="point_rate" maxlength="3" value="<%'=point_rate%>" style="text-align:right;" readonly>%-->
				<input type="text" size="3" name="point_rate" maxlength="3" value="<%=point_rate%>" style="text-align:right;">%
				<Br>��å�� ������ ����Ʈ�������� 0% �Դϴ�. ����Ʈ ������ ���Ͻø� �Է��ϼ���.
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<input type="image" src="/images/icon_save.gif">
		<a href="javascript:history.back();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
