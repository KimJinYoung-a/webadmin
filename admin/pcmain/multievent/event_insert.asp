<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : event_insert.asp
' Discription : pcmain �̺�Ʈ����
' History : 2014.06.23 ����ȭ
'         : 2018-11-26 ������ �Ǿ����� �������� �߰�
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/pcmain_multieventCls.asp" -->
<%
'###############################################
'�̺�Ʈ �ű� ��Ͻ�
'###############################################
%>
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
Dim eCode
Dim idx , evtimg , isusing , mode
Dim srcSDT , srcEDT
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim linkurl
Dim maincopy
Dim event_info
Dim prevDate , ordertext
Dim startdate
Dim enddate
dim tag_only
dim dispOption
dim contentType

Dim cEvtCont
Dim ename , stdt , eddt , sortnum , todaybanner , Molistbanner , subcopy , event_info_option , subname
Dim sale_per , coupon_per, contentImg, tmpEventCode
dim isSale, isGift, isCoupon, isCommnet, isOnlyTen, isOneplusOne, isFreedelivery, isNew, saleCPer, salePer

	contentType = request("contentType")
	dispOption = request("dispOption")
	tmpEventCode = requestCheckvar(Request("eC"),10)
	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")

If idx = "" Then
	mode = "add"
Else
	mode = "modify"
End If

if contentType = "" then
	contentType = 1
end if

if contentType = "" then
	contentType = "1"
end if

'// ������
If idx <> "" then
	dim oEnjoyeventOne
	set oEnjoyeventOne = new CMainbanner
	oEnjoyeventOne.FRectIdx = idx
	oEnjoyeventOne.GetOneContents()

	linkurl				=	oEnjoyeventOne.FOneItem.Flinkurl
	maincopy			=	oEnjoyeventOne.FOneItem.Fmaincopy
	startdate			=	left(oEnjoyeventOne.FOneItem.Fevtstdate, 10) 
	enddate				=	left(oEnjoyeventOne.FOneItem.Fevteddate, 10)
	mainStartDate		=	oEnjoyeventOne.FOneItem.Fstartdate
	mainEndDate			=	oEnjoyeventOne.FOneItem.Fenddate
	isusing				=	oEnjoyeventOne.FOneItem.Fisusing
	ordertext			=	oEnjoyeventOne.FOneItem.Fordertext
	sortnum				=	oEnjoyeventOne.FOneItem.Fsortnum
	eCode				=	oEnjoyeventOne.FOneItem.Feventid
	Molistbanner		=	oEnjoyeventOne.FOneItem.Fevtmolistbanner
	subcopy				=	oEnjoyeventOne.FOneItem.Fsubcopy
	sale_per			=	oEnjoyeventOne.FOneItem.Fsale_per
	coupon_per			=	oEnjoyeventOne.FOneItem.Fcoupon_per
	tag_only			=	oEnjoyeventOne.FOneItem.Ftag_only
	dispOption			=   oEnjoyeventOne.FOneItem.FdispOption
	contentImg			=   oEnjoyeventOne.FOneItem.FcontentImg
	contentType			=   oEnjoyeventOne.FOneItem.FcontentType
	event_info  		= 	oEnjoyeventOne.FOneItem.FEventInfo
	event_info_option  		= 	oEnjoyeventOne.FOneItem.FEventInfoOption
	'�߰�
	isSale					= oEnjoyeventOne.FOneItem.FESale
	isGift					= oEnjoyeventOne.FOneItem.FEGift		
	isCoupon				= oEnjoyeventOne.FOneItem.FECoupon			
	isCommnet				= oEnjoyeventOne.FOneItem.FECommnet			
	isOnlyTen				= oEnjoyeventOne.FOneItem.FSisOnlyTen			
	isOneplusOne		= oEnjoyeventOne.FOneItem.FEOneplusOne					
	isFreedelivery	= oEnjoyeventOne.FOneItem.FEFreedelivery						
	isNew						= oEnjoyeventOne.FOneItem.FENew		

	if event_info_option = "1" then '����
		saleCPer = oEnjoyeventOne.FOneItem.FECsalePer
		salePer = event_info
	elseif event_info_option = "2" then '����
		saleCPer = event_info
		salePer = oEnjoyeventOne.FOneItem.FESalePer
	else
		saleCPer = oEnjoyeventOne.FOneItem.FECsalePer
		salePer = oEnjoyeventOne.FOneItem.FESalePer	
	end if
	
	set oEnjoyeventOne = Nothing
End If

'// �Է½�
IF tmpEventCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = tmpEventCode	'�̺�Ʈ �ڵ�
	eCode = tmpEventCode
	'�̺�Ʈ ���� ��������
	cEvtCont.fnGetEventCont
	maincopy	=	db2html(cEvtCont.FEName)
	subcopy	=	db2html(cEvtCont.FENamesub)
	stdt	=	left(cEvtCont.FESDay, 10)
	eddt	=	left(cEvtCont.FEEDay, 10)
	'�̺�Ʈ ȭ�鼳�� ���� ��������
	cEvtCont.fnGetEventDisplay
	todaybanner = cEvtCont.FEBImgMoToday
	Molistbanner = cEvtCont.FEBImgMoListBanner
	'�߰�
	isSale = cEvtCont.FESale
	isGift = cEvtCont.FEGift
	isCoupon = cEvtCont.FECoupon
	isCommnet = cEvtCont.FECommnet
	isOnlyTen = cEvtCont.FSisOnlyTen
	isOneplusOne = cEvtCont.FEOneplusOne
	isFreedelivery = cEvtCont.FEFreedelivery
	isNew = cEvtCont.FENew	
	saleCPer = cEvtCont.FECsalePer
	salePer = cEvtCont.FESalePer	
	event_info = salePer

	startdate = stdt
	enddate = eddt
	dim tmpename
		tmpename = Split(maincopy,"|")

	if Ubound(tmpename)>0 then
		maincopy = tmpename(0)
	end if

	set cEvtCont = nothing
END IF

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
		if prevDate = "" then
			prevDate = sDt
		end if 
	elseif prevDate <> "" then
		sDt = prevDate
	else
		sDt = date
	end if
	sTm = "00:00:00"
end if

if Not(mainEndDate="" or isNull(mainEndDate)) then
	eDt = left(mainEndDate,10)
	eTm = Num2Str(hour(mainEndDate),2,"0","R") &":"& Num2Str(minute(mainEndDate),2,"0","R") &":"& Num2Str(second(mainEndDate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
		if prevDate = "" then 
			prevDate = eDt
		end if 
	elseif prevDate <> "" then
		eDt = prevDate
	else
		eDt = date
	end if
	eTm = "23:59:59"
end If
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;

		if (frm.linkurl.value.indexOf("�̺�Ʈ��ȣ") > 0 || frm.linkurl.value.indexOf("��ǰ�ڵ�") > 0){
			alert("��ũ ���� Ȯ�� ���ּ���.");
			frm.linkurl.focus();
			return;
		}

		if (confirm('���� �Ͻðڽ��ϱ�?')){
			//frm.target = "blank";
			frm.action = "event_proc.asp";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/pcmain/multievent/";
	}

	$(function(){
	//�޷´�ȭâ ����
	var arrDayMin = ["��","��","ȭ","��","��","��","��"];
	var arrMonth = ["1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��"];
    $("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
    	showCurrentAtPos: 1,
      	showOn: "button",
      	<% if Idx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#eDt" ).datepicker( "option", "minDate", selectedDate );
    	}
    });
    $("#eDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 2,
      	showOn: "button",
      	<% if Idx<>"" then %>minDate: "<%=sDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });
		controlDisp('<%=dispOption%>');
			
		$('input[name=event_info_option]').click(function(){	
			var salePer = "<%=salePer%>";
			var saleCper = "<%=saleCPer%>";

			var tmpValueTxt = $(this).attr("valueTxt");
			var valueTxt = tmpValueTxt

			if(valueTxt == "����"){
				valueTxt = salePer;
			}else if(valueTxt == "����"){
				valueTxt = saleCper;
			}		
			$("#event_info").val(valueTxt)
		})

});

//-- jsPopCal : �޷� �˾� --//
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function putLinkText(key,gubun) {
	var frm = document.frm;
	var urllink
	if (gubun == "1" )
	{
		urllink = frm.linkurl;
	}
	switch(key) {
		case 'event':
			urllink.value='/event/eventmain.asp?eventid=�̺�Ʈ��ȣ';
			break;
		case 'itemid':
			urllink.value='/category/category_itemprd.asp?itemid=��ǰ�ڵ�';
			break;
	}
}
//���� �̺�Ʈ �ҷ�����
function jsLastEvent(){
  var valsdt , valedt , valgcode, vDispOption
	valsdt = document.frm.StartDate.value;
	valedt = document.frm.EndDate.value;
	vDispOption = document.frm.dispOption.value;

// alert(vDispOption);
// return false;
  var winLast,eKind;
  winLast = window.open('pop_event_list.asp?menupos=<%=menupos%>&sDt='+valsdt+'&eDt='+valedt+'&dispOption='+vDispOption+'&idx=<%=idx%>','pLast','width=550,height=600, scrollbars=yes')
  winLast.focus();
}
function changeForm(){
  	var valsdt , valedt
	valsdt = document.frm.sDt.value;
	valedt = document.frm.eDt.value;

	var contentType = document.frm.contentType.value;	
	var dispOption = document.frm.dispOption.value;	
	var link = contentType == 1 ? "event_insert.asp" : "item_insert.asp"
	document.location.href= link + "?menupos=<%=menupos%>&prevDate=<%=sDt%>&dispOption="+dispOption+"&contentType="+contentType;	
}
function controlDisp(optionVal){	
	var row = document.getElementById("contentType");
	if(optionVal==2){
		row.style.display = "";
		$("#prdInfo").css("display","none")	
		$("#normalTag").css("display","none")			
		$("#eventInfo").css("display","")  				
		$("#evtInfo").css("display","")					
	}else{
		row.style.display = "none";
		$("#prdInfo").css("display","")	
		$("#normalTag").css("display","")	
		$("#eventInfo").css("display","none")  						
		$("#evtInfo").css("display","none")			
	}
}
</script>
<form name="frm" method="post" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="eventid" value="<%=eCode%>">
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="10%">���� ��ġ</td>
    <td >
		<select name="dispOption" class="select" onchange="controlDisp(this.value)">				
			<option value="1" <%=chkiif(dispOption="1"," selected","")%>>�⺻</option>
			<option value="2" <%=chkiif(dispOption="2"," selected","")%>>���λ�ܱ�ȹ��</option>
		</select>					
    </td>
</tr>
<tr bgcolor="#FFFFFF" id="contentType" style="display:<%=chkIIF(dispOption = "1" or dispOption = "" , "none", "")%>;">
    <td bgcolor="#FFF999" align="center" width="10%">����������</td>
	<td>
		<div style="float:left;">
			<input type="radio" onclick="changeForm()" name="contentType" value="1" <%=chkiif(contentType = "1","checked","")%> checked />�̺�Ʈ &nbsp;&nbsp;&nbsp; 
			<input type="radio" onclick="changeForm()" name="contentType" value="2" <%=chkiif(contentType = "2","checked","")%>/>��ǰ
		</div>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="10%">����Ⱓ</td>
    <td >
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="30">�̺�Ʈ�ҷ�����</td>
	<td >
		<% If Molistbanner <> "" Then %>
		<img src="<%=Molistbanner%>" width="200"><br/><%=Molistbanner%>
		<% End If %>
		<input type="button" value="�̺�Ʈ �ҷ�����" onclick="jsLastEvent();"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="30">�̺�Ʈ URL</td>
	<td >
		<% IF eCode <> "" THEN %>
			<input type="text" name="linkurl" size="80" value="/event/eventmain.asp?eventid=<%=eCode%>">
		<% Else %>
			<input type="text" name="linkurl" size="80" value="<%=linkurl%>"/>
		<% End If %>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('event','1')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','1')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�̺�Ʈ ����</td>
	<td>
		<input type="text" name="maincopy" value="<%=maincopy%>" size="40" maxlength="200"/>
		</br>
		<input type="text" name="subcopy" id="subcopy" value="<%=subcopy%>" size="70" maxlength="400"/>
		<font color="red"></font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�̺�Ʈ ������ - ������</td>
	<td >
		<input type="text" name="evtstdate" size="10" value="<%=chkiif(mode="add",stdt,startdate)%>" onClick="jsPopCal('evtstdate');"/>
		-
		<input type="text" name="evteddate" size="10" value="<%=chkiif(mode="add",eddt,enddate)%>" onClick="jsPopCal('evteddate');"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="prdInfo">
  <td bgcolor="#FFF999" align="center">����/����</td>
  <td>
	<input type="text" name="sale_per" value="<%=sale_per%>"> : ������ ex)<font color="red"><strong>~45%</strong></font></br>
	<input type="text" name="coupon_per" value="<%=coupon_per%>"> : ���������� ex)<font color="green"><strong>10%</strong></font></br>
	<font color="red"><strong>���ִ� ��츸 �Է� �ϼ���.��</strong></font>
  </td>
</tr>
<tr bgcolor="#FFFFFF" id="normalTag">
	<td bgcolor="#FFF999" align="center">�±�</td>
	<td >
		<div style="float:left;">
			<input type="checkbox" name="tag_only" value="Y" <%=chkiif(tag_only = "Y","checked","")%>/> �ܵ� 
		</div> <br/>
		<div style="float:right;margin-top:5px;margin-right:10px;"></div>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="eventInfo">
	<td bgcolor="#FFF999" align="center">�̺�Ʈ���� ����</td>
	<td colspan="3">
		<div style="float:left;">		
		<input type="radio" name="event_info_option" value="1" valueTxt="����" <%=chkiif(event_info_option = "1","checked","")%> checked/>���� &nbsp;&nbsp;&nbsp;  
		<% if isCoupon then %> 			<input type="radio" name="event_info_option" value="2" valueTxt="����" <%=chkiif(event_info_option = "2","checked","")%> />���� &nbsp;&nbsp;&nbsp; <% end if %>
		<% if isGift then %>   			<input type="radio" name="event_info_option" value="3" valueTxt="GIFT" <%=chkiif(event_info_option = "3","checked","")%> />GIFT &nbsp;&nbsp;&nbsp; <% end if %>
		<% if isOneplusOne then %>  <input type="radio" name="event_info_option" value="4" valueTxt="1+1"  <%=chkiif(event_info_option = "4","checked","")%> />1+1 &nbsp;&nbsp;&nbsp; <% end if %>
		<% if isNew then %> 				<input type="radio" name="event_info_option" value="5" valueTxt="��Ī" <%=chkiif(event_info_option = "5","checked","")%> />��Ī &nbsp;&nbsp;&nbsp; <% end if %>
		<% if isCommnet then %> 		<input type="radio" name="event_info_option" value="6" valueTxt="����" <%=chkiif(event_info_option = "6","checked","")%> />���� &nbsp;&nbsp;&nbsp; <% end if %>
		<% if isOnlyTen then %> 		<input type="radio" name="event_info_option" value="7" valueTxt="�ܵ�" <%=chkiif(event_info_option = "7","checked","")%> />�ܵ� &nbsp;&nbsp;&nbsp; <% end if %>
		</div> 
		<font color="red"><strong>�� �Ѱ��� ���� ���� ��</strong></font>				
	</td>		
</tr>
<tr bgcolor="#FFFFFF" id="evtInfo">
  <td bgcolor="#FFF999" align="center">�̺�Ʈ����</td>
  <td colspan="3">
	<input type="text" id="event_info" name="event_info" value="<%=event_info%>"> : ������ ex)<font color="red"><strong>~45%</strong></font></br>	
	<font color="red"><strong>���ִ� ��츸 �Է� �ϼ���.��</strong></font>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td ><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">���� ��ȣ</td>
	<td ><input type="text" name="sortnum" size="10" value="<%=chkiif(mode="add","99",sortnum)%>" maxlength="3"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�۾��� ���û���</td>
	<td ><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" �� �� " onClick="jsgolist();"/><input type="button" value=" �� �� " onClick="jsSubmit();"/></td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->