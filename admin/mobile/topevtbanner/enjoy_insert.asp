<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : enjoy_insert.asp
' Discription : ����� ī�װ� TOP 2 EVENT
' History : 2015-09-16 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/topeventCls.asp" -->
<!-- #include virtual="/lib/classes/mobile/topcatecodeCls.asp" -->
<%
'###############################################
'�̺�Ʈ �ű� ��Ͻ�
'###############################################
%>
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
Dim eCode
Dim idx , evtimg , subImage2 , subImage3 , subImage4 , isusing , mode , gcode
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate
Dim sDt, sTm, eDt, eTm
Dim evtalt, linkurl 
Dim evttitle
Dim issalecoupontxt
Dim prevDate , ordertext
Dim startdate
Dim enddate
Dim issalecoupon , linktype

Dim cEvtCont
Dim ename , stdt , eddt , sortnum , todaybanner , Molistbanner , evttitle2

	eCode = requestCheckvar(Request("eC"),10)
	gcode = requestCheckvar(Request("gcode"),3)
	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")
	linktype = request("linktype") '�̺�Ʈ��ũŸ��

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

'// �Է½�
IF eCode <> "" And mode = "add" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
	'�̺�Ʈ ���� ��������
	cEvtCont.fnGetEventCont
	ename	=	db2html(cEvtCont.FEName)
	stdt	=	cEvtCont.FESDay
	eddt	=	cEvtCont.FEEDay
	'�̺�Ʈ ȭ�鼳�� ���� ��������
	cEvtCont.fnGetEventDisplay
	todaybanner = cEvtCont.FEBImgMoToday
	Molistbanner = cEvtCont.FEBImgMoListBanner
	
	set cEvtCont = nothing
END IF

'// ������
If idx <> "" then
	dim oTopevtOne
	set oTopevtOne = new CMainbanner
	oTopevtOne.FRectIdx = idx
	oTopevtOne.GetOneContents()

	linktype			=	oTopevtOne.FOneItem.Flinktype
	evtalt				=	oTopevtOne.FOneItem.Fevtalt
	linkurl				=	oTopevtOne.FOneItem.Flinkurl
	evtimg				=	oTopevtOne.FOneItem.Fevtimg
	evttitle			=	oTopevtOne.FOneItem.Fevttitle
	issalecoupontxt		=	oTopevtOne.FOneItem.Fissalecoupontxt
	startdate			=	oTopevtOne.FOneItem.Fevtstdate
	enddate				=	oTopevtOne.FOneItem.Fevteddate
	issalecoupon		=	oTopevtOne.FOneItem.Fissalecoupon
	mainStartDate		=	oTopevtOne.FOneItem.Fstartdate
	mainEndDate			=	oTopevtOne.FOneItem.Fenddate 
	isusing				=	oTopevtOne.FOneItem.Fisusing
	ordertext			=	oTopevtOne.FOneItem.Fordertext
	sortnum				=	oTopevtOne.FOneItem.Fsortnum
	todaybanner			=	oTopevtOne.FOneItem.Ftodaybanner
	eCode				=	oTopevtOne.FOneItem.Fevt_code
	Molistbanner		=	oTopevtOne.FOneItem.Fevtmolistbanner
	evttitle2			=	oTopevtOne.FOneItem.Fevttitle2
	gcode				=	oTopevtOne.FOneItem.Fgnbcode

	set oTopevtOne = Nothing
End If 

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
	else
		sDt = Date()
	end if
	sTm = "00:00:00"
end if

if Not(mainEndDate="" or isNull(mainEndDate)) then
	eDt = left(mainEndDate,10)
	eTm = Num2Str(hour(mainEndDate),2,"0","R") &":"& Num2Str(minute(mainEndDate),2,"0","R") &":"& Num2Str(second(mainEndDate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
	else
		eDt = Date()
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

		if (!frm.gcode.value)
		{
			alert('���� GNB ������ ���� ���ּ���.');
			frm.gcode.focus();
			return;
		}

		if (confirm('���� �Ͻðڽ��ϱ�?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/topevtbanner/";
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
    	numberOfMonths: 1,
    	showCurrentAtPos: 0,
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
  var valsdt , valedt , valgcode
	valsdt = document.frm.sDt.value;
	valedt = document.frm.eDt.value;
	valgcode = document.frm.gcode.value;

	if (valgcode == ""){
		valgcode = "<%=gcode%>";
	}else{
		valgcode = document.frm.gcode.value;
	}

  var winLast,eKind;
  winLast = window.open('pop_event_list.asp?menupos=<%=menupos%>&gcode='+valgcode+'&sDt='+valsdt+'&eDt='+valedt,'pLast','width=550,height=600, scrollbars=yes')
  winLast.focus();
}
function chgmu(v){
	if (v == "1")
	{
		$("#sel11").css("display","block");
		$("#sel21").css("display","none");
		$("#sel22").css("display","none");
	}else{
		$("#sel11").css("display","none");
		$("#sel21").css("display","block");
		$("#sel22").css("display","block");
	}
}
</script>
<table width="1000" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/topeventbanner_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">����Ⱓ</td>
    <td colspan="3">
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<!-- �ű��̺�Ʈ ��Ͻ� -->
<% If mode = "add" Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">���� GNB����</td>
	<td colspan="3"><% Call drawSelectBoxGNB("gcode" , gcode) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="30">�̺�Ʈ ��ũŸ��</td>
	<td colspan="3"><label for="load">�̺�Ʈ �ҷ�����</label><input type="radio" value="1" name="linktype" id="load" onclick="chgmu('1');" <%=chkiif(linktype="1","checked","")%>/> <label for="self">�����Է�</label><input type="radio" value="2" name="linktype" id="self" onclick="chgmu('2');"/></td>
</tr>
<tr bgcolor="#FFFFFF" id="sel11" style="display:<%=chkiif(linktype="1","block","none")%>;">
	<td bgcolor="#FFF999" align="center" height="30">�̺�Ʈ�ҷ�����</td>
	<td colspan="3"><input type="button" value="�̺�Ʈ �ҷ�����" onclick="jsLastEvent();"/><br/><img src="<%=Molistbanner%>" width="200"><br/><%=Molistbanner%></td>
</tr>
<tr bgcolor="#FFFFFF" id="sel21" style="display:none;">
	<td bgcolor="#FFF999" align="center" height="30">�̺�Ʈ URL</td>
	<td colspan="3">
		<% IF eCode <> "" And mode = "add" THEN %>
			<input type="hidden" name="linkurl" value="/event/eventmain.asp?eventid=<%=eCode%>">
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
<tr bgcolor="#FFFFFF" id="sel22" style="display:none;">
	<td bgcolor="#FFF999" align="center" width="15%">�̺�Ʈ �̹���</td>
	<td width="45%">
		<input type="file" name="evtimg" class="file" title="�̺�Ʈ #1" require="N" style="width:80%;" />
		<% if evtimg<>"" then %>
		<br>
		<img src="<%= evtimg %>" width="100" /><br><%= evtimg %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">�̺�Ʈ<br/>�̹��� alt</td>
	<td width="20%"><input type="text" name="evtalt" value="<%=evtalt%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�̺�Ʈ ����</td>
	<td width="45%"><input type="text" name="evttitle" value="<%=ename%>" size="40"/><!--</br><input type="text" name="evttitle2" value="" size="40"/>--></td>
	<td bgcolor="#FFF999" align="center" width="10%">�̺�Ʈ ����</td>
	<td width="20%">���� : <input type="radio" name="issalecoupon" value="1" <%=chkiif(issalecoupon = 1,"checked","")%>/> ���� : <input type="radio" name="issalecoupon" value="2" <%=chkiif(issalecoupon = 2,"checked","")%>/> <input type="text" name="issalecoupontxt" size="10" value="<%=issalecoupontxt%>" maxlength="10"/><br/> GIFT : <input type="radio" name="issalecoupon" value="3" <%=chkiif(issalecoupon = 3,"checked","")%>/> ���� : <input type="radio" name="issalecoupon" value="4" <%=chkiif(issalecoupon = 4,"checked","")%>/>&nbsp;</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�̺�Ʈ ������ - ������</td>
	<td colspan="3">
		<input type="text" name="evtstdate" size="10" value="<%=stdt%>" onClick="jsPopCal('evtstdate');"/>
		-
		<input type="text" name="evteddate" size="10" value="<%=eddt%>" onClick="jsPopCal('evteddate');"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">���� ��ȣ</td>
	<td colspan="3"><input type="text" name="sortnum" size="10" value="99" maxlength="3"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�۾��� ���û���</td>
	<td colspan="3"><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<% Else %>
<!-- �̺�Ʈ ������ -->
<input type="hidden" value="<%=linktype%>" name="linktype"/>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">���� GNB����</td>
	<td colspan="3"><% Call drawSelectBoxGNB("gcode" , gcode) %></td>
</tr>
<% If linktype = "1" Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">�̺�Ʈ �̹���</td>
	<td colspan="3"><!-- ������<img src="<%=todaybanner%>" width="100"><br/><%=todaybanner%><br/>�Ź��� --><img src="<%=Molistbanner%>" width="200"><br/><%=Molistbanner%></td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">�̺�Ʈ �̹���</td>
	<td width="45%">
		<input type="file" name="evtimg" class="file" title="�̺�Ʈ #1" require="N" style="width:80%;" />
		<% if evtimg<>"" then %>
		<br>
		<img src="<%= evtimg %>" width="100" /><br><%= evtimg %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">�̺�Ʈ<br/>�̹��� alt</td>
	<td width="20%"><input type="text" name="evtalt" value="<%=evtalt%>" size="20" maxlength="20"/></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�̺�Ʈ ����</td>
	<td width="45%"><input type="text" name="evttitle" value="<%=evttitle%>" size="40"/><!--</br><input type="text" name="evttitle2" value="<%=evttitle2%>" size="40"/>--></td>
	<td bgcolor="#FFF999" align="center" width="10%">�̺�Ʈ ����</td>
	<td width="20%">���� : <input type="radio" name="issalecoupon" value="1" <%=chkiif(issalecoupon = 1,"checked","")%>/> ���� : <input type="radio" name="issalecoupon" value="2" <%=chkiif(issalecoupon = 2,"checked","")%>/> <input type="text" name="issalecoupontxt" size="10" value="<%=issalecoupontxt%>" maxlength="10"/><br/> GIFT : <input type="radio" name="issalecoupon" value="3" <%=chkiif(issalecoupon = 3,"checked","")%>/> ���� : <input type="radio" name="issalecoupon" value="4" <%=chkiif(issalecoupon = 4,"checked","")%>/>&nbsp;</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�̺�Ʈ ������ - ������</td>
	<td colspan="3">
		<input type="text" name="evtstdate" size="10" value="<%=startdate%>" onClick="jsPopCal('evtstdate');"/>
		-
		<input type="text" name="evteddate" size="10" value="<%=enddate%>" onClick="jsPopCal('evteddate');"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�̺�Ʈ URL</td>
	<td colspan="3"><input type="text" name="linkurl" size="80" value="<%=linkurl%>"/>
	<br/><br/>ex)
		<font color="#707070">
		- <span style="cursor:pointer" onClick="putLinkText('event','1')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
		- <span style="cursor:pointer" onClick="putLinkText('itemid','1')">��ǰ�ڵ� ��ũ : /category/category_itemprd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
		</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">���� ��ȣ</td>
	<td colspan="3"><input type="text" name="sortnum" size="10" value="<%=sortnum%>" maxlength="3"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�۾��� ���û���</td>
	<td colspan="3"><textarea name="ordertext" cols="80" rows="8"/><%=ordertext%></textarea></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="4"><input type="button" value=" �� �� " onClick="jsgolist();"/><input type="button" value=" �� �� " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->