<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : enjoy_insert.asp
' Discription : ����� enjoybanner_new
' History : 2014.06.23 ����ȭ
' 		  : 2018.11.28 ������ ���� ��� ��ȹ�� �߰�
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/todayenjoyCls.asp" -->
<%
'###############################################
'�̺�Ʈ �ű� ��Ͻ�
'###############################################
%>
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
Dim eCode
Dim idx , evtimg , subImage2 , subImage3 , subImage4 , isusing , mode
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
Dim ename , stdt , eddt , sortnum , todaybanner , Molistbanner , evttitle2 , etc_opt , subname , modify_Molistbanner
Dim tag_gift , tag_plusone , tag_launching , tag_actively , sale_per , coupon_per , tag_only
Dim itemid1 , itemid2 , itemid3 , addtype , iteminfo
Dim itemname1 ,  itemname2 , itemname3
Dim itemimg1 ,  itemimg2 , itemimg3


	eCode = requestCheckvar(Request("eC"),10)
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
IF eCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
	'�̺�Ʈ ���� ��������
	cEvtCont.fnGetEventCont
	ename	=	db2html(cEvtCont.FEName)
	subname	=	db2html(cEvtCont.FENamesub)
	stdt	=	left(cEvtCont.FESDay, 10)
	eddt	=	left(cEvtCont.FEEDay, 10)

	'�̺�Ʈ ȭ�鼳�� ���� ��������
	cEvtCont.fnGetEventDisplay
	todaybanner = cEvtCont.FEBImgMoToday
	If mode = "add" then
		Molistbanner = cEvtCont.FEBImgMoListBanner
	Else 
		modify_Molistbanner = cEvtCont.FEBImgMoListBanner
	End If 


	dim tmpename
		tmpename = Split(ename,"|") 
			 
	if Ubound(tmpename)>0 then
		ename = tmpename(0)
	end if
	
	set cEvtCont = nothing
END IF



'// ������
If idx <> "" then
	dim oEnjoyeventOne
	set oEnjoyeventOne = new CMainbanner
	oEnjoyeventOne.FRectIdx = idx
	oEnjoyeventOne.GetOneContents()

	linktype			=	oEnjoyeventOne.FOneItem.Flinktype
	evtalt				=	oEnjoyeventOne.FOneItem.Fevtalt
	linkurl				=	oEnjoyeventOne.FOneItem.Flinkurl
	evtimg				=	oEnjoyeventOne.FOneItem.Fevtimg
	evttitle			=	oEnjoyeventOne.FOneItem.Fevttitle
	issalecoupontxt		=	oEnjoyeventOne.FOneItem.Fissalecoupontxt
	startdate			=	left(oEnjoyeventOne.FOneItem.Fevtstdate, 10)
	enddate				=	left(oEnjoyeventOne.FOneItem.Fevteddate, 10)
	issalecoupon		=	oEnjoyeventOne.FOneItem.Fissalecoupon
	mainStartDate		=	oEnjoyeventOne.FOneItem.Fstartdate
	mainEndDate			=	oEnjoyeventOne.FOneItem.Fenddate 
	isusing				=	oEnjoyeventOne.FOneItem.Fisusing
	ordertext			=	oEnjoyeventOne.FOneItem.Fordertext
	sortnum				=	oEnjoyeventOne.FOneItem.Fsortnum
	todaybanner			=	oEnjoyeventOne.FOneItem.Ftodaybanner
	eCode				=	oEnjoyeventOne.FOneItem.Fevt_code
	Molistbanner		=	oEnjoyeventOne.FOneItem.Fevtmolistbanner
	evttitle2			=	oEnjoyeventOne.FOneItem.Fevttitle2
	etc_opt				=	oEnjoyeventOne.FOneItem.Fetc_opt

	tag_only			=	oEnjoyeventOne.FOneItem.Ftag_only
	tag_gift			=	oEnjoyeventOne.FOneItem.Ftag_gift
	tag_plusone			=	oEnjoyeventOne.FOneItem.Ftag_plusone
	tag_launching		=	oEnjoyeventOne.FOneItem.Ftag_launching
	tag_actively		=	oEnjoyeventOne.FOneItem.Ftag_actively
	sale_per			=	oEnjoyeventOne.FOneItem.Fsale_per
	coupon_per			=	oEnjoyeventOne.FOneItem.Fcoupon_per

	itemid1				=	oEnjoyeventOne.FOneItem.Fitemid1
	itemid2				=	oEnjoyeventOne.FOneItem.Fitemid2
	itemid3				=	oEnjoyeventOne.FOneItem.Fitemid3
	addtype				=	oEnjoyeventOne.FOneItem.Faddtype
	iteminfo			=	oEnjoyeventOne.FOneItem.Fiteminfo

	set oEnjoyeventOne = Nothing

	Dim ii
	If addtype = 2 then
		If ubound(Split(iteminfo,"^^")) > 0 Then ' �̹��� 3�� ����
			For ii = 0 To ubound(Split(iteminfo,","))
				If CStr(itemid1) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) Then
					itemname1 = Split(Split(iteminfo,",")(ii),"|")(1)
					itemimg1 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(itemid1) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If

				If CStr(itemid2) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) Then
					itemname2 = Split(Split(iteminfo,",")(ii),"|")(1)
					itemimg2 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(itemid2) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If

				If CStr(itemid3) = CStr(Split(Split(iteminfo,",")(ii),"|")(0)) Then
					itemname3 = Split(Split(iteminfo,",")(ii),"|")(1)
					itemimg3 =  webImgUrl & "/image/small/" & GetImageSubFolderByItemid(itemid3) & "/" & Split(Split(iteminfo,",")(ii),"|")(2)
				End If
			Next 
		End If 
	End If 
End If 

dim dateOption
dateOption = request("dateoption")

if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
		if prevDate = "" then 
			prevDate = sDt
		end if 
	elseif dateOption <> "" then
		sDt = dateOption
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
	elseif dateOption <> "" then
		eDt = dateOption
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

		if (frm.addtype[1].checked && frm.addtype[1].value == 2){
			if (frm.itemid1.value == ""){
				alert("��ǰ�ڵ�1�� �־��ּ���.");
				frm.itemid1.focus();
				return;
			}
			if (frm.itemid2.value == ""){
				alert("��ǰ�ڵ�2�� �־��ּ���.");
				frm.itemid2.focus();
				return;
			}
			if (frm.itemid3.value == ""){
				alert("��ǰ�ڵ�3�� �־��ּ���.");
				frm.itemid3.focus();
				return;
			}
		}

		if (confirm('���� �Ͻðڽ��ϱ�?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/appmanage/today/enjoy/";
//		self.location.href="/admin/appmanage/today/enjoy/?menupos=1633&tabs=1";
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
	valsdt = document.frm.StartDate.value;
	valedt = document.frm.EndDate.value;

  var winLast,eKind;
  winLast = window.open('pop_event_list.asp?menupos=<%=menupos%>&sDt='+valsdt+'&eDt='+valedt,'pLast','width=550,height=600, scrollbars=yes')
  winLast.focus();
}

function chgmu(v){
	if (v == "1")
	{
		$("#sel11").css("display","");
		$("#sel21").css("display","none");
		$("#sel22").css("display","none");
	}else{
		$("#sel11").css("display","none");
		$("#sel21").css("display","");
		$("#sel22").css("display","");
	}
}
function changeForm(){
	var dispOption = document.frm.addtype.value;	
	var link = dispOption == 1 ? "enjoy_insert.asp" : "mainTopExhibition_insert.asp"
	document.location.href= link + "?menupos=<%=menupos%>&prevDate=<%=sDt%>&dispOption="+dispOption;	
}
function chgtype(v){
	if (v == "1"){
		$("#additem1").css("display","none");
		$("#additem2").css("display","none");
		$("#additem3").css("display","none");
		$("#evttitle2").attr("maxlength",60);
	}else if(v == "3"){
		changeForm();
	}else{
		$("#additem1").css("display","");
		$("#additem2").css("display","");
		$("#additem3").css("display","");
		$("#evttitle2").attr("maxlength",30);
	}
}

// ��ǰ���� ����
function fnGetItemInfo(iid,v) {
	$.ajax({
		type: "GET",
		url: "/admin/sitemaster/wcms/act_iteminfo.asp?itemid="+iid,
		dataType: "xml",
		cache: false,
		async: false,
		timeout: 5000,
		beforeSend: function(x) {
			if(x && x.overrideMimeType) {
				x.overrideMimeType("text/xml;charset=euc-kr");
			}
		},
		success: function(xml) {
			if($(xml).find("itemInfo").find("item").length>0) {
				var rst = "<img src='" + $(xml).find("itemInfo").find("item").find("smallImage").text() + "' height='70' /><br/>"
					rst += $(xml).find("itemInfo").find("item").find("itemname").text();
				$("#lyItemInfo"+v).fadeIn();
				$("#lyItemInfo"+v).html(rst);
			} else {
				$("#lyItemInfo"+v).fadeOut();
			}
		},
		error: function(xhr, status, error) {
			$("#lyItemInfo"+v).fadeOut();
			/*alert(xhr + '\n' + status + '\n' + error);*/
		}
	});
}
</script>
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/todayenjoy_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<%'2017 ��ǰ �߰� ver %>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#FFF999" align="center">��ȹ�� Ÿ��</td>
  <td colspan="3">
	<input type="radio" name="addtype" id="typeA" value="1" onclick="chgtype('1');" checked/> <label for="typeA">�⺻��</label>
	<!--<input type="radio" name="addtype" id="typeB" value="2" onclick="chgtype('2');" disabled/> <label for="typeB">�⺻�� + ��ǰ3��</label>&nbsp;<br/>-->
	<input type="radio" name="addtype" id="typeC" value="3" onclick="chgtype('3');" /> <label for="typeB">���λ�ܱ�ȹ��</label>&nbsp;<br/>	
  </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="10%">����Ⱓ</td>
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
	<td bgcolor="#FFF999" align="center" height="30">�̺�Ʈ ��ũŸ��</td>
	<td colspan="3">
		<label for="load">�̺�Ʈ �ҷ�����</label>
		<input type="radio" value="1" name="linktype" id="load" onclick="chgmu('1');" <%=chkiif(linktype="1","checked","")%>/>
		<label for="self">�����Է�</label>
		<input type="radio" value="2" name="linktype" id="self" onclick="chgmu('2');"/>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="sel11" style="display:<%=chkiif(linktype="1","","none")%>;">
	<td bgcolor="#FFF999" align="center" height="30">�̺�Ʈ�ҷ�����</td>
	<td colspan="3"><input type="button" value="�̺�Ʈ �ҷ�����" onclick="jsLastEvent();"/><img src="<%=Molistbanner%>" width="200"><br/><%=Molistbanner%></td>
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
	<td bgcolor="#FFF999" align="center" width="10%">�̺�Ʈ �̹���</td>
	<td>
		<input type="file" name="evtimg" class="file" title="�̺�Ʈ #1" require="N" style="width:50%;" />
		<% if evtimg<>"" then %>
		<br>
		<img src="<%= evtimg %>" width="100" /><br><%= evtimg %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">�̺�Ʈ<br/>�̹��� alt</td>
	<td width="40%"><input type="text" name="evtalt" value="<%=evtalt%>" size="20" maxlength="20"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�̺�Ʈ ����</td>
	<td colspan="3"><input type="text" name="evttitle" value="<%=ename%>" size="40"/></br><input type="text" name="evttitle2" id="evttitle2" value="<%=subname%>" size="70"/></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�̺�Ʈ ������ - ������</td>
	<td colspan="3">
		<input type="text" name="evtstdate" size="10" value="<%=stdt%>" onClick="jsPopCal('evtstdate');"/>
		-
		<input type="text" name="evteddate" size="10" value="<%=eddt%>" onClick="jsPopCal('evteddate');"/>
	</td>
</tr>

<tr bgcolor="#FFFFFF" style="display:none;" id="additem1">
    <td bgcolor="#FFF999" align="center">��ǰ�ڵ�1</td>
    <td colspan="3">
        <input type="text" name="itemid1" value="" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'1')" title="��ǰ�ڵ�" />
        <div id="lyItemInfo1" style="display:none;"></div>
    </td>
</tr>
<tr bgcolor="#FFFFFF" style="display:none;" id="additem2">
    <td bgcolor="#FFF999" align="center">��ǰ�ڵ�2</td>
    <td colspan="3">
        <input type="text" name="itemid2" value="" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'2')" title="��ǰ�ڵ�" />
        <div id="lyItemInfo2" style="display:none;"></div>
    </td>
</tr>
<tr bgcolor="#FFFFFF" style="display:none;" id="additem3">
    <td bgcolor="#FFF999" align="center">��ǰ�ڵ�3</td>
    <td colspan="3">
        <input type="text" name="itemid3" value="" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'3')" title="��ǰ�ڵ�" />
        <div id="lyItemInfo3" style="display:none;"></div>
    </td>
</tr>
<%'2017 ��ǰ �߰� ver %>
<tr bgcolor="#FFFFFF">
  <td bgcolor="#FFF999" align="center">�±�</td>
  <td>
  	<input type="checkbox" name="tag_only" id="tag_only" value="Y"/> <label for="tag_only">�ܵ�</label>
	<input type="checkbox" name="tag_gift" id="tag_gift" value="Y"/> <label for="tag_gift">GIFT</label>
	<input type="checkbox" name="tag_plusone" id="tag_plusone" value="Y"/> <label for="tag_plusone">1+1</label>&nbsp;
	<input type="checkbox" name="tag_launching" id="tag_launching" value="Y"/> <label for="tag_launching">��Ī</label>&nbsp;
	<input type="checkbox" name="tag_actively" id="tag_actively" value="Y"/> <label for="tag_actively">����(�ڸ�Ʈ, �Խ��� , ��ǰ�ı�)</label>&nbsp;<br/>
	<font color="red"><strong>�� �ܵ� > GIFT > 1+1 > ��Ī > ���� ������ ���� �˴ϴ�.��</strong></font>
  </td>
  <td bgcolor="#FFF999" align="center">����/����</td>
  <td>
	<input type="text" name="sale_per" value=""> : ������ ex)<font color="red"><strong>~45%</strong></font></br>
	<input type="text" name="coupon_per" value=""> : ���������� ex)<font color="green"><strong>10%</strong></font></br>
	<font color="red"><strong>���ִ� ��츸 �Է� �ϼ���.��</strong></font>
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
<% If linktype = "1" Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">�̺�Ʈ �̹���</td>
	<td colspan="3"><!-- ������<img src="<%=todaybanner%>" width="100"><br/><%=todaybanner%><br/>�Ź��� --><img src="<%=chkiif(Molistbanner="",modify_Molistbanner,Molistbanner)%>" width="200"><br/><%=chkiif(Molistbanner="",modify_Molistbanner,Molistbanner)%>
	<% If Molistbanner= "" And modify_Molistbanner <> "" then%>
	<br/>�� �ش� �̺�Ʈ�� �̹����� ��� �Ǿ����ϴ� ������ �Ͻø� ������������ ������ �˴ϴ�. �� 
	<% End If %>
	</td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="10%">�̺�Ʈ �̹���</td>
	<td>
		<input type="file" name="evtimg" class="file" title="�̺�Ʈ #1" require="N" style="width:80%;" />
		<% if evtimg<>"" then %>
		<br>
		<img src="<%= evtimg %>" width="100" /><br><%= evtimg %>
		<% end if %>		
	</td>
	<td bgcolor="#FFF999" align="center" width="10%">�̺�Ʈ<br/>�̹��� alt</td>
	<td width="40%"><input type="text" name="evtalt" value="<%=evtalt%>" size="20" maxlength="20"/></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�̺�Ʈ ����</td>
	<td colspan="3"><input type="text" name="evttitle" value="<%=evttitle%>" size="40"/></br><input type="text" name="evttitle2" id="evttitle2" value="<%=evttitle2%>" size="70"/></td>
	
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�̺�Ʈ ������ - ������</td>
	<td colspan="3">
		<input type="text" name="evtstdate" size="10" value="<%=startdate%>" onClick="jsPopCal('evtstdate');"/>
		-
		<input type="text" name="evteddate" size="10" value="<%=enddate%>" onClick="jsPopCal('evteddate');"/>
	</td>
</tr>

<tr bgcolor="#FFFFFF" style="display:<%=chkiif(addtype="2","","none")%>;" id="additem1">
    <td bgcolor="#FFF999" align="center">��ǰ�ڵ�1</td>
    <td colspan="3">
        <input type="text" name="itemid1" value="<%=itemid1%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'1')" title="��ǰ�ڵ�" />
        <div id="lyItemInfo1" style="display:<%=chkIIF(itemid1="","none","")%>;">
		<%
        	if Not(itemName1="" or isNull(itemName1)) then
        		Response.Write "<img src='" & itemimg1 & "' height='70' /><br/>"
        		Response.Write itemName1
        	end if
        %>
		</div>
    </td>
</tr>
<tr bgcolor="#FFFFFF" style="display:<%=chkiif(addtype="2","","none")%>;" id="additem2">
    <td bgcolor="#FFF999" align="center">��ǰ�ڵ�2</td>
    <td colspan="3">
        <input type="text" name="itemid2" value="<%=itemid2%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'2')" title="��ǰ�ڵ�" />
        <div id="lyItemInfo2" style="display:<%=chkIIF(itemid2="","none","")%>;">
		<%
        	if Not(itemName2="" or isNull(itemName2)) then
        		Response.Write "<img src='" & itemimg2 & "' height='70' /><br/>"
        		Response.Write itemName2
        	end if
        %>
		</div>
    </td>
</tr>
<tr bgcolor="#FFFFFF" style="display:<%=chkiif(addtype="2","","none")%>;" id="additem3">
    <td bgcolor="#FFF999" align="center">��ǰ�ڵ�3</td>
    <td colspan="3">
        <input type="text" name="itemid3" value="<%=itemid3%>" size="8" maxlength="8" class="text" require="N" onblur="fnGetItemInfo(this.value,'3')" title="��ǰ�ڵ�" />
        <div id="lyItemInfo3" style="display:<%=chkIIF(itemid3="","none","")%>;">
		<%
        	if Not(itemName3="" or isNull(itemName3)) then
        		Response.Write "<img src='" & itemimg3 & "' height='70' /><br/>"
        		Response.Write itemName3
        	end if
        %>
		</div>
    </td>
</tr>
<%'2017 ��ǰ �߰� ver %>

<tr bgcolor="#FFFFFF">
  <td bgcolor="#FFF999" align="center">�±�</td>
  <td>
  	<input type="checkbox" name="tag_only" id="tag_only" value="Y" <%=chkiif(tag_only = "Y","checked","")%>/> <label for="tag_only">�ܵ�</label><br/>
	<input type="checkbox" name="tag_gift" id="tag_gift" value="Y" <%=chkiif(tag_gift = "Y","checked","")%>/> <label for="tag_gift">GIFT</label>
	<input type="checkbox" name="tag_plusone" id="tag_plusone" value="Y" <%=chkiif(tag_plusone = "Y","checked","")%>/> <label for="tag_plusone">1+1</label>&nbsp;
	<input type="checkbox" name="tag_launching" id="tag_launching" value="Y" <%=chkiif(tag_launching = "Y","checked","")%>/> <label for="tag_launching">��Ī</label>&nbsp;
	<input type="checkbox" name="tag_actively" id="tag_actively" value="Y" <%=chkiif(tag_actively = "Y","checked","")%>/> <label for="tag_actively">����(�ڸ�Ʈ, �Խ��� , ��ǰ�ı�)</label>&nbsp;<br/>
	<font color="red"><strong>�� GIFT > 1+1 > ��Ī > ���� ������ ���� �˴ϴ�.��</strong></font>
  </td>
  <td bgcolor="#FFF999" align="center">����/����</td>
  <td>
	<input type="text" name="sale_per" value="<%=sale_per%>"> : ������ ex)<font color="red"><strong>~45%</strong></font></br>
	<input type="text" name="coupon_per" value="<%=coupon_per%>"> : ���������� ex)<font color="green"><strong>10%</strong></font></br>
	<font color="red"><strong>���ִ� ��츸 �Է� �ϼ���.��</strong></font>
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