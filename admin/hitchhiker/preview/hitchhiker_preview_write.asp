<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : ��ġ����Ŀ ���� ������ ��� ������
'	History		: 2014.08.01 ���¿� ����
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhiker_previewCls.asp"-->

<%
Dim i, mode
Dim sdate, edate
Dim sDt, sTm, eDt, eTm
Dim srcSDT , srcEDT, stdt, eddt
Dim idx, title, isusing, sortnum, regdate, preview_detail, cash, mileage
Dim sqlstr, sqlsearch, arrlist, resultcount
Dim cEvtCont
	idx = request("idx")
	mode = request("mode")
	cash = request("cash")
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	sdate = request("sdate")
	edate = request("edate")
	title = request("title")
	mileage = request("mileage")
	isusing = request("isusing")
	regdate = request("regdate")
	sortnum = request("sortnum")
	preview_detail = request("preview_detail")
	preview_thumbimg = request("preview_thumbimg")

if idx = "" then 
	mode="NEW"
else
	mode="EDIT"
end if

dim opart, preview_thumbimg
	set opart = new CHitchhikerPreview
		opart.FrectIdx = idx

	if idx <> "" then
		opart.sbpreviewwrite
	end if
	
	if opart.ftotalcount > 0 then
		idx = opart.FOneItem.Fidx
		title = opart.FOneItem.FReqTitle
		preview_detail = opart.FOneItem.FReqpreview_detail
		preview_thumbimg = opart.FOneItem.FReqpreview_thumbimg
		isusing = opart.FOneItem.FReqIsusing
		sortnum = opart.FOneItem.FReqSortnum
		regdate = opart.FOneItem.FReqregdate
		sdate = opart.FOneItem.FReqsdate
		edate = opart.FOneItem.FReqedate
		cash = opart.FOneItem.FReqcash
		mileage = opart.FOneItem.FReqmileage
	end if

if Not(sdate="" or isNull(sdate)) then
	sDt = left(sdate,10)
	sTm = Num2Str(hour(sdate),2,"0","R") &":"& Num2Str(minute(sdate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
	else
		sDt = date
	end if
	sTm = "00:00"
end if

if Not(edate="" or isNull(edate)) then
	eDt = left(edate,10)
	eTm = Num2Str(hour(edate),2,"0","R") &":"& Num2Str(minute(edate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
	else
		eDt = date
	end if
	eTm = "23:59"
end If

IF sortnum = "" then sortnum = "99"
%>

<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
function frmedit(){
	if(frm.StartDate.value==""){
		alert("Ÿ��Ʋ�� �Է��� �ּ���");
		frm.title.focus();
		return;
	}
	if(frm.StartDate.value==""){
		alert("�������� ������ �ּ���");
		frm.StartDate.focus();
		return;
	}
	if(frm.EndDate.value==""){
		alert("�������� ������ �ּ���");
		frm.EndDate.focus();
		return;
	}
	if(frm.sortnum.value==""){
		alert("�켱������ �Է��� �ּ���");
		frm.sortnum.focus();
		return;
	}
	if(frm.cash.value!=''){
		if (!IsDouble(frm.cash.value)){
			alert('��ǰ�ڵ�� ���ڸ� �����մϴ�.');
			frm.cash.focus();
			return;
		}
	}
	if(frm.mileage.value!=''){
		if (!IsDouble(frm.mileage.value)){
			alert('��ǰ�ڵ�� ���ڸ� �����մϴ�.');
			frm.mileage.focus();
			return;
		}
	}

	frm.submit();
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

//�̹��� ��â Ȯ�뺸��
function jsImgView(sImgUrl){
	var wImgView;
	wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}
//�̹��� ����
function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	eval("document.all."+sName).value = "";
	eval("document.all."+sSpan).style.display = "none";
	}
}
//�̹��� ���
function jsSetImg(sImg, sName, sSpan){	
	document.domain ="10x10.co.kr";	
	var winImg;
	winImg = window.open('/admin/hitchhiker/preview/hitchhiker_preview_imagereg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}
//�ѱ� �Է� �ȵǰ�
function onlyNumDecimalInput(){
	var code = window.event.keyCode; 
	if ((code >= 48 && code <= 57) || (code >= 96 && code <= 105) || code == 110 || code == 190 || code == 8 || code == 9 || code == 13 || code == 46){ 
		window.event.returnValue = true; 
		return; 
	} 
	window.event.returnValue = false; 
}
</script>

<!-- #include virtual="/admin/hitchhiker/inc_HichHead.asp"-->
<img src="/images/icon_arrow_link.gif"> <b>��ġ����Ŀ ������ ���</b>
<form name="frm" method="post" action="hitchhiker_preview_proc.asp">
<input type = "hidden" name = "idx" value = "<%=idx %>">
<input type = "hidden" name = "mode" value = "<%=mode %>">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="preview_thumbimg" value="<%= preview_thumbimg %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if mode = "EDIT"  then %>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ȣ</td>
			<td colspan="2"><%=idx%></td>
		</tr>
	<% end if %>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">Ÿ��Ʋ</td>
		<td colspan="2">
			<input type="text" name="title" style="width:100%" value="<%=title%>"/>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�󼼳���</td>
		<td colspan="2">
			<input type="text" name="preview_detail" style="width:100%" value="<%=preview_detail%>"/>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�Ⱓ</td>
		<td colspan="2">
			<% if mode = "NEW" then %>
				<input type="text" id="sDt" name="StartDate" size="10" value="<%=stdt%>" />
				<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
				<input type="text" id="eDt" name="EndDate" size="10" value="<%=eddt%>" />
				<input type="text" name="eTm" size="8" value="<%=eTm%>" />
			<% else %>
				<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
				<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
				<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
				<input type="text" name="eTm" size="8" value="<%=eTm%>" />
			<% end if %>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����</td>
		<td bgcolor="#FFFFFF">
			<input type="button" name="btnmainbannerimg" value="����� �̹������" onClick="jsSetImg('<%= preview_thumbimg %>','preview_thumbimg','preview_thumbimgdiv')" class="button">
			<div id="preview_thumbimgdiv" style="padding: 5 5 5 5">
				<% IF preview_thumbimg <> "" THEN %>			
					<img src="<%=preview_thumbimg%>" border="0" width=300 height=150 onclick="jsImgView('<%=preview_thumbimg %>');" alt="�����ø� Ȯ�� �˴ϴ�">
					<a href="javascript:jsDelImg('preview_thumbimg','preview_thumbimgdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
				<% END IF %>
			</div>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�켱����</td>
		<td colspan="2"><input type="text" name="sortnum" size="10" value="<%=sortnum%>" maxlength="3" onkeydown='onlyNumDecimalInput();' style="ime-mode:disabled"/></td>
	</tr>	
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ǰ�ڵ�</td>
		<td colspan="2">
			���ݱ��� :		<input type="text" name="cash" size="10" value="<%= cash %>" maxlength="10" />&nbsp;&nbsp;&nbsp;
			���ϸ������� :	<input type="text" name="mileage" size="10" value="<%= mileage %>" maxlength="10">
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center"> ��뿩�� </td>
		<td colspan="2">
			<input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp;
			<input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������
		</td>
	</tr>
	
	<% If mode = "EDIT" Then %>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">������<br>�̹���(WWW)</td>
			<td bgcolor="#FFFFFF">
				<iframe id="iframG" frameborder="0" width="100%" src="/admin/hitchhiker/preview/iframe_hitchhiker_preview.asp?idx=<%=idx%>" height=300></iframe>
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">������<br>�̹���(�����)<br>�ִ� 4��</td>
			<td bgcolor="#FFFFFF">
				<iframe id="iframF" frameborder="0" width="100%" src="/admin/hitchhiker/preview/iframe_hitchhiker_preview_M.asp?idx=<%=idx%>" height=300></iframe>
			</td>
		</tr>
	<% else %>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">������<br>�̹���</td>
			<td bgcolor="#FFFFFF">
				�űԵ�� �Ϸ��� PreView�̹����� �Է� �ϽǼ� �ֽ��ϴ�.
			</td>
		</tr>
	<% End If %>

	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="3">
				<% if mode = "EDIT" or mode = "NEW" then %>
					<input type="button" class="button" uname="editsave" value="����" onclick="frmedit()" />	
				<% end if %>
					<input type="button" class="button" name="editclose" value="���" onclick="self.close()" />
		</td>
	</tr>
</table>
</form>
<%
set opart = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->