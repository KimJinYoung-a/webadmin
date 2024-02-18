<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : ��ġ����Ŀ ���� ���ι�� ��� ������
'	History		: 2014.07.09 ���¿� ����
'#############################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhiker_mainbannerCls.asp"-->

<%
Dim i, mode
Dim sDt, sTm, eDt, eTm
Dim sdate, edate, gubun
Dim srcSDT , srcEDT, stdt, eddt
Dim idx, isusing, sortnum, regdate, linkurl, layerpopurl, SearchGubun, map
Dim sqlstr, sqlsearch, arrlist, resultcount
Dim cEvtCont
	idx = request("idx")
	map = request("map")
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	sdate = request("sdate")
	edate = request("edate")
	mode = request("mode")
	gubun = request("gubun")
	isusing = request("isusing")
	regdate = request("regdate")
	sortnum = request("sortnum")
	linkurl = request("linkurl")
	layerpopurl = request("layerpopurl")
	
dim opart, con_viewthumbimg
	set opart = new CAbouthitchhiker
		opart.fnGetHitchhikerList

if idx = "" then 
	mode="NEW"
else
	mode="EDIT"
end if

if mode="EDIT" then
	if idx <> "" then
		sqlsearch = sqlsearch & " and idx="& idx &""
	end if
		
		sqlstr = "select top 1"
		sqlstr = sqlstr & " idx, linkurl, sdate, edate, isusing, sortnum, gubun, con_viewthumbimg"
		sqlstr = sqlstr & " from db_sitemaster.dbo.tbl_hitchhiker_mainbanner_list"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " order by idx desc"
		
		rsget.Open sqlstr, dbget, 1
		
		resultcount = rsget.recordcount
		
	if not rsget.EOF then
		'suserid = userid
		arrlist = rsget.getrows()
	end if
	
		rsget.close
		
		idx = arrlist(0,0)
		linkurl = arrlist(1,0)
		sdate = arrlist(2,0)
		edate = arrlist(3,0)
		isusing = arrlist(4,0)
		sortnum = arrlist(5,0)
		gubun = arrlist(6,0)
		con_viewthumbimg = arrlist(7,0)
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
	if(frm.SearchGubun.value==""){
		alert("������ ������ �ּ���");
		frm.SearchGubun.focus();
		return;
	}

	if(frm.con_viewthumbimg.value==""){
		alert("�̹����� ����� �ּ���");
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


function chghicprogbn(comp){
    var frm=comp.form;
	location.href="/admin/hitchhiker/mainbanner/hitchhiker_mainbanner_write.asp?idx=<%= idx %>&gubun="+comp;
}

//�̹��� Ȯ��ȭ�� ��â���� �����ֱ�
function jsImgView(sImgUrl){
	var wImgView;
	wImgView = window.open('/admin/itemmaster/colortrend_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
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
	winImg = window.open('/admin/hitchhiker/mainbanner/hitchhiker_mainbanner_imagereg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
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

//��ũ �ؽ�Ʈ�ڽ� ����(����)
function clearFieldColor(field) {
  if (field.value == field.defaultValue) {
      field.style.backgroundColor = "#FFFFFF";
  }
}
function checkFieldColor(field) {
  if (!field.value) {
      field.style.backgroundColor = "#FFDDDD";
  }
} 
</script>

<img src="/images/icon_arrow_link.gif"> <b>��ġ����Ŀ ���ι�� ���</b>
<form name="frm" method="post" action="hitchhiker_mainbanner_proc.asp">
<input type = "hidden" name = "idx" value = "<%=idx %>">
<input type = "hidden" name = "mode" value = "<%=mode %>">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="con_viewthumbimg" value="<%= con_viewthumbimg %>">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if mode = "EDIT"  then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ȣ</td>
		<td colspan="2"><%=idx%></td>
	</tr>
	<% end if %>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
		<td colspan="2">
			<select name="SearchGubun" onChange='chghicprogbn(this.value)'>
				<option value ="" style="color:blue">�� ��</option>
				<option value="1" <% If "1" = cstr(gubun) Then%> selected <%End if%>>��ũ</option>
				<option value="2" <% If "2" = cstr(gubun) Then%> selected <%End if%>>���̾��˾�</option>
				<option value="3" <% If "3" = cstr(gubun) Then%> selected <%End if%>>OnlyView</option>
				<option value="4" <% If "4" = cstr(gubun) Then%> selected <%End if%>>����&�߰�</option>
			</select>
		<% if mode = "NEW" then %>
			<font color="red">�������� �� ���� ������ �ּ���!!</font>
		<% end if %>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�̹���</td>
		<td bgcolor="#FFFFFF">
			<input type="button" name="btnmainbannerimg" value="�̹������" onClick="jsSetImg('<%= con_viewthumbimg %>','con_viewthumbimg','con_viewthumbimgdiv')" class="button">
			<div id="con_viewthumbimgdiv" style="padding: 5 5 5 5">
				<% IF con_viewthumbimg <> "" THEN %>			
					<img src="<%=con_viewthumbimg%>" border="0" width=600 height=300 onclick="jsImgView('<%=con_viewthumbimg %>');" alt="�����ø� Ȯ�� �˴ϴ�">
					<a href="javascript:jsDelImg('con_viewthumbimg','con_viewthumbimgdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
				<% END IF %>
			</div>
		</td>
	</tr>

<% If "1" = cstr(gubun) Then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ũ</td>
		<td colspan="2">
			<font color="red"> �� ��ġ����Ŀ ��ǰ����Ʈ�� �������� ��ũ�϶� �ƹ��͵� �Է����� ������.</font>
			<input type="text" name="linkurl" style="width:100%; background-color:#FFDDDD;" value="<%=linkurl%>" onBlur="checkFieldColor(this);" onFocus="clearFieldColor(this);"/>
		</td>
	</tr>
<% elseif "2" = cstr(gubun) Then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">���̾� �˾�</td>
		<td colspan="2">
			<textarea name="linkurl" class="textarea" style="width:100%; height:150px; background-color:#FFF";"><%= trim(linkurl)%></textarea>
		</td>
	</tr>
<% elseif "4" = cstr(gubun) Then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">����&�߰�</td>
		<td colspan="2">
			<textarea name="linkurl" class="textarea" style="width:100%; height:150px; background-color:#FFF";"><%= trim(linkurl)%></textarea>
		</td>
	</tr>
<% end if %>

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
		<td bgcolor="<%= adminColor("tabletop") %>" align="center"> ��뿩�� </td>
		<td colspan="2">
			<input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; 
			<input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�켱����</td>
		<td colspan="2"><input type="text" name="sortnum" size="10" value="<%=sortnum%>" maxlength="3" onkeydown='onlyNumDecimalInput();' style="ime-mode:disabled"/></td>
	</tr>
	
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