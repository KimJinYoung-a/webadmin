<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ġ����Ŀ ������
' Hieditor : 2014.07.17 ���¿� ����
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/test/jtest/about_hitchhiker_contentsCls.asp"-->

<%
Dim i, mode
Dim hicprogbn
Dim sDt, sTm, eDt, eTm
Dim sdate, edate, ename, eCode
Dim sqlstr, sqlsearch, arrlist, resultcount
Dim srcSDT , srcEDT, stdt, eddt, todaybanner
Dim contentsidx, con_title, isusing, sortnum, regdate, con_detail, con_movieurl
Dim cEvtCont
	contentsidx = request("contentsidx")
	hicprogbn = requestCheckvar(Request("hicprogbn"),1)

dim opart, con_viewthumbimg
	set opart = new CAbouthitchhiker
		opart.Frectcontentsidx=contentsidx
		if contentsidx <> "" then
			opart.getHitchhiker_oneitem
			if opart.FResultCount > 0 then
				stdt = opart.Foneitem.FSdate
				eddt = opart.Foneitem.FEdate
				isusing = opart.Foneitem.FIsusing
				hicprogbn = opart.Foneitem.Fgubun
				contentsidx = opart.Foneitem.Fcontentsidx
				con_title = db2html(opart.Foneitem.Fcon_title)
				con_detail = db2html(opart.Foneitem.Fcon_detail)
				con_movieurl = db2html(opart.Foneitem.Fcon_movieurl)
				con_viewthumbimg = opart.Foneitem.Fcon_viewthumbimg
			end if
		end if

'���� idx���� �������(�űԵ��) NEW, �ƴҰ��(����) EDIT
if contentsidx = "" then
	mode="NEW"
else
	mode="EDIT"
end if

dim odevice
set odevice=new CAbouthitchhiker

	if hicprogbn="1" then
		odevice.Frectisusing="Y"
		odevice.Frectgubun="1"
		if contentsidx <> "" then
			odevice.Frectcontentsidx=contentsidx
			odevice.fnGetContents_link
		else
			odevice.fnGetDeviceList
		end if
	elseif hicprogbn="2" then
		odevice.Frectisusing="Y"
		odevice.Frectgubun="2"
		if contentsidx <> "" then
			odevice.Frectcontentsidx=contentsidx
			odevice.fnGetContents_link
		else
			odevice.fnGetDeviceList
		end if
	end if
%>

<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>

function frmedit(){
	if(frm.hicprogbn.value==""){
		alert('���а��� ������ �ּ���');
		frm.hicprogbn.focus();
		return;
	}
	var tmphicprogbn = frm.hicprogbn.value;

	if(tmphicprogbn == "1"){ //���а��� PC �϶� üũ�ؾ� �� �͵�

	}else if (tmphicprogbn == "2"){ //���а��� PC �϶� üũ�ؾ� �� �͵�


	}else if (tmphicprogbn == "3"){ //���а��� MOVIE �϶� üũ�ؾ� �� �͵�
		if(frm.con_detail.value==""){
			alert('�� ������ �Է��� �ּ���');
			return;
		}
	}
	if(frm.con_title.value==""){
		alert('Ÿ��Ʋ�� �Է��� �ּ���');
		frm.con_title.focus();
		return;
	}
	if(frm.con_sdate.value==""){
		alert('�������� �Է��� �ּ���');
		frm.con_sdate.focus();
		return;
	}

	var tmpisusing = "";
	for(var i = 0;  i < frm.isusing.length; i++){
		if(frm.isusing[i].checked==true){
		tmpisusing = frm.isusing[i].value;
		}
	}

	if(tmpisusing == ""){
		alert('��뿩�θ� �����ϼ���');
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
		<% if contentsidx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
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
		<% if contentsidx<>"" then %>minDate: "<%=sDt%>",<% end if %>
		onClose: function( selectedDate ) {
			$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
		}
	});
});

function chghicprogbn(comp){
    var frm=comp.form;
	location.href="/test/jtest/about_contents_write.asp?contentsidx=<%= contentsidx %>&hicprogbn="+comp;
}

//�̹��� Ȯ��ȭ�� ��â���� �����ֱ�
function jsImgView(sImgUrl){
	var wImgView;
	wImgView = window.open('/admin/itemmaster/colortrend_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	eval("document.all."+sName).value = "";
	eval("document.all."+sSpan).style.display = "none";
	}
}

function jsSetImg(sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/test/jtest/hitchhiker_imagereg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}
</script>

<form name="frm" method="post" action="/test/jtest/about_contents_proc.asp">
<input type="hidden" name="mode" value="<%=mode %>">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="contentsidx" value="<%=contentsidx %>">
<input type="hidden" name="con_viewthumbimg" value="<%= con_viewthumbimg %>">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="left">
			<b>����ġ����Ŀ ������ ���</b>
		</td>
	</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% IF contentsidx <> "" THEN%>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ȣ</td>
		<td colspan="2"><%=contentsidx%></td>
	</tr>
	<% End if %>

	<tr bgcolor="#FFFFFF">
		<td  bgcolor="<%= adminColor("tabletop") %>" align="center" width="100px">����</td>
			<td>
				<% Call DrawSelectBoxHitchhikerGubun("hicprogbn",hicprogbn,"onChange='chghicprogbn(this.value)'") %><% if mode = "NEW" then %><font color="red">������ ������ �ּ���!!</font><% end if %>
			</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">Ÿ��Ʋ</td>
		<td colspan="2">
			<input type="text" name="con_title" size="50" value="<%=trim(con_title)%>"/>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">������</td>
		<td colspan="2">
				<input type="text" id="sDt" name="con_sdate" size="10" value="<%=stdt%>" />
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
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="btnhicthumbimg" value="<%= con_viewthumbimg %>" >
			<div id="con_viewthumbimgdiv" style="padding: 5 5 5 5">
				<% IF con_viewthumbimg <> "" THEN %>
					<img src="<%=con_viewthumbimg%>" border="0" width=100 height=100 onclick="jsImgView('<%=con_viewthumbimg %>');" alt="�����ø� Ȯ�� �˴ϴ�">
					<a href="javascript:jsDelImg('con_viewthumbimg','con_viewthumbimgdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
				<% END IF %>
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center"> con_movieurl </td>
		<td colspan="2"><textarea name="con_movieurl" class="textarea" style="width:100%; height:150px;"><%= trim(con_movieurl)%></textarea></td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center"> con_detail </td>
		<td colspan="2"><input type="text" name="con_detail" size="50" value="<%= trim(con_detail) %>"/></td>
	</tr>


	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="3">
			<% if mode = "EDIT" or mode = "NEW" then %>
				<input type="button" name="editsave" value="����" onclick="frmedit()" />
			<% end if %>

			<input type="button" name="editclose" value="���" onclick="self.close()" />
		</td>
	</tr>
</table>
</form>
<%
set opart = nothing
set odevice = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
