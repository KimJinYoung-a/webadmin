<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/just1DayCls2018.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , isusing , mode, paramisusing
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate , lp
Dim sDt, sTm, eDt, eTm , gubun , title , prevDate , is1day
Dim extraurl
Dim subtitle , saleper

	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")
	paramisusing = request("paramisusing")
	is1day = request("is1day")

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

If idx <> "" then
	dim just1dayList
	set just1dayList = new Cjust1Day
	just1dayList.FRectIdx = idx
	just1dayList.GetOneContents()

	title	=	just1dayList.FOneItem.Ftitle '// ����(front�� ǥ�õǴ� ��� ����)
	mainStartDate	=	just1dayList.FOneItem.Fstartdate '// ������
	mainEndDate		=	just1dayList.FOneItem.Fenddate '// ������
	isusing			=	just1dayList.FOneItem.Fisusing '// ��뿩��
	is1day			=	just1dayList.FOneItem.Fis1day '// ���а�(J-just1day, E-��ȹ��, W-�ָ�, H-����)

	subImage1		=	just1dayList.FOneItem.FsubImage1 '// ��ϵ� ��� �̹���(��ȹ�������� ���)
	extraurl		=	just1dayList.FOneItem.Fextraurl '// �ָ�, ����, ��ȹ�������� ���
	subtitle		=	just1dayList.FOneItem.Fsubtitle '// �ָ�, ���޿����� ���
	saleper			=	just1dayList.FOneItem.Fsaleper '// just1day, �ָ�, ���޿����� ���

	set just1dayList = Nothing
End If 

Dim oSubItemList
set oSubItemList = new Cjust1Day
	oSubItemList.FPageSize = 100
	oSubItemList.FRectlistIdx = idx
	If idx <> "" then
		oSubItemList.GetContentsItemList()
	End If 


if Not(mainStartDate="" or isNull(mainStartDate)) then
	sDt = left(mainStartDate,10)
	sTm = Num2Str(hour(mainStartDate),2,"0","R") &":"& Num2Str(minute(mainStartDate),2,"0","R") &":"& Num2Str(second(mainStartDate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
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
	else
		eDt = date
	end if
	eTm = "23:59:59"
end If

If Trim(is1day)="" Then
	is1day="Y"
End If
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;

		<% Select Case trim(is1day) %>
			<% case "Y" %>
				if (!frm.title.value){
					alert("������ �Է����ּ���.");
					frm.title.focus();
					return;
				}

				if (!frm.saleper.value){
					alert("�������� Ȯ�� ���ּ���.");
					frm.saleper.focus();
					return;
				}
			<% case "W" %>
				if (!frm.title.value){
					alert("������ �Է����ּ���.");
					frm.title.focus();
					return;
				}

				if (!frm.subtitle.value){
					alert("subtitle�� Ȯ�� ���ּ���.");
					frm.subtitle.focus();
					return;
				}				

				if (!frm.saleper.value){
					alert("�������� Ȯ�� ���ּ���.");
					frm.saleper.focus();
					return;
				}

				if (!frm.extraurl.value){
					alert("��� ��ũ�� �Է����ּ���.");
					frm.extraurl.focus();
					return;
				}

				if (frm.extraurl.value.indexOf("�̺�Ʈ��ȣ") > 0 || frm.extraurl.value.indexOf("��ǰ�ڵ�") > 0){
					alert("��� ��ũ ���� Ȯ�� ���ּ���.");
					frm.extraurl.focus();
					return;
				}
			<% case "H" %>
				if (!frm.title.value){
					alert("������ �Է����ּ���.");
					frm.title.focus();
					return;
				}

				if (!frm.subtitle.value){
					alert("subtitle�� Ȯ�� ���ּ���.");
					frm.subtitle.focus();
					return;
				}				

				if (!frm.saleper.value){
					alert("�������� Ȯ�� ���ּ���.");
					frm.saleper.focus();
					return;
				}

				if (!frm.extraurl.value){
					alert("��� ��ũ�� �Է����ּ���.");
					frm.extraurl.focus();
					return;
				}
			
				if (frm.extraurl.value.indexOf("�̺�Ʈ��ȣ") > 0 || frm.extraurl.value.indexOf("��ǰ�ڵ�") > 0){
					alert("��� ��ũ ���� Ȯ�� ���ּ���.");
					frm.extraurl.focus();
					return;
				}
			<% case "E" %>
				if (!frm.title.value){
					alert("������ �Է����ּ���.");
					frm.title.focus();
					return;
				}
				//if (!frm.todayban.value){
				//	alert("��� �̹����� ������ּ���.");
				//	return;
				//}
				if (!frm.extraurl.value){
					alert("��� ��ũ�� �Է����ּ���.");
					frm.extraurl.focus();
					return;
				}
			
				if (frm.extraurl.value.indexOf("�̺�Ʈ��ȣ") > 0 || frm.extraurl.value.indexOf("��ǰ�ڵ�") > 0){
					alert("��� ��ũ ���� Ȯ�� ���ּ���.");
					frm.extraurl.focus();
					return;
				}

		<% end select %>

		if (!frm.isusing[0].checked && !frm.isusing[1].checked)
		{
			alert("��뿩�θ� �����ϼ���!")
			return false;
		}


		if (confirm('���� �Ͻðڽ��ϱ�?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/sitemaster/chance/?menupos=<%=request("menupos")%>&isusing=<%=paramisusing%>";
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
	
	//������ư
    $(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	$( "#subList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).parent().find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).parent().find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});

});

//����
function popSubEdit(subidx) {
<% if idx <>"" then %>
    var popwin = window.open('popSubItemEdit.asp?listIdx=<%=idx%>&subIdx='+subidx,'popTemplateManage','width=800,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

// ��ǰ�˻� �ϰ� ���
function popRegSearchItem() {
<% if idx <> "" then %>
	var is1day=document.frm.is1day.value;
	if(is1day=="Y")
	{
		is1day="just1day"
	}
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/sitemaster/chance/doSubRegItemCdArray.asp?listidx=<%=idx%>&ptype="+is1day, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

// ��ǰ�ڵ� �ϰ� ���
function popRegArrayItem() {
<% if idx<>"" then %>
    var popwin = window.open('popSubRegItemCdArray.asp?listIdx=<%=idx%>','popRegArray','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}



function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� ���縦 �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.action="doListModify.asp";
		document.frmList.submit();
	}
}

function putLinkText(key) {
	var frm = document.frm;
	switch(key) {
		case 'event':
			frm.extraurl.value='/event/eventmain.asp?eventid=�̺�Ʈ��ȣ';
			break;
		case 'itemid':
			frm.extraurl.value='/shopping/category_prd.asp?itemid=��ǰ�ڵ�';
			break;
	}
}

//-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}


function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<form name="frm" method="post" action="dojust1day.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="paramisusing" value="<%=paramisusing%>">
<input type="hidden" name="todayban" value="<%=subImage1%>">
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
	<% If idx = ""  Then %>
	<td colspan="2" align="center" height="35">��� ���� �� �Դϴ�.</td>
	<% Else %>
	<td bgcolor="#FFF999" colspan="2" align="center" height="35">���� ���� �� �Դϴ�.</td>
	<% End If %>
</tr>
<% If idx <> ""  Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">idx</td>
	<td><%=idx%></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">����Ⱓ</td>
    <td>
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">JUST 1 DAY ����</td>
	<td>
		<div style="float:left;">
			<input type="radio" name="is1day" value="Y" <%=chkiif(is1day = "Y","checked","")%> onclick="location.href='/admin/sitemaster/chance/just1day_insert.asp?idx=<%=idx%>&sDt=<%=srcSDT%>&eDt=<%=srcEDT%>&prevDate=<%=prevDate%>&paramisusing=<%=paramisusing%>&is1day=Y';"/>JUST 1 DAY
			
			&nbsp;&nbsp;&nbsp; <input type="radio" name="is1day" value="W"  <%=chkiif(is1day = "W","checked","")%> onclick="location.href='/admin/sitemaster/chance/just1day_insert.asp?idx=<%=idx%>&sDt=<%=srcSDT%>&eDt=<%=srcEDT%>&prevDate=<%=prevDate%>&paramisusing=<%=paramisusing%>&is1day=W';"/>�ָ�Ư��

			&nbsp;&nbsp;&nbsp; <input type="radio" name="is1day" value="H"  <%=chkiif(is1day = "H","checked","")%> onclick="location.href='/admin/sitemaster/chance/just1day_insert.asp?idx=<%=idx%>&sDt=<%=srcSDT%>&eDt=<%=srcEDT%>&prevDate=<%=prevDate%>&paramisusing=<%=paramisusing%>&is1day=H';"/>����Ư��
			
			&nbsp;&nbsp;&nbsp; <input type="radio" name="is1day" value="E"  <%=chkiif(is1day = "E","checked","")%> onclick="location.href='/admin/sitemaster/chance/just1day_insert.asp?idx=<%=idx%>&sDt=<%=srcSDT%>&eDt=<%=srcEDT%>&prevDate=<%=prevDate%>&paramisusing=<%=paramisusing%>&is1day=E';"/>��ȹ��</div> <div style="float:right;margin-top:5px;margin-right:10px;">
		</div>
	</td>
</tr>
<% Select Case Trim(is1day) %>
	<%'// Just1Day %>
	<% Case "Y" %>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="#FFF999" align="center" width="15%">����</td>
			<td>
				<input type="text" name="title" size="50" value="<%=title%>" /> <font color="red">Front�� ǥ�� �ȵ˴ϴ�.</font>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="#FFF999" align="center" width="15%">������</td>
			<td>
				<input type="text" name="saleper" size="5" value="<%=saleper%>" /> <font color="red">ex) ~60%</font>
			</td>
		</tr>
	<%'// ��ȹ�� %>
	<% Case "E" %>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="#FFF999" align="center" width="15%">����</td>
			<td>
				<input type="text" name="title" size="50" value="<%=title%>" /> <font color="red">Front�� ǥ�� �ȵ˴ϴ�.</font>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" id="tmpstyle1">
			<td bgcolor="#DDDDFF" align="center" width="15%" id="imagetitle">��ȹ���̹���</td>
			<td><input type="button" name="wimg" value="��ȹ�� �̹��� ���" onClick="jsSetImg('today','<%=subImage1%>','todayban','weekendimg')" class="button">
				<div id="weekendimg" style="padding: 5 5 5 5">
					<%IF subImage1 <> "" THEN %>
					<a href="javascript:jsImgView('<%=subImage1%>')"><img  src="<%=subImage1%>" width="400" border="0"></a>
					<a href="javascript:jsDelImg('todayban','weekendimg');"><img src="/images/icon_delete2.gif" border="0"></a>
					<%END IF%>
				</div>
				<%=subImage1%>
			</td>
		</tr>	
		<tr bgcolor="#FFFFFF" id="tmpstyle2">
			<td bgcolor="#DDDDFF"  align="center" width="15%">��ȹ�� �̹������ ��ũ</td>
			<td>
				<input type="text" name="extraurl" value="<%=extraurl%>" style="width:80%;" /><br/>
				- <span style="cursor:pointer" onClick="putLinkText('event')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('itemid')">��ǰ�ڵ� ��ũ : /shopping/category_prd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="#FFF999" align="center" width="15%">�۾����޻���</td>
			<td>
				<textarea name="subtitle" cols="40" rows="5"><%=subtitle%></textarea>
			</td>
		</tr>
	<%'// �ָ�Ư�� %>
	<% Case "W" %>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="#FFF999" align="center" width="15%">����</td>
			<td>
				<input type="text" name="title" size="50" value="<%=title%>" /> <font color="red">Front�� ǥ�� �ȵ˴ϴ�.</font>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="#FFF999" align="center" width="15%">Front�� Subtitle</td>
			<td>
				<input type="text" name="subtitle" size="50" value="<%=subtitle%>" /> <font color="red">�ѱ۷θ� �Է� ����) ���� ��,��,�� 3�ϰ��� Ư���� ����!</font>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="#FFF999" align="center" width="15%">������</td>
			<td>
				<input type="text" name="saleper" size="5" value="<%=saleper%>" /> <font color="red">ex) ~60%</font>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" id="tmpstyle2">
			<td bgcolor="#DDDDFF"  align="center" width="15%">�ָ�Ư�� ��� ��ũ</td>
			<td>
				<input type="text" name="extraurl" value="<%=extraurl%>" style="width:80%;" /><br/>
				- <span style="cursor:pointer" onClick="putLinkText('event')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('itemid')">��ǰ�ڵ� ��ũ : /shopping/category_prd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
			</td>
		</tr>
	<%'// ����Ư�� %>
	<% Case "H" %>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="#FFF999" align="center" width="15%">����</td>
			<td>
				<input type="text" name="title" size="50" value="<%=title%>" /> <font color="red">Front�� ǥ�� �ȵ˴ϴ�.</font>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="#FFF999" align="center" width="15%">Front�� Subtitle</td>
			<td>
				<input type="text" name="subtitle" size="50" value="<%=subtitle%>" /> <font color="red">�ѱ۷θ� �Է� ����) ���޵����� Ư���� ����!</font>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="#FFF999" align="center" width="15%">������</td>
			<td>
				<input type="text" name="saleper" size="5" value="<%=saleper%>" /> <font color="red">ex) ~60%</font>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" id="tmpstyle2">
			<td bgcolor="#DDDDFF"  align="center" width="15%">����Ư�� ��� ��ũ</td>
			<td>
				<input type="text" name="extraurl" value="<%=extraurl%>" style="width:80%;" /><br/>
				- <span style="cursor:pointer" onClick="putLinkText('event')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('itemid')">��ǰ�ڵ� ��ũ : /shopping/category_prd.asp?itemid=<font color="darkred">��ǰ�ڵ� (O)</font></span><br>
			</td>
		</tr>
<% End Select %>

<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>

<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2"><input type="button" value=" �� �� " onClick="jsgolist();"/><input type="button" value=" �� �� " onClick="jsSubmit();"/></td>
</tr>
</table>
</form>

<%
	If idx <> "" then
%>
<!--
<p><b>�� ���� ����</b></p>
<p>
	<strong>
		�� [ON]��ǰ���� >> ��ǰ�������� ��� ��ǰ �˾� Ŭ�� >> ��ǰ��ȣ �ϴ� URL���� >> ������� ��ũ ����<br/>��) http://m.10x10.co.kr/category/category_itemprd.asp?itemid=663507&<span style="color:blue">ldv=<span style="color:red">MzIwMCAg</span></span>
	</strong>
</p>
-->
<!-- // ��ϵ� ���� ��� --------->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="900" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="10">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="left">
		    	�� <%=oSubItemList.FTotalCount%> �� /
		    	<input type="button" value="��ü����" class="button" onClick="chkAllItem()">
		    	<input type="button" value="��������" class="button" onClick="saveList()" title="��뿩�θ� �ϰ������մϴ�.">
		    </td>
		    <td align="right">
		    	<!--<input type="button" value="��ǰ�ڵ�� ���" class="button" onClick="popRegArrayItem()" />//-->
		    	<input type="button" value="��ǰ �߰�" class="button" onClick="popRegSearchItem()" />
		    	<!--<img src="/images/icon_new_registration.gif" border="0" onclick="popSubEdit('')" style="cursor:pointer;" align="absmiddle">//-->
		    </td>
		</tr>
		</table>
	</td>
</tr>
<col width="30" />
<col width="30" />
<col width="30" />
<col span="3" width="0*" />
<col width="30" />
<col width="30" />
<col width="30" />
<col width="30" />
<tr align="center" bgcolor="#DDDDFF">
    <td>&nbsp;</td>
    <td>�����ȣ</td>
    <td>�̹���</td>
    <td>��ǰ�ڵ�</td>
    <td>��ǰ��</td>
    <td>���Ĺ�ȣ</td>
    <td>��</td>
    <td>��뿩��</td>
</tr>
<tbody id="subList">
<%	For lp=0 to oSubItemList.FResultCount-1 %>
<tr align="center" bgcolor="<%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","#FFFFFF","#F3F3F3")%>#FFFFFF">
    <td><input type="checkbox" name="chkIdx" value="<%=oSubItemList.FItemList(lp).FsubIdx%>" /></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FsubIdx%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;">
    <%
    	if Not(oSubItemList.FItemList(lp).FsmallImage="" or isNull(oSubItemList.FItemList(lp).FsmallImage)) then
    		Response.Write "<img src='" & oSubItemList.FItemList(lp).FsmallImage & "' height='50' />"
    	end if
    %>
    </td>
    <td>
    <%
    	if Not(oSubItemList.FItemList(lp).FItemid="0" or isNull(oSubItemList.FItemList(lp).FItemid) or oSubItemList.FItemList(lp).FItemid="") then
    		Response.Write "[<input type='text' value='" & oSubItemList.FItemList(lp).FItemid & "' size=5 />]"
    	end if
    %>
	<!--
	&nbsp;ldv=<input type="text" name="ldv<%=oSubItemList.FItemList(lp).FsubIdx%>" value="<%=oSubItemList.FItemList(lp).Fldv%>"/ size="5">
	-->
    </td>
	<td align="left" style="padding-left:5px;">
		��ǰ�� : <input type="text" name="itemname<%=oSubItemList.FItemList(lp).FsubIdx%>" value="<%=oSubItemList.FItemList(lp).Fitemname%>"/ size="35"><br/>
	</td>
    <td><input type="text" name="sort<%=oSubItemList.FItemList(lp).FsubIdx%>" size="3" class="text" value="<%=oSubItemList.FItemList(lp).Fsortnum%>" style="text-align:center;" /></td>
    <td>
		<select name="label<%=oSubItemList.FItemList(lp).FsubIdx%>">
			<option value="0" <%=chkiif(oSubItemList.FItemList(lp).Flabel="","selected","")%>>����</option>
			<option value="1" <%=chkiif(oSubItemList.FItemList(lp).Flabel="1","selected","")%>>10x10 ONLY</option>
			<option value="2" <%=chkiif(oSubItemList.FItemList(lp).Flabel="2","selected","")%>>HOT ITEM</option>
			<option value="3" <%=chkiif(oSubItemList.FItemList(lp).Flabel="3","selected","")%>>WISH NO.1</option>
			<option value="4" <%=chkiif(oSubItemList.FItemList(lp).Flabel="4","selected","")%>>BEST ITEM</option>
			<option value="5" <%=chkiif(oSubItemList.FItemList(lp).Flabel="5","selected","")%>>1Day</option>
			<option value="6" <%=chkiif(oSubItemList.FItemList(lp).Flabel="6","selected","")%>>����</option>
			<option value="7" <%=chkiif(oSubItemList.FItemList(lp).Flabel="7","selected","")%>>1+1</option>
			<option value="8" <%=chkiif(oSubItemList.FItemList(lp).Flabel="8","selected","")%>>����</option>
			<option value="9" <%=chkiif(oSubItemList.FItemList(lp).Flabel="9","selected","")%>>������</option>
		</select>
	</td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">���</label><input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">����</label>
		</span>
    </td>
</tr>
<% Next %>
</tbody>
</table>
</form>
<%
	End If 
	set oSubItemList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->