<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : PC���ΰ��� ����Ʈ������
' History : ������ ����
'			2022.07.04 �ѿ�� ����(isms�������ġ)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/just1DayCls2018New.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , isusing , mode, paramisusing, bannerimage
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate , lp
Dim sDt, sTm, eDt, eTm , gubun , title , prevDate , is1day
Dim linkurl, workertext, vplatform
Dim subtitle, saleper
Dim vType

	idx = requestCheckvar(getNumeric(request("idx")),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")
	paramisusing = request("paramisusing")
	vType = request("type")
	vplatform = "pc"

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

		title	=	ReplaceBracket(just1dayList.FOneItem.Ftitle) '// ����(front�� ǥ�õǴ� ��� ����)
		mainStartDate	=	just1dayList.FOneItem.Fstartdate '// ������
		mainEndDate		=	just1dayList.FOneItem.Fenddate '// ������
		isusing			=	just1dayList.FOneItem.Fisusing '// ��뿩��
		saleper			=	just1dayList.FOneItem.Fsaleper '// ������
		vType			=	just1dayList.FOneItem.FType '// type(just1day, event)
		bannerimage		=	just1dayList.FOneItem.FbannerImage '// ��ȹ���� ����̹���
		linkurl			=	just1dayList.FOneItem.FlinkUrl '// ��ȹ���� ��� ��ũurl
		workertext		=	just1dayList.FOneItem.FworkerText '// �۾��� ���޻���(��ȹ���� ���)	
		vplatform		=	just1dayList.FOneItem.Fplatform '// �÷���(pc,mobile)	

		'// 2019-12-11 �������� ��û���� ���� �ָ�Ư��(event) Just1Day �κ� �̹��� ���� ����ϴ��� �ؽ�Ʈ�� ��ü
		'// �ؼ�.. �ָ�Ư��(vType=event)�� ��� Front�� ǥ�õ��� �ʴ� Ÿ��Ʋ�� �̺�Ʈ ����ī�Ƿ�,
		'// workertext�� �̺�Ʈ ����ī�Ƿ� ����ϰ� linkUrl�� webadmin �󿡼� �̺�Ʈ ���ý� �ڵ����� ��������.
		'// �ٸ� �ָ�Ư���� �ƴ� ���� Just1Day�� �״�ΰ�(title Front�� ǥ�� �ȵ� �� ��Ÿ ��� �״��..)

		set just1dayList = Nothing
	End If 

	If Trim(vType)="" Then
		vType = "just1day"
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
			if prevDate = "" then 
				prevDate = sDt
			end if 
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
			if prevDate = "" then 
				prevDate = eDt
			end if 
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

		<% if vType="just1day" then %>
			if (!frm.title.value){
				alert("������ �Է����ּ���.");
				frm.title.focus();
				return;
			}

			if (!frm.saleper.value){
				alert("�ִ��������� Ȯ�� ���ּ���.");
				frm.saleper.focus();
				return;
			}
		<% end if %>

		<% if vType="event" then %>		
			if (!frm.linkurl.value){
				alert("��ȹ���� �������ּ���.");
				frm.linkurl.focus();
				return;
			}

			if (!frm.title.value){
				alert("��ȹ�� ����ī�Ǹ� �Է����ּ���.");
				frm.title.focus();
				return;
			}

			if (!frm.workertext.value){
				alert("��ȹ�� ����ī�Ǹ� �Է����ּ���.");
				frm.workertext.focus();
				return;
			}

			<%'// �ִ� �������� �Է� ���ϴ� ��쵵 �ִٰ� �ؼ� üũ ���� %>
			/*
			if (!frm.saleper.value){
				alert("�ִ��������� Ȯ�� ���ּ���.");
				frm.saleper.focus();
				return;
			}
			*/			
		
			if (frm.linkurl.value.indexOf("�̺�Ʈ��ȣ") > 0 || frm.linkurl.value.indexOf("��ǰ�ڵ�") > 0){
				alert("��� ��ũ ���� Ȯ�� ���ּ���.");
				frm.linkurl.focus();
				return;
			}
		<% end if %>

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
		self.location.href="/admin/sitemaster/just1day2018/?menupos=<%=request("menupos")%>&isusing=<%=paramisusing%>";
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
    var popwin = window.open('popSubItemEdit.asp?listIdx=<%=idx%>&subIdx='+subidx,'popTemplateManage','width=800,height=500,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

function popRegSearchItem() {
<% if idx <> "" then %>
    var popwin = window.open("/admin/sitemaster/just1day2018/popSubItemEdit.asp?listidx=<%=idx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
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
			frm.linkurl.value='/event/eventmain.asp?eventid=�̺�Ʈ��ȣ';
			break;
		case 'itemid':
			frm.linkurl.value='/shopping/category_prd.asp?itemid=��ǰ�ڵ�';
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

function jsChangeTypeJust1Day(typevalue)
{
	location.href='/admin/sitemaster/just1day2018/just1day_insert.asp?menupos=<%=request("menupos")%>&idx=<%=idx%>&sDt='+document.frm.sDt.value+'&eDt='+document.frm.eDt.value+'&prevDate=<%=prevDate%>&paramisusing=<%=paramisusing%>&type='+typevalue;
}

//-- jsLastEvent : ���� �̺�Ʈ �ҷ����� --//
function jsLastEvent(){
	winLast = window.open('pop_event_lastlist.asp','pLast','width=800,height=600, scrollbars=yes')
	winLast.focus();
}

function jsCheckJust1DayEventUrl(){
	if(document.frm.linkurl.value=="") {
		alert("��ȹ���� ���� �������ּ���.");
		return;
	}
	else {
		<% If LCase(application("Svr_Info"))="dev" Then %>
			window.open('http://2015www.10x10.co.kr'+document.frm.linkurl.value);
		<% ElseIf LCase(application("Svr_Info"))="staging" Then %>
			window.open('http://stgwww.10x10.co.kr'+document.frm.linkurl.value);
		<% Else %>
			window.open('http://www.10x10.co.kr'+document.frm.linkurl.value);
		<% End If %>
	}
}
</script>
<form name="frm" method="post" action="dojust1day.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="paramisusing" value="<%=paramisusing%>">
<input type="hidden" name="bannerimage" value="<%=bannerimage%>">
<input type="hidden" name="platform" value="<%=vplatform%>">
<table width="1100" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
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
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=chkiif(mode="add",prevDate,sDt)%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=chkiif(mode="add",prevDate,eDt)%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">����</td>
	<td>
		<div style="float:left;">
			<% If idx="" Then %>
				<input type="radio" name="type" value="just1day" <%=chkiif(vType = "just1day","checked","")%> onclick="jsChangeTypeJust1Day('just1day');"/>JUST 1 DAY
				
				&nbsp;&nbsp;&nbsp; <input type="radio" name="type" value="event"  <%=chkiif(vType = "event","checked","")%> onclick="jsChangeTypeJust1Day('event');"/>��ȹ��
			<% Else %>
				<% If vType="just1day" Then %>
					JUST 1 DAY
					<input type="hidden" name="type" value="just1day">
				<% End If %>
				<% If vType="event" Then %>
					��ȹ��
					<input type="hidden" name="type" value="event">
				<% End If %>
			<% End If %>
		</div>
	</td>
</tr>
<% If vType="just1day" Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">����</td>
	<td>
		<input type="text" name="title" size="50" value="<%=title%>" /> <font color="red">Front�� ǥ�� �ȵ˴ϴ�.</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">������</td>
	<td>
		<input type="text" name="saleper" size="50" value="<%=saleper%>" /> <font color="red">ex) ~91%</font>
	</td>
</tr>
<% End If %>
<% If vtype="event" Then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999"  align="center" width="15%">��ȹ�� ����</td>
		<td>
			<input type="text" name="linkurl" value="<%=linkurl%>" size="50" readonly />&nbsp;&nbsp;<a href="" onclick="jsLastEvent();return false;">[��ȹ�� �ҷ�����]</a>&nbsp;&nbsp;<a href="" onclick="jsCheckJust1DayEventUrl();return false;">[��ȹ�� Ȯ���ϱ�]</a>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999" align="center" width="15%">��ȹ�� ����ī��</td>
		<td>
			<input type="text" name="title" size="50" maxlength="30" value="<%=title%>" /> <font color="red">30�� ���� �Է� �����մϴ�.</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999" align="center" width="15%">��ȹ�� ����ī��</td>
		<td><input type="text" name="workertext" size="80" maxlength="55" value="<%=workertext%>"> <font color="red">55�� ���� �Է� �����մϴ�.</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999" align="center" width="15%">������</td>
		<td>
			<input type="text" name="saleper" size="50" value="<%=saleper%>" /> <font color="red">ex) ~91%</font>
		</td>
	</tr>	
<% End If %>
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
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="1100" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="15">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="left">
		    	�� <%=oSubItemList.FTotalCount%> �� 
		    	<!--input type="button" value="��ü����" class="button" onClick="chkAllItem()">
		    	<input type="button" value="��������" class="button" onClick="saveList()" title="��뿩�θ� �ϰ������մϴ�."-->
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
<col width="50" />
<col width="50" />
<col width="50" />
<col width="200" />
<col width="30" />
<col width="80" />
<col width="80" />
<col width="80" />
<col width="50" />
<col width="50" />
<col width="50" />
<col width="50" />
<tr align="center" bgcolor="#DDDDFF">
    <td>����</td>
    <td>IDX</td>
    <td>��ǰ�ڵ�</td>
    <td>�����</td>
    <td>FrontIMAGE</td>
    <td>����</td>
    <td>������</td>
    <td>���ļ���</td>
    <td>��뿩��</td>
    <td></td>
</tr>
<tbody id="subList">
<%	For lp=0 to oSubItemList.FResultCount-1 %>
<tr align="center" bgcolor="<%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","#FFFFFF","#F3F3F3")%>#FFFFFF">
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;">
	<%
		If oSubItemList.FItemList(lp).Fitemdiv="21" Then
			response.write "����ǰ"
		Else
			response.write "�Ϲݻ�ǰ"
		End If
	%>
	</td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FsubIdx%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FItemid%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FTitle%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;">
    <%
    	if Not(oSubItemList.FItemList(lp).FitemFrontimage="") then
    		Response.Write "<img src='" & oSubItemList.FItemList(lp).FitemFrontimage & "' height='50' />"
		Else
    		Response.Write "<img src='" & oSubItemList.FItemList(lp).FsmallImage & "' height='50' />"
    	end if
    %>
    </td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FitemPrice%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).Fitemsaleper%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).Fsortnum%></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;">
		<%
			If Trim(oSubItemList.FItemList(lp).Fisusing="Y") Then 
				Response.write "���"
			Else
				Response.write "������"
			End If
		%>
	</td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><input type="button" value="����"></td>
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