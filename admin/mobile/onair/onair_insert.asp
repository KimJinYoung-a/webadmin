<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : onair_insert.asp
' Discription : ����� onair
' History : 2013.12.15 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/onairCls.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , isusing , mode
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate , lp
Dim sDt, sTm, eDt, eTm , gubun , onairtitle
Dim ctitle , cper , cnum , cgubun , prevDate

	idx = requestCheckvar(request("idx"),16)
	srcSDT = request("sDt")
	srcEDT = request("eDt")
	prevDate = request("prevDate")

If idx = "" Then 
	mode = "add" 
Else 
	mode = "modify" 
End If 

If idx <> "" then
	dim oOnAirList
	set oOnAirList = new COnAir
	oOnAirList.FRectIdx = idx
	oOnAirList.GetOneContents()

	gubun				=	oOnAirList.FOneItem.Fgubun
	onairtitle			=	oOnAirList.FOneItem.Fonairtitle
	mainStartDate	=	oOnAirList.FOneItem.Fstartdate
	mainEndDate	=	oOnAirList.FOneItem.Fenddate 
	isusing				=	oOnAirList.FOneItem.Fisusing

	ctitle					=	oOnAirList.FOneItem.Fctitle
	cper					=	oOnAirList.FOneItem.Fcper
	cnum				=	oOnAirList.FOneItem.Fcnum
	cgubun				=	oOnAirList.FOneItem.Fcgubun

	set oOnAirList = Nothing
End If 

Dim oSubItemList
set oSubItemList = new COnAir
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
	sTm = "08:00:00"
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
	eTm = "11:59:59"
end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;
	
		if (confirm('���� �Ͻðڽ��ϱ�?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/onair/";
	}
	$(function(){
	//�޷´�ȭâ ����
	var arrDayMin = ["��","��","ȭ","��","��","��","��"];
	var arrMonth = ["1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��"];
	var sTime = document.frm.sTm.value;
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
			 document.frm.onairtitle.value = selectedDate +" "+ sTime + " ���� onair �Դϴ�";
			
			 if (document.frm.gubun[3].checked)
			 {
				nextdate(selectedDate);// ����Ⱓ ������ ǥ�� �ٲٱ�
			 }
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

function nextdate(date){
	var ndate = date
	var dateinfo = ndate.split("-");
	var tdate =  new Date(dateinfo[0],dateinfo[1]-1,dateinfo[2]);

	var tyear = tdate.getFullYear(dateinfo[0]);
	var tmonth=tdate.getMonth(dateinfo[1]-1) + 1;
	tmonth = (tmonth<10)? '0'+ tmonth : tmonth;
	var tday=tdate.getDate(dateinfo[2])+1; 
	tday = (tday<10)? '0'+ tday : tday;

	var nextdate = tyear+'-'+tmonth+'-'+tday;

	frm.eDt.value = nextdate;
}

function jsautotime(v){
	var frm = document.frm;
	var ndate = frm.StartDate.value;
	var dateinfo = ndate.split("-");
	var tdate =  new Date(dateinfo[0],dateinfo[1]-1,dateinfo[2]);

	var tyear = tdate.getFullYear(dateinfo[0]);
	var tmonth=tdate.getMonth(dateinfo[1]-1) + 1;
	tmonth = (tmonth<10)? '0'+ tmonth : tmonth;
	var tday=tdate.getDate(dateinfo[2])+1; 
	tday = (tday<10)? '0'+ tday : tday;
	var nextdate = tyear+'-'+tmonth+'-'+tday;

	if (v == 1){
		frm.sDt.value = ndate;
		frm.eDt.value = ndate;
		frm.sTm.value = "08:00:00";
		frm.eTm.value = "11:59:59";
		frm.onairtitle.value= frm.sDt.value +" "+ frm.sTm.value + " ���� onair �Դϴ�";
	}else if (v == 2){
		frm.sDt.value = ndate;
		frm.eDt.value = ndate;
		frm.sTm.value = "12:00:00";
		frm.eTm.value = "17:59:59";
		frm.onairtitle.value= frm.sDt.value +" "+ frm.sTm.value + " ���� onair �Դϴ�";
	}else if (v == 3){
		frm.sDt.value = ndate;
		frm.eDt.value = ndate;
		frm.sTm.value = "18:00:00";
		frm.eTm.value = "22:59:59";
		frm.onairtitle.value= frm.sDt.value +" "+ frm.sTm.value + " ���� onair �Դϴ�";
	}else{
		frm.sDt.value = ndate;
		frm.eDt.value = nextdate;
		frm.sTm.value = "23:00:00";
		frm.eTm.value = "07:59:59";
		frm.onairtitle.value= frm.sDt.value +" "+ frm.sTm.value + " ���� onair �Դϴ�";
	}
}
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
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/mobile/onair/doSubRegItemCdArray.asp?listidx=<%=idx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
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
</script>
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="doonair.asp">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<tr bgcolor="#FFFFFF">
	<% If idx = ""  Then %>
	<td colspan="5" align="center" height="35">��� ���� �� �Դϴ�.</td>
	<% Else %>
	<td bgcolor="#FFF999" colspan="5" align="center" height="35">���� ���� �� �Դϴ�.</td>
	<% End If %>
</tr>
<% If idx <> ""  Then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" width="15%">idx</td>
	<td colspan="5"><%=idx%></td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">����</td>
    <td colspan="3">
		&nbsp;<label for="t1">8��</label><input type="radio" id= "t1" name="gubun" value="1" onclick="jsautotime(1)" <%=chkiif(gubun="1" Or gubun="" ,"checked","")%>/>
		<label for="t2">12��</label><input type="radio" id= "t2" name="gubun" value="2" onclick="jsautotime(2)" <%=chkiif(gubun="2","checked","")%>/>
		<label for="t3">18��</label><input type="radio" id= "t3" name="gubun" value="3" onclick="jsautotime(3)" <%=chkiif(gubun="3","checked","")%>/>
		<label for="t4">23��</label><input type="radio" id= "t4" name="gubun" value="4" onclick="jsautotime(4)" <%=chkiif(gubun="4","checked","")%>/>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">����Ⱓ</td>
    <td colspan="3">
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">����</td>
    <td colspan="3">
		<input type="text" name="onairtitle" size="100" value="<%If idx="" then%><%=sDt%>&nbsp;<%=sTm%>&nbsp;���� onair �Դϴ�<% Else %><%=onairtitle%><% End if%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td colspan="3"><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center">Ÿ������</td>
    <td width="40%">
		���� : <input type="text" name="ctitle" size="30" value="<%=ctitle%>" /><br/>
		�������� : <select name="cgubun" >
							<option value="I" <%=chkiif(cgubun="I","selected","")%>>��ǰ����</option>
							<option value="B" <%=chkiif(cgubun="B","selected","")%>>���ʽ���������</option>
					   </select>
    </td>
	<td width="25%">
		���η�% : <input type="text" name="cper" size="5" value="<%=cper%>" />%
	</td>
	<td width="25%">
		������ȣ : <input type="text" name="cnum" size="7" value="<%=cnum%>" />
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="5"><input type="button" value=" �� �� " onClick="jsgolist();"/><input type="button" value=" �� �� " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>

<%
	If idx <> "" then
%>
<p><b>�� ���� ����</b></p>
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
		    	<input type="button" value="��ü����" class="button" onClick="chkAllItem()" />
		    	<input type="button" value="��������" class="button" onClick="saveList()" title="ǥ�ü��� �� ��뿩�θ� �ϰ������մϴ�."/>
		    </td>
		    <td align="right">
		    	<input type="button" value="��ǰ�ڵ�� ���" class="button" onClick="popRegArrayItem()" />
		    	<input type="button" value="��ǰ �߰�" class="button" onClick="popRegSearchItem()" />
		    	<img src="/images/icon_new_registration.gif" border="0" onclick="popSubEdit('')" style="cursor:pointer;" align="absmiddle">
		    </td>
		</tr>
		</table>
	</td>
</tr>
<col width="30" />
<col width="60" />
<col span="4" width="0*" />
<col width="70" />
<col width="110" />
<col width="80" />
<col width="80" />
<tr align="center" bgcolor="#DDDDFF">
    <td>&nbsp;</td>
    <td>�����ȣ</td>
    <td>�̹���</td>
    <td>��ǰ�ڵ�</td>
    <td>ǥ�ü���</td>
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
    		Response.Write "[" & oSubItemList.FItemList(lp).FItemid & "]" & oSubItemList.FItemList(lp).Fitemname
    	end if
    %>
    </td>
    <td><input type="text" name="sort<%=oSubItemList.FItemList(lp).FsubIdx%>" size="3" class="text" value="<%=oSubItemList.FItemList(lp).Fsortnum%>" style="text-align:center;" /></td>
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