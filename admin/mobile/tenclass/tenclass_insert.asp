<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : tenclass_insert.asp
' Discription : ����� tenclass
' History : 2018-02-27 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/tenclass_Cls.asp" -->
<%
Dim idx , isusing , mode
Dim srcSDT , srcEDT
Dim mainStartDate, mainEndDate , lp , ii
Dim sDt, sTm, eDt, eTm , gubun , prevDate
Dim mainimg , maincopy , subcopy , adminnotice , mainimage
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
	dim tenClassList
	set tenClassList = new tenClass
	tenClassList.FRectIdx = idx
	tenClassList.GetOneContents()

	maincopy		=	tenClassList.FOneItem.Fmaincopy
	subcopy			=	tenClassList.FOneItem.Fsubcopy
	mainStartDate	=	tenClassList.FOneItem.Fstartdate
	mainEndDate		=	tenClassList.FOneItem.Fenddate
	isusing			=	tenClassList.FOneItem.Fisusing
	adminnotice		=	tenClassList.FOneItem.Fadminnotice
	mainimage		=	tenClassList.FOneItem.Fmainimage

	set tenClassList = Nothing
End If

Dim oSubItemList
set oSubItemList = new tenClass
	oSubItemList.FPageSize = 100
	oSubItemList.FRectidx = idx
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
end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;

		if (frm.sTm.value.length != 8) {
			alert("�ð��� ��Ȯ�� �Է��ϼ���");
			frm.sTm.focus();
			return;
		}

		if (frm.eTm.value.length != 8) {
			alert("�ð��� ��Ȯ�� �Է��ϼ���");
			frm.eTm.focus();
			return;
		}

		if (confirm('���� �Ͻðڽ��ϱ�?')){
			//frm.target = "blank";
			frm.submit();
		}
	}
	function jsgolist(){
		self.location.href="/admin/mobile/tenclass/";
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
			$(this).find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
});

// ��ǰ�˻� �ϰ� ��� (������)
function popRegSearchItem() {
<% if idx <> "" then %>
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/mobile/tenclass/doSubRegItemCdArray.asp?idx=<%=idx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

// ��ǰ�ڵ� �ϰ� ���
function popRegArrayItem() {
<% if idx<>"" then %>
    var popwin = window.open('popSubRegItemCdArray.asp?idx=<%=idx%>','popRegArray','width=600,height=300,scrollbars=yes,resizable=yes');
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

//'������ ����
function itemdel(v){
	if (confirm("��ǰ�� �����˴ϴ� ���� �Ͻðڽ��ϱ�?")){
		document.frmdel.chkIdx.value = v;
		document.frmdel.mode.value = "itemdel";
		document.frmdel.action="doListModify.asp";
		document.frmdel.submit();
	}
}
</script>
<form name="frmdel" method="POST" action="">
<input type="hidden" name="mode" />
<input type="hidden" name="chkIdx" />
</form>
<form name="frm" method="post" action="<%=uploadUrl%>/linkweb/mobile/tenclass_proc.asp" enctype="multipart/form-data" style="margin:0px;">
<table width="900" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<input type="hidden" name="mode" value="<%=mode%>" />
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>" />
<input type="hidden" name="idx" value="<%=idx%>" />
<input type="hidden" name="prevDate" value="<%=prevDate%>" />
<input type="hidden" name="menupos" value="<%=menupos%>" />
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
	<td bgcolor="#FFF999" align="center" width="10%">���� �̹���</td>
	<td>
		<input type="file" name="mainimage" class="file" title="�̺�Ʈ #1" require="N" style="width:50%;" />
		<% if mainimage<>"" then %>
		<br>
		<img src="<%= mainimage %>" width="200" /><br><%= mainimage %>
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">����ī��</td>
    <td>
		<input type="text" name="maincopy" size="50" value="<%=maincopy%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#FFF999" align="center" width="15%">����ī��</td>
    <td>
		<input type="text" name="subcopy" size="80" value="<%=subcopy%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�۾��� ���û���</td>
	<td colspan="3"><textarea name="adminnotice" cols="80" rows="8"/><%=adminnotice%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="1" <%=chkiif(isusing = "1","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="0"  <%=chkiif(isusing = "0","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2">
		<input type="button" value=" �� �� " onClick="jsgolist();"/><input type="button" value=" �� �� " onClick="jsSubmit();"/>
	</td>
</tr>
</table>
</form>

<%
	If idx <> "" then
%>
<p><b>�� Ŭ���� ����</b></p>
<!-- // ��ϵ� ���� ��� --------->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="900" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td colspan="8">
			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
			<tr>
				<td align="left">
					�� <%=oSubItemList.FTotalCount%> �� /
					<input type="button" value="��ü����" class="button" onClick="chkAllItem()">
					<input type="button" value="��������" class="button" onClick="saveList()" title="ǥ�ü��� �� ��뿩�θ� �ϰ������մϴ�.">
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<col width="30" />
	<col width="30" />
	<col width="30" />
	<col width="150" />
	<col width="30" />
	<col width="80" />
	<col width="30" />
	<tr align="center" bgcolor="#DDDDFF">
		<td>&nbsp;</td>
		<td>�̹���</td>
		<td>��ǰ�ڵ�</td>
		<td>��ǰ��</td>
		<td>ǥ�ü���</td>
		<td>��뿩��</td>
		<td>��ǰ����</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="8">
			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
			<tr>
				<td align="right">
					<input type="button" value="��ǰ�ڵ�� ���" class="button" onClick="popRegArrayItem()" />
					<input type="button" value="��ǰ �߰�" class="button" onClick="popRegSearchItem()" />
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tbody id="subList">
	<% For lp=0 to oSubItemList.FResultCount-1 %>
	<tr align="center" bgcolor="<%=chkIIF(oSubItemList.FItemList(lp).FIsUsing,"#FFFFFF","#F3F3F3")%>#FFFFFF">
		<td><input type="checkbox" name="chkIdx" value="<%=oSubItemList.FItemList(lp).Fdidx%>" /></td>
		<td>
		<%
			if Not(oSubItemList.FItemList(lp).FsmallImage="" or isNull(oSubItemList.FItemList(lp).FsmallImage)) then
				Response.Write "<img src='" & oSubItemList.FItemList(lp).FsmallImage & "' height='50' />"
			end if
		%>
		</td>
		<td>
		<%
			if Not(oSubItemList.FItemList(lp).FItemid="0" or isNull(oSubItemList.FItemList(lp).FItemid) or oSubItemList.FItemList(lp).FItemid="") then
				Response.Write "<input type='text' value='" & oSubItemList.FItemList(lp).FItemid & "' readonly size='5'/>"
			end if
		%>
		</td>
		<td><input type="text" name="itemname<%=oSubItemList.FItemList(lp).Fdidx%>" value="<%=oSubItemList.FItemList(lp).Fitemname%>" size="40"></td>
		<td><input type="text" name="sort<%=oSubItemList.FItemList(lp).Fdidx%>" size="3" class="text" value="<%=oSubItemList.FItemList(lp).Fsortno%>" style="text-align:center;" /></td>
		<td>
			<span class="rdoUsing">
			<input type="radio" name="use<%=oSubItemList.FItemList(lp).Fdidx%>" id="rdoUsing<%=lp%>_1" value="1" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing,"checked","")%> /><label for="rdoUsing<%=lp%>_1">���</label>
			<input type="radio" name="use<%=oSubItemList.FItemList(lp).Fdidx%>" id="rdoUsing<%=lp%>_2" value="0" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing,"","checked")%> /><label for="rdoUsing<%=lp%>_2">����</label>
			</span>
		</td>
		<td><input type="button" value="��ǰ����" onclick="itemdel('<%=oSubItemList.FItemList(lp).Fdidx%>');"/></td>
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
