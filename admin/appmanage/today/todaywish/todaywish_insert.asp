<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : todaywish_insert.asp
' Discription : ����� todaywish
' History : 2014.01.28 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/todaywishCls.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , isusing , mode
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate , lp
Dim sDt, sTm, eDt, eTm , gubun , wishtitle , prevDate
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
	dim todaywishList
	set todaywishList = new Ctodaywish
	todaywishList.FRectIdx = idx
	todaywishList.GetOneContents()

	gubun				=	todaywishList.FOneItem.Fgubun
	wishtitle			=	todaywishList.FOneItem.Fwishtitle
	mainStartDate		=	todaywishList.FOneItem.Fstartdate
	mainEndDate			=	todaywishList.FOneItem.Fenddate 
	isusing				=	todaywishList.FOneItem.Fisusing

	set todaywishList = Nothing
End If 

Dim oSubItemList
set oSubItemList = new Ctodaywish
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
		self.location.href="/admin/appmanage/today/todaywish/";
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
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/appmanage/today/todaywish/doSubRegItemCdArray.asp?listidx=<%=idx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
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
<form name="frm" method="post" action="dotodaywish.asp">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
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
    <td bgcolor="#FFF999" align="center" width="15%">����</td>
    <td>
		<input type="text" name="wishtitle" size="100" value="<%If idx="" then%><%=sDt%>&nbsp;���� todaywish �Դϴ�<% Else %><%=wishtitle%><% End if%>" />
    </td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2"><input type="button" value=" �� �� " onClick="jsgolist();"/><input type="button" value=" �� �� " onClick="jsSubmit();"/></td>
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
		    	<input type="button" value="��ü����" class="button" onClick="chkAllItem()">
		    	<input type="button" value="��������" class="button" onClick="saveList()" title="ǥ�ü��� �� ��뿩�θ� �ϰ������մϴ�.">
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
<col width="30" />
<col width="30" />
<col span="3" width="0*" />
<col width="30" />
<col width="30" />
<col width="30" />
<tr align="center" bgcolor="#DDDDFF">
    <td>&nbsp;</td>
    <td>�����ȣ</td>
    <td>�̹���</td>
    <td>��ǰ�ڵ�</td>
    <td>��ǰ��</td>
    <td>����ī��Ʈ</td>
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
    		Response.Write "[" & oSubItemList.FItemList(lp).FItemid & "]"
    	end if
    %>
    </td>
	<td><input type="text" name="itemname<%=oSubItemList.FItemList(lp).FsubIdx%>" value="<%=oSubItemList.FItemList(lp).Fitemname%>"/ size="60"></td>
	<td><%=oSubItemList.FItemList(lp).Ffavcount%></td>
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