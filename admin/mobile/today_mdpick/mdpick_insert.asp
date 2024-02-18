<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : mdpick_insert.asp
' Discription : ����� mdpick
' History : 2014.01.28 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/new_mdpickCls.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , isusing , mode
Dim srcSDT , srcEDT
Dim mainStartDate, mainEndDate , lp , ii
Dim sDt, sTm, eDt, eTm , gubun , mdpicktitle , prevDate , topview , userlevelgubun
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
	dim mdpickList
	set mdpickList = new Cmdpick
	mdpickList.FRectIdx = idx
	mdpickList.GetOneContents()

	gubun			=	mdpickList.FOneItem.Fgubun
	mdpicktitle		=	mdpickList.FOneItem.Fmdpicktitle
	mainStartDate	=	mdpickList.FOneItem.Fstartdate
	mainEndDate		=	mdpickList.FOneItem.Fenddate
	isusing			=	mdpickList.FOneItem.Fisusing
	topview			=	mdpickList.FOneItem.Ftopview
	userlevelgubun	=	mdpickList.FOneItem.Fuserlevelgubun

	set mdpickList = Nothing
End If

If Isnull(topview) Or topview = "0" Or  topview = "" Then topview = 4

Dim oSubItemList
set oSubItemList = new Cmdpick
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
		self.location.href="/admin/mobile/today_mdpick/";
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
	$(".rdoIsLowestPrice").buttonset().children().next().attr("style","font-size:11px;");	

	for ( ii = 1 ; ii<=4 ; ii++ )
	{
		$( "#subList"+ii ).sortable({
			placeholder: "ui-state-highlight",
			start: function(event, ui) {
				ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
				$(this).find('.etcInfo').hide();
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
				$(this).find('.etcInfo').show();
			}
		});
	}
});

//����
function popSubEdit(subidx,v,t) {
<% if idx <>"" then %>
    var popwin = window.open('popSubItemEdit.asp?listIdx=<%=idx%>&subIdx='+subidx+'&gubun='+v+'&topview='+t,'popTemplateManage','width=800,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

// ��ǰ�˻� �ϰ� ��� (������)
function popRegSearchItem(v,t) {
<% if idx <> "" then %>
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/mobile/today_mdpick/doSubRegItemCdArray.asp?listidx=<%=idx%>||gubun="+v+'||topview='+t, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

// ��ǰ�˻� �ϰ� ��� (�Ź���)
function popRegMdPickRcmdItem(v,t) {
<% if idx <> "" then %>
    var popwin = window.open("/admin/itemmaster/pop_itemAutoPick.asp?sellyn=Y&usingyn=Y&defaultmargin=0&gubun="+v+"&acURL=/admin/mobile/today_mdpick/doSubRegItemCdArray.asp?listidx=<%=idx%>||gubun="+v+'||topview='+t, "popup_item", "width=1150,height=600,scrollbars=yes,resizable=yes");
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

// ��ǰ�ڵ� �ϰ� ���
function popRegArrayItem(v,t) {
<% if idx<>"" then %>
    var popwin = window.open('popSubRegItemCdArray.asp?listIdx=<%=idx%>&gubun='+v+'&topview='+t,'popRegArray','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

// ��ǰ�ڵ� �ϰ� ���
function popRegArrayDealItem(v,t) {
<% if idx<>"" then %>
    var popwin = window.open('popSubRegItemCdArray.asp?dealyn=Y&listIdx=<%=idx%>&gubun='+v+'&topview='+t,'popRegArray','width=600,height=300,scrollbars=yes,resizable=yes');
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

// ��ǰ�̹��� ����
function editItemImage(itemid) {
	var param = "itemid=" + itemid;
	popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage','width=1000,height=900,scrollbars=yes,resizable=yes');
	popwin.focus();
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

function popCurrentStock(itemid) {
	var popwin;
    var popwin = window.open('/admin/stock/itemcurrentstock.asp?menupos=<%= request("menupos") %>&itemid='+itemid,'popRegArray','width=1024,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>
<form name="frmdel" method="POST" action="">
<input type="hidden" name="mode" />
<input type="hidden" name="chkIdx" />
</form>
<table width="1140" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post" action="domdpick.asp">
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
		<input type="text" name="mdpicktitle" size="100" value="<%If idx="" then%><%=sDt%>&nbsp;���� mdpick �Դϴ�<% Else %><%=mdpicktitle%><% End if%>" />
    </td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">���ݱ���</td>
	<td><div style="float:left;"><input type="radio" name="userlevelgubun" value="1" <%=chkiif(userlevelgubun = "1","checked","")%> checked />���� &nbsp;&nbsp;&nbsp; <input type="radio" name="userlevelgubun" value="2" <%=chkiif(userlevelgubun = "2","checked","")%>/>VIP�̻�</div> <div style="float:right;margin-top:5px;margin-right:10px;"></div></td>
</tr>

<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2">
		<input type="button" value=" �� �� " onClick="jsgolist();"/><input type="button" value=" �� �� " onClick="jsSubmit();"/>
	</td>
</tr>
</form>
</table>

<%
	If idx <> "" then
%>
<p><b>�� ���� ����</b></p>
<!-- // ��ϵ� ���� ��� -->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="1140" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="9">
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
<col width="30" />
<col width="400" />
<col width="40" />
<col width="90" />
<col width="105" />
<col width="30" />
<tr align="center" bgcolor="#DDDDFF">
    <td>&nbsp;</td>
    <td>�����ȣ</td>
    <td>�̹���</td>
    <td>��ǰ�ڵ�</td>
    <td>��ǰ����</td>
    <td>ǥ�ü���</td>
    <td>��뿩��</td>
    <td>������ ����</td>	
    <td>��ǰ����</td>
</tr>
<% For ii = 1 To 4%>
<% If  ii = 1 Then %><tr><td align="center" bgcolor="#F3F3F3" colspan="9">MD`S PICK</td></tr><% End If %>
<% If  ii = 2 Then %><tr><td align="center" bgcolor="#F3F3F3" colspan="9">NEW</td></tr><% End If %>
<% If  ii = 3 Then %><tr><td align="center" bgcolor="#F3F3F3" colspan="9">BEST (2017 �������� ������)</td></tr><% End If %>
<% If  ii = 4 Then %><tr><td align="center" bgcolor="#F3F3F3" colspan="9">SALE</td></tr><% End If %>
<tr bgcolor="#FFFFFF">
	<td colspan="9">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="right">
		    	<input type="button" value="��ǰ�ڵ�� ���" class="button" onClick="popRegArrayItem('<%=ii%>','<%=topview%>')" />
				<input type="button" value="����ǰ ���" class="button" onClick="popRegArrayDealItem('<%=ii%>','<%=topview%>')" />
		    	<input type="button" value="��ǰ �߰�<%=chkiif(ii>1 ,"(��)","")%>" class="button" onClick="popRegSearchItem('<%=ii%>','<%=topview%>')" />
				<% If ii > 1 Then %>
				<input type="button" value="��õ ��ǰ �߰�" class="button" onClick="popRegMdPickRcmdItem('<%=ii%>','<%=topview%>')" />
				<% End If %>
		    	<img src="/images/icon_new_registration.gif" border="0" onclick="popSubEdit('','<%=ii%>','<%=topview%>')" style="cursor:pointer;" align="absmiddle">
		    </td>
		</tr>
		</table>
	</td>
</tr>
<tbody id="subList<%=ii%>">
<% For lp=0 to oSubItemList.FResultCount-1 %>
<% If oSubItemList.FItemList(lp).Fgubun = ii Then %>
<tr align="center" bgcolor="<%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","#FFFFFF","#F3F3F3")%>#FFFFFF">
    <td><input type="checkbox" name="chkIdx" value="<%=oSubItemList.FItemList(lp).FsubIdx%>" /><input type="hidden" name="chkgubun<%=oSubItemList.FItemList(lp).FsubIdx%>" value="<%=oSubItemList.FItemList(lp).Fgubun%>"/></td>
    <td onclick="popSubEdit(<%=oSubItemList.FItemList(lp).FsubIdx%>,'<%=ii%>','<%=topview%>')" style="cursor:pointer;"><%=oSubItemList.FItemList(lp).FsubIdx%></td>
    <td style="cursor:pointer;">	
    <%		
		if oSubItemList.FItemList(lp).FFrontImage <> "" then
			Response.Write "<img src='" & oSubItemList.FItemList(lp).FFrontImage & "' height='50' />"
		else
			response.write "<img src='" & chkiif(oSubItemList.FItemList(lp).FTentenImg = "" or isnull(oSubItemList.FItemList(lp).FTentenImg), oSubItemList.FItemList(lp).FsmallImage , oSubItemList.FItemList(lp).FTentenImg ) & "' height='50' />"
		end if
    %>
	<br/><input type="button" onclick="editItemImage(<%=oSubItemList.FItemList(lp).FItemid%>,'<%=ii%>','<%=topview%>')" value="����"/>
    </td>
    <td>
    <%
    	if Not(oSubItemList.FItemList(lp).FItemid="0" or isNull(oSubItemList.FItemList(lp).FItemid) or oSubItemList.FItemList(lp).FItemid="") then
    		Response.Write "<input type='text' value='" & oSubItemList.FItemList(lp).FItemid & "' readonly size='5'/>"
    	end if
    %>
    </td>
	<td style="text-align:left;">
		<input type="text" name="itemname<%=oSubItemList.FItemList(lp).FsubIdx%>" value="<%=oSubItemList.FItemList(lp).Fitemname%>" size="60">
		<table cellpadding="1" cellspacing="1" class="a etcInfo" style="padding-top:15px;">
			<tr>
				<td>�ǸŰ� :</td>
				<td><%=FormatNumber(oSubItemList.FItemList(lp).Forgprice,0)%> <%=oSubItemList.FItemList(lp).saleCouponPriceCheck(oSubItemList.FItemList(lp).Fsailyn , oSubItemList.FItemList(lp).FitemCouponYn , oSubItemList.FItemList(lp).Forgprice , oSubItemList.FItemList(lp).Fsailprice , oSubItemList.FItemList(lp).FitemCouponType)%></td>
			</tr>
			<tr>
				<td>���� :</td>
				<td><%=fnPercent(oSubItemList.FItemList(lp).Forgsuplycash,oSubItemList.FItemList(lp).Forgprice,1)%> <%=oSubItemList.FItemList(lp).priceMarginCheck( oSubItemList.FItemList(lp).Fsailyn , oSubItemList.FItemList(lp).FitemCouponYn , oSubItemList.FItemList(lp).FitemCouponType , oSubItemList.FItemList(lp).Fsailsuplycash , oSubItemList.FItemList(lp).Fsailprice , oSubItemList.FItemList(lp).Fcouponbuyprice , oSubItemList.FItemList(lp).Fbuycash)%></td>
			</tr>
			<tr>
				<td>��౸�� :</td>
				<td><%=fnColor(oSubItemList.FItemList(lp).Fmwdiv,"mw")%>-<%=oSubItemList.FItemList(lp).deliveryTypeName(oSubItemList.FItemList(lp).Fdeliverytype)%></td>
			</tr>
			<tr>
				<td>�����Ȳ :</td>
				<td><a href="javascript:popCurrentStock('<%= oSubItemList.FItemList(lp).FItemid %>');">[����]</a></td>
			</tr>
		</table>
	</td>
    <td><input type="text" name="sort<%=oSubItemList.FItemList(lp).FsubIdx%>" size="3" class="text" value="<%=oSubItemList.FItemList(lp).Fsortnum%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">���</label><input type="radio" name="use<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oSubItemList.FItemList(lp).FIsUsing="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">����</label>
		</span>
    </td>
    <td>
		<span class="rdoIsLowestPrice">
		<input type="radio" name="lowestPrice<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoIsLowestPrice<%=lp%>_1" value="Y" <%=chkIIF(oSubItemList.FItemList(lp).FLowestPrice="Y","checked","")%> /><label for="rdoIsLowestPrice<%=lp%>_1">���</label><input type="radio" name="lowestPrice<%=oSubItemList.FItemList(lp).FsubIdx%>" id="rdoIsLowestPrice<%=lp%>_2" value="N" <%=chkIIF(oSubItemList.FItemList(lp).FLowestPrice="N" or oSubItemList.FItemList(lp).FLowestPrice="","checked","")%> /><label for="rdoIsLowestPrice<%=lp%>_2">������</label>
		</span>
    </td>	
	<td><input type="button" value="��ǰ����" onclick="itemdel('<%=oSubItemList.FItemList(lp).FsubIdx%>');"/></td>
</tr>
<% End If %>
<% Next %>
</tbody>
<% Next %>
</table>
</form>
<%
	End If
	set oSubItemList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
