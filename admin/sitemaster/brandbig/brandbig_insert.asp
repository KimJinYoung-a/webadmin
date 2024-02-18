<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/brandbigCls.asp" -->
<%
Dim idx , subImage1 , subImage2 , subImage3 , isusing , mode, paramisusing
Dim srcSDT , srcEDT 
Dim mainStartDate, mainEndDate , lp
Dim sDt, sTm, eDt, eTm , gubun , title , prevDate , is1day
Dim extraurl
Dim subtitle , saleper, makerid
Dim bannerImg, linkUrl, altName, bannerNameEng, bannerNameKor, subCopy, sortnum

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
	dim brandbigList
	set brandbigList = new Cbrandbig
	brandbigList.FRectIdx = idx
	brandbigList.GetOneContents()

	mainStartDate	=	brandbigList.FOneItem.Fstartdate
	mainEndDate		=	brandbigList.FOneItem.Fenddate
	isusing			=	brandbigList.FOneItem.Fisusing
	bannerImg		=	brandbigList.FOneItem.FbannerImg
	linkUrl			=	brandbigList.FOneItem.FlinkUrl
	altName			=	brandbigList.FOneItem.FaltName
	bannerNameEng	=	brandbigList.FOneItem.FbannerNameEng
	bannerNameKor	=	brandbigList.FOneItem.FbannerNameKor
	subCopy			=	brandbigList.FOneItem.FsubCopy
	makerid	=	brandbigList.FOneItem.Fmakerid
	subcopy = Replace(subcopy,"<br>",Chr(13)&Chr(10))

	sortnum			=	brandbigList.FOneItem.Fsortnum

	set brandbigList = Nothing
End If 

Dim oSubItemList
set oSubItemList = new CbrandBig
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

		if (!frm.linkurl.value){
			alert("��� ��ũ�� �Է����ּ���.");
			frm.linkurl.focus();
			return;
		}

		if (frm.linkurl.value.indexOf("�귣����̵�") > 0){
			alert("��� ��ũ ���� Ȯ�� ���ּ���.");
			frm.linkurl.focus();
			return;
		}

		if (!frm.altname.value){
			alert("Alt���� �Է����ּ���.");
			frm.altname.focus();
			return;
		}

		if (!frm.bannernameeng.value){
			alert("�귣�� �������� �Է����ּ���.");
			frm.bannernameeng.focus();
			return;
		}

		if (!frm.bannernamekor.value){
			alert("�귣�� �ѱ۸��� �Է����ּ���.");
			frm.bannernamekor.focus();
			return;
		}

		if (!frm.subcopy.value){
			alert("����ī�Ǹ� �Է����ּ���.");
			frm.subcopy.focus();
			return;
		}

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
		self.location.href="/admin/sitemaster/brandbig/?menupos=<%=request("menupos")%>&isusing=<%=paramisusing%>";
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
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/sitemaster/brandbig/doSubRegItemCdArray.asp?listidx=<%=idx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
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
		case 'brand':
			frm.linkurl.value='/street/street_brand_sub06.asp?makerid=�귣����̵�';
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

//�귣�� ID �˻� �˾�â
function jsSearchBrandIDNew(frmName,compName){
	var compVal = "";
	try{
		compVal = eval("document.all." + frmName + "." + compName).value;
	}catch(e){
		compVal = "";
	}

	var popwin = window.open("popBrandSearch.asp?frmName=" + frmName + "&compName=" + compName + "&rect=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

	popwin.focus();
}

function fnRegItemAuto(){
	document.frmList.mode.value="auto";
	document.frmList.action="doSubRegItemCdArray.asp";
	document.frmList.submit();
}
</script>
<form name="frm" method="post" action="dobrandbig.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="prevDate" value="<%=prevDate%>">
<input type="hidden" name="paramisusing" value="<%=paramisusing%>">
<input type="hidden" name="bannerImg" value="<%=bannerImg%>">
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
			<input type="text" id="sDt" name="StartDate" size="10" value="<%=chkiif(mode="add",prevDate,sDt)%>" />
			<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
			<input type="text" id="eDt" name="EndDate" size="10" value="<%=chkiif(mode="add",prevDate,eDt)%>" />
			<input type="text" name="eTm" size="8" value="<%=eTm%>" />
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" id="tmpstyle1">
		<td bgcolor="#DDDDFF" align="center" width="15%" id="imagetitle">�귣�� ��� �̹���</td>
		<td><input type="button" name="wimg" value="�귣�� ��� �̹��� ���" onClick="jsSetImg('brandbig','<%=bannerImg%>','bannerImg','brandbigimg')" class="button">
			<div id="brandbigimg" style="padding: 5 5 5 5">
				<%IF bannerImg <> "" THEN %>
				<a href="javascript:jsImgView('<%=bannerImg%>')"><img  src="<%=bannerImg%>" width="400" border="0"></a>
				<a href="javascript:jsDelImg('bannerImg','brandbigimg');"><img src="/images/icon_delete2.gif" border="0"></a>
				<%END IF%>
			</div>
			<%=bannerImg%>
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF" id="tmpstyle2">
		<td bgcolor="#DDDDFF"  align="center" width="15%">�귣�� ��� ��ũ</td>
		<td>
			<input type="text" name="linkurl" value="<%=linkurl%>" style="width:80%;" /><br/>
			- <span style="cursor:pointer" onClick="putLinkText('brand')">�귣�� ��ũ : /street/street_brand_sub06.asp?makerid=<font color="darkred">�귣����̵�</font></span><br>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999" align="center" width="15%">Alt ��</td>
		<td>
			<input type="text" name="altname" size="50" value="<%=altname%>" />
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999" align="center" width="15%">�귣�弱��</td>
		<td>
			<%	NewDrawSelectBoxDesignerwithNameEvent "makerid", makerid %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999" align="center" width="15%">�귣��� ����</td>
		<td>
			<input type="text" name="bannernameeng" size="50" value="<%=bannernameeng%>" />
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999" align="center" width="15%">�귣��� �ѱ�</td>
		<td>
			<input type="text" name="bannernamekor" size="50" value="<%=bannernamekor%>" />
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999" align="center" width="15%">����ī��</td>
		<td>
			<textarea name="subcopy" cols="40" rows="5"><%=subcopy%></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="#FFF999" align="center" width="15%">�켱����</td>
		<td>
			<input type="text" name="sortnum" size="10" value="<% If sortnum="" Then  Response.write 0 Else Response.write sortnum End If %>" />
		</td>
	</tr>
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
<input type="hidden" name="makerid" value="<%=makerid%>">
<input type="hidden" name="listidx" value="<%=idx%>">
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
				<input type="button" value="��ǰ�ڵ��߰�" class="button" onClick="fnRegItemAuto()" />
		    	<input type="button" value="��ǰ �߰�" class="button" onClick="popRegSearchItem()" />
		    	<!--<img src="/images/icon_new_registration.gif" border="0" onclick="popSubEdit('')" style="cursor:pointer;" align="absmiddle">//-->
		    </td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
    <td>&nbsp;</td>
    <td>�����ȣ</td>
    <td>�̹���</td>
    <td>��ǰ�ڵ�</td>
    <td>��ǰ��</td>
	<td>����</td>
    <td>���Ĺ�ȣ</td>
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
		<%=oSubItemList.FItemList(lp).Fitemname%>
	</td>
	<td align="left" style="padding-left:5px;">
		<%
			Response.Write "<a href=""javascript:editItemPriceInfo('" & oSubItemList.FItemList(lp).Fitemid & "')"" title='�ǸŰ� �� ���ް� ����'>" & FormatNumber(oSubItemList.FItemList(lp).Forgprice,0) & "</a>"
			'���ΰ�
			if oSubItemList.FItemList(lp).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>("&CLng((oSubItemList.FItemList(lp).Forgprice-oSubItemList.FItemList(lp).Fsailprice)/oSubItemList.FItemList(lp).Forgprice*100) & "%��)" & FormatNumber(oSubItemList.FItemList(lp).Fsailprice,0) & "</font>"
			end if
			'������
			if oSubItemList.FItemList(lp).FitemCouponYn="Y" then
				Select Case oSubItemList.FItemList(lp).FitemCouponType
					Case "1"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oSubItemList.FItemList(lp).GetCouponAssignPrice(),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oSubItemList.FItemList(lp).GetCouponAssignPrice(),0) & "</font>"
				end Select
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