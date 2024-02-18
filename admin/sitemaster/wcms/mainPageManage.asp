<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/mainWCMSCls.asp" -->
<%
'###############################################
' PageName : mainPageManage.asp
' Discription : ����Ʈ ���� ���/���� �� ���� ����
' History : 2013.04.01 ������ : �ű� ����
'###############################################

'// ���� ����
Dim siteDiv, pageDiv
Dim oTemplate, oMainCont, oSubList, lp
Dim MainIdx, mainStartDate, mainEndDate, mainTitle, mainTitleYn, mainSortNo, mainTimeYN, mainIcon, mainSubNum, mainExtDataCd
Dim mainIsPreOpen, mainIsUsing, mainRegUserId, mainRegDate, mainLastModiUserid, mainLastModiDate, mainWorkRequest, mainStat
Dim tplIdx, tplType, tplName, isTimeUse, isIconUse, isSubNumUse, isTopImgUse, isTopLinkUse
Dim isImageUse, isTextUse, isLinkUse, isItemUse, isVideoUse, isBGColorUse, isExtDataUse, isImgDescUse, tplinfoDesc, tplSortNo
Dim sDt, sTm, eDt, eTm
Dim srcSDT, srcEDT

'// �Ķ���� ����
siteDiv = request("site")
pageDiv = request("pDiv")
MainIdx = request("MainIdx")
srcSDT = request("sDt")
srcEDT = request("eDt")
mainSortNo = 0
mainSubNum = 1

'// ������ ����
	set oMainCont = new CCMSContent
	oMainCont.FRectMainIdx = MainIdx
    if MainIdx<>"" then
    	oMainCont.GetOneMainPage

		if oMainCont.FResultCount>0 then
			mainIdx = oMainCont.FOneItem.FmainIdx
			tplIdx = oMainCont.FOneItem.FtplIdx
			mainStartDate = oMainCont.FOneItem.FmainStartDate
			mainEndDate = oMainCont.FOneItem.FmainEndDate
			mainTitle = oMainCont.FOneItem.FmainTitle
			mainTitleYn = oMainCont.FOneItem.FmainTitleYn
			mainSortNo = oMainCont.FOneItem.FmainSortNo
			mainTimeYN = oMainCont.FOneItem.FmainTimeYN
			mainIcon = oMainCont.FOneItem.FmainIcon
			mainSubNum = oMainCont.FOneItem.FmainSubNum
			mainExtDataCd = oMainCont.FOneItem.FmainExtDataCd
			mainIsPreOpen = oMainCont.FOneItem.FmainIsPreOpen
			mainIsUsing = oMainCont.FOneItem.FmainIsUsing
			mainRegUserId = oMainCont.FOneItem.FmainRegUserId
			mainRegDate = oMainCont.FOneItem.FmainRegDate
			mainLastModiUserid = oMainCont.FOneItem.FmainLastModiUserid
			mainLastModiDate = oMainCont.FOneItem.FmainLastModiDate
			mainWorkRequest = oMainCont.FOneItem.FmainWorkRequest
			mainStat = oMainCont.FOneItem.FmainStat
		end if
    end if
    set oMainCont = Nothing

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

'// ���ø� ����
if Not(tplIdx="" or isNull(tplIdx)) then
	set oTemplate = new CCMSContent
	oTemplate.FRectTplIdx = tplIdx
    if tplIdx<>"" then
    	oTemplate.GetOneTemplate
		if oTemplate.FResultCount>0 then
			tplType			= oTemplate.FOneItem.FtplType
			tplName			= oTemplate.FOneItem.FtplName
			siteDiv			= oTemplate.FOneItem.FsiteDiv
			isTimeUse		= oTemplate.FOneItem.FisTimeUse
			isIconUse		= oTemplate.FOneItem.FisIconUse
			isSubNumUse		= oTemplate.FOneItem.FisSubNumUse
			isTopImgUse		= oTemplate.FOneItem.FisTopImgUse
			isTopLinkUse	= oTemplate.FOneItem.FisTopLinkUse
			isImageUse		= oTemplate.FOneItem.FisImageUse
			isTextUse		= oTemplate.FOneItem.FisTextUse
			isLinkUse		= oTemplate.FOneItem.FisLinkUse
			isItemUse		= oTemplate.FOneItem.FisItemUse
			isVideoUse		= oTemplate.FOneItem.FisVideoUse
			isBGColorUse	= oTemplate.FOneItem.FisBGColorUse
			isExtDataUse	= oTemplate.FOneItem.FisExtDataUse
			isImgDescUse	= oTemplate.FOneItem.FisImgDescUse
			tplinfoDesc		= oTemplate.FOneItem.FtplinfoDesc
			tplSortNo		= oTemplate.FOneItem.FtplSortNo
		end if
    end if
    set oTemplate = Nothing
end if

'// ���ø� ���
	set oTemplate = new CCMSContent
	oTemplate.FPageSize = 100
	oTemplate.FRectSiteDiv = siteDiv
	oTemplate.FRectPageDiv = pageDiv
    oTemplate.GetTemplateList

'// ���� ���
	set oSubList = new CCMSContent
	oSubList.FPageSize = 100
	oSubList.FRectMainIdx = MainIdx
    if MainIdx<>"" then
    	oSubList.GetMainSubItem
    end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
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
      	<% if MainIdx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
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
      	<% if MainIdx<>"" then %>minDate: "<%=sDt%>",<% end if %>
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });
    $("#rdoTime").buttonset().children().attr("style","font-size:11px;");
    $("#rdoPre").buttonset().children().attr("style","font-size:11px;");
	$("#rdoTitle").buttonset().children().attr("style","font-size:11px;");
    $("#rdoUsing").buttonset().children().attr("style","font-size:11px;");
    $(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	// �� ����
	$( "#subList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="<%=chkIIF(isTopImgUse="Y" or isImageUse="Y" or isItemUse="Y","54","30")%>" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
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

// ���ø� ��ȯ
function chgTemplate() {
	var opt = $("select[name='tplIdx'] option:selected").attr("opt");
	$(".optRow").hide();
	$("#rdoTime").hide();
	if(opt.substr(0,1)=="Y") {
		$("#rowTime").show();
		$("#rdoTime").show();
	}
	if(opt.substr(1,1)=="Y") $("#rowIcon").show();
	if(opt.substr(2,1)=="Y") $("#rowSubNum").show();
	if(opt.substr(3,1)=="Y") $("#rowExt").show();

	if($("#lyTplInfo").css("display")!="none") {
		var tplCd = $("select[name='tplIdx'] option:selected").val();
		if(tplCd!="") {
			$("#tplInfoImg").html("<img src='/images/wcms/tplInfo"+tplCd+".JPG' />");
		} else {
			$("#tplInfoImg").html("������ ���� ���ø��� �������ּ���.");
		}
	}
}

// ���˻�
function SaveTemplate(frm) {
	if(frm.mainTitle.value=="") {
		alert("���������� �Է����ּ���.");
		frm.mainTitle.focus();
		return;
	}
	if(frm.tplIdx.value=="") {
		alert("����� ���ø��� �������ּ���.");
		frm.tplIdx.focus();
		return;
	}
	if(frm.mainSortNo.value=="") {
		alert("������ �켱������ �Է����ּ���.");
		frm.mainSortNo.focus();
		return;
	}

	if($("#rowIcon").css("display")!="none"&&frm.mainSubNum.value<"1") {
		alert("������ �ּ� ��� ������ �Է����ּ���.");
		frm.mainSubNum.focus();
		return;
	}

	if($("#rowExt").css("display")!="none"&&frm.mainExtDataCd.value=="") {
		alert("�ܺ� ������ ������ �������ּ���.");
		frm.mainExtDataCd.focus();
		return;
	}

	frm.submit();
}

// ���� ������ ���/����
function popSubEdit(subidx) {
<% if MainIdx<>"" then %>
    var popwin = window.open('/admin/sitemaster/wcms/popSubItemEdit.asp?mainIdx=<%=mainIdx%>&subIdx='+subidx,'popTemplateManage','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

// ��ǰ�ڵ� �ϰ� ���
function popRegArrayItem() {
<% if MainIdx<>"" then %>
    var popwin = window.open('/admin/sitemaster/wcms/popSubRegItemCdArray.asp?mainIdx=<%=mainIdx%>','popRegArray','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
<% else %>
	alert("���ø� ������ ������ ���� ������ּ���.");
<% end if %>
}

// ��ǰ�˻� �ϰ� ���
function popRegSearchItem() {
<% if MainIdx<>"" then %>
    var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&defaultmargin=0&acURL=/admin/sitemaster/wcms/doSubRegItemCdArray.asp?mainIdx=<%=mainIdx%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
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

// ���ø� ���� On/Off
function fnViewTplInfo() {
	var tplCd = $("select[name='tplIdx']").val();
	if(tplCd=="") {
		alert("������ ���� ���ø��� �������ּ���.");
		return;
	}
	if($("#lyTplInfo").css("display")=="none") {
		$("#lyTplInfo").show();
		$("#tplInfoImg").html("<img src='/images/wcms/tplInfo"+tplCd+".JPG' />");
	} else {
		$("#lyTplInfo").hide();
	}
}
</script>
<!-- ���������� ���� ���� -->
<form name="frm" method="POST" action="doMainContents.asp" style="margin:0;">
<input type="hidden" name="site" value="<%= siteDiv %>" />
<input type="hidden" name="pDiv" value="<%= pageDiv %>" />
<input type="hidden" name="srcSDT" value="<%= srcSDT %>" />
<input type="hidden" name="srcEDT" value="<%= srcEDT %>" />
<input type="hidden" name="menupos" value="<%= request("menupos") %>" />
<p><b>�� ����Ʈ ����</b></p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout: fixed;">
<colgroup>
	<col width="120" />
	<col width="*" />
	<col width="120" />
	<col width="*" />
</colgroup>
<% if MainIdx<>"" then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">������ȣ</td>
    <td colspan="3">
        <%=MainIdx %>
        <input type="hidden" name="MainIdx" value="<%=MainIdx %>" />
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">����Ⱓ</td>
    <td colspan="3">
		<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
		<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
		<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
		<input type="text" name="eTm" size="8" value="<%=eTm%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">������ ����</td>
    <td colspan="3">
		<input type="text" name="mainTitle" size="60" maxlength="128" value="<%=mainTitle%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���ø� ����</td>
    <td colspan="3">
		<select name="tplIdx" class="select" onchange="chgTemplate()">
		<option value="" opt="">::����::</option>
		<%
			if oTemplate.FResultCount>0 then
				for lp=0 to (oTemplate.FResultCount-1)
					Response.Write "<option value='" & oTemplate.FItemList(lp).FtplIdx & "' " & chkIIF(cStr(oTemplate.FItemList(lp).FtplIdx)=cStr(tplIdx),"selected","") &_
							" opt='" & oTemplate.FItemList(lp).FisTimeUse & oTemplate.FItemList(lp).FisIconUse & oTemplate.FItemList(lp).FisSubNumUse & oTemplate.FItemList(lp).FisExtDataUse & "'>" &_
							oTemplate.FItemList(lp).FtplName & "</option>"
				next
			end if
		%>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���� ����</td>
    <td colspan="3">
		<% if (datediff("h",mainEndDate,now)>=0 and MainIdx<>"") or mainStat="9" then %>
		<strong><%=chkIIF(mainStat="9","��������","����")%></strong>
		<input type="hidden" name="mainStat" value="<%=mainStat%>" />
		<% else %>
		<select name="mainStat" class="select">
		<option value="0" <%=chkIIF(mainStat="0","selected","")%>>��ϴ��</option>
		<option value="3" <%=chkIIF(mainStat="3","selected","")%>>�̹�����Ͽ�û</option>
		<option value="5" <%=chkIIF(mainStat="5","selected","")%>>���¿�û(�̹��� ��ϿϷ�)</option>
		<option value="7" <%=chkIIF(mainStat="7","selected","")%>>����</option>
		<option value="9" <%=chkIIF(mainStat="9","selected","")%>>��������</option>
		</select>
		<% end if %>
    </td>
</tr>
<tr id="rowTime" class="optRow" bgcolor="#FFFFFF" style="display:<%=chkIIF(isTimeUse="Y","","none")%>;">
    <td bgcolor="#DDDDFF">�ð�ǥ�� ����</td>
    <td colspan="3">
		<div id="rdoTime">
		<input type="radio" name="mainTimeYN" id="rdoTm1" value="Y" <%=chkIIF(mainTimeYN="Y","checked","")%> /><label for="rdoTm1">ǥ��</label><input type="radio" name="mainTimeYN" id="rdoTm2" value="N" <%=chkIIF(mainTimeYN="N" or mainTimeYN="","checked","")%> /><label for="rdoTm2">����</label>
		</div>
    </td>
</tr>
<tr id="rowIcon" class="optRow" bgcolor="#FFFFFF" style="display:<%=chkIIF(isIconUse="Y","","none")%>;">
    <td bgcolor="#DDDDFF">������ ����</td>
    <td colspan="3">
		<select name="mainIcon" class="select">
		<option value="">::����::</option>
		<option value="I" <%=chkIIF(mainIcon="I","selected","")%>>��ǰ���� ������</option>
		<option value="T" <%=chkIIF(mainIcon="T","selected","")%>>Today Hot</option>
		</select>
    </td>
</tr>
<tr id="rowSubNum" class="optRow" bgcolor="#FFFFFF" style="display:<%=chkIIF(isSubNumUse="Y","","none")%>;">
    <td bgcolor="#DDDDFF">�׸񰳼�</td>
    <td colspan="3">
		<input type="text" name="mainSubNum" size="4" value="<%=mainSubNum%>" class="text" />
		(�� �ּ� ��ǰ ����)
    </td>
</tr>
<tr id="rowExt" class="optRow" bgcolor="#FFFFFF" style="display:<%=chkIIF(isExtDataUse="Y","","none")%>;">
    <td bgcolor="#DDDDFF">�ܺ� �ڷ� ���</td>
    <td colspan="3">
		<select name="mainExtDataCd" class="select">
		<option value="">::����::</option>
		<option value="BA" <%=chkIIF(mainExtDataCd="BA","selected","")%>>����Ʈ �����</option>
		<option value="DF" <%=chkIIF(mainExtDataCd="DF","selected","")%>>������ �ΰŽ�</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" title="������ �� ����Ʈ ���⿩��">�����⿩��</td>
    <td>
		<div id="rdoPre">
		<input type="radio" name="mainIsPreOpen" id="rdoPre1" value="Y" <%=chkIIF(mainIsPreOpen="Y","checked","")%> /><label for="rdoPre1">����</label><input type="radio" name="mainIsPreOpen" id="rdoPre2" value="N" <%=chkIIF(mainIsPreOpen="N" or mainIsPreOpen="","checked","")%> /><label for="rdoPre2">�������</label>
		</div>
    </td>
    <td bgcolor="#DDDDFF">����ǥ��</td>
    <td>
		<div id="rdoTitle">
		<input type="radio" name="mainTitleYn" id="rdoTitle1" value="Y" <%=chkIIF(mainTitleYn="Y" or mainTitleYn="","checked","")%> /><label for="rdoTitle1">ǥ��</label><input type="radio" name="mainTitleYn" id="rdoTitle2" value="N" <%=chkIIF(mainTitleYn="N","checked","")%> /><label for="rdoTitle2">ǥ�þ���</label>
		</div>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�켱����</td>
    <td>
		<input type="text" name="mainSortNo" size="4" value="<%=mainSortNo%>" class="text" />
    </td>
    <td bgcolor="#DDDDFF">��뿩��</td>
    <td>
		<div id="rdoUsing">
		<input type="radio" name="mainIsUsing" id="rdoUsg1" value="Y" <%=chkIIF(mainIsUsing="Y" or mainIsUsing="","checked","")%> /><label for="rdoUsg1">���</label><input type="radio" name="mainIsUsing" id="rdoUsg2" value="N" <%=chkIIF(mainIsUsing="N","checked","")%> /><label for="rdoUsg2">������</label>
		</div>
    </td>
</tr>
<% if MainIdx<>"" then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�����</td>
    <td><%=getStaffUserName(mainRegUserId)%></td>
    <td bgcolor="#DDDDFF">�����</td>
    <td><%=mainRegDate%></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�۾���</td>
    <td><%=getStaffUserName(mainLastModiUserid)%></td>
    <td bgcolor="#DDDDFF">�۾���</td>
    <td><%=mainLastModiDate%></td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">�۾���û����</td>
    <td colspan="3">
        <textarea name="mainWorkRequest" class="textarea" style="width:100%; height:160px;"><%=mainWorkRequest%></textarea>
    </td>
</tr>
<tr bgcolor="#F8F8F8">
    <td colspan="4" align="center">
    	<table width="100%" cellpadding="0" cellspacing="0" border="0">
    	<tr>
    		<td width="80" align="left"><input type="button" value=" ���� " onClick="fnViewTplInfo()" class="button" style="background-color:#FFDDCC"></td>
    		<td align="center"><input type="button" value=" �� �� " onClick="SaveTemplate(this.form);" class="button"></td>
    		<td width="80" align="left"><input type="button" value=" ��� " onClick="location.href='/admin/sitemaster/wcms/index.asp?site=<%=siteDiv%>&pDiv=<%=pageDiv%>&menupos=<%= request("menupos") %>&sDt=<%=srcSDT%>&eDt=<%=srcEDT%>'" class="button"></td>
    	</tr>
    	</table>
    </td>
</tr>
<tr id="lyTplInfo" bgcolor="#FFFFFF" style="display:none;">
    <td bgcolor="#DDDDFF">���ø� �ȳ�</td>
    <td colspan="3" id="tplInfoImg" align="center"></td>
</tr>
</table>
</form>
<p><b>�� ���� ����</b></p>
<!-- // ��ϵ� ���� ��� --------->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="10">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr>
		    <td align="left">
		    	�� <%=oSubList.FTotalCount%> �� /
		    	<input type="button" value="��ü����" class="button" onClick="chkAllItem()">
		    	<input type="button" value="��������" class="button" onClick="saveList()" title="ǥ�ü��� �� ��뿩�θ� �ϰ������մϴ�.">
		    </td>
		    <td align="right">
		    	<% if isItemUse="Y" then %>
		    	<!-- ��ǰ�ϰ� ��Ϲ�ư (��ǰ���Ӽ��϶�) -->
		    	<input type="button" value="��ǰ�ڵ�� ���" class="button" onClick="popRegArrayItem()" />
		    	<input type="button" value="��ǰ �߰�" class="button" onClick="popRegSearchItem()" />
		    	<% end if %>
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
    <td>�ؽ�Ʈ</td>
    <td>��ǰ�ڵ�</td>
    <td>��ũ</td>
    <td>ǥ�ü���</td>
    <td>��뿩��</td>
    <td>�����</td>
    <td>�����</td>
</tr>
<tbody id="subList">
<%	For lp=0 to oSubList.FResultCount-1 %>
<tr align="center" bgcolor="<%=chkIIF(oSubList.FItemList(lp).FsubIsUsing="Y","#FFFFFF","#F3F3F3")%>">
    <td><input type="checkbox" name="chkIdx" value="<%=oSubList.FItemList(lp).FsubIdx%>" /></td>
    <td onclick="popSubEdit(<%=oSubList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubList.FItemList(lp).FsubIdx%></td>
    <td onclick="popSubEdit(<%=oSubList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;">
    <%
    	if Not(oSubList.FItemList(lp).FsubImage1="" or isNull(oSubList.FItemList(lp).FsubImage1)) then
    		Response.Write "<img src='" & oSubList.FItemList(lp).getImageUrl(1) & "' height='50' />"
    	end if
    	if Not(oSubList.FItemList(lp).FsmallImage="" or isNull(oSubList.FItemList(lp).FsmallImage)) then
    		Response.Write "<img src='" & oSubList.FItemList(lp).FsmallImage & "' height='50' />"
    	end if
    %>
    </td>
    <td onclick="popSubEdit(<%=oSubList.FItemList(lp).FsubIdx%>)" style="cursor:pointer;"><%=oSubList.FItemList(lp).FsubText1%></td>
    <td>
    <%
    	if Not(oSubList.FItemList(lp).FsubItemid="0" or isNull(oSubList.FItemList(lp).FsubItemid) or oSubList.FItemList(lp).FsubItemid="") then
    		Response.Write "[" & oSubList.FItemList(lp).FsubItemid & "]" & oSubList.FItemList(lp).Fitemname
    	end if
    %>
    </td>
    <td><%=oSubList.FItemList(lp).FsubLinkUrl%></td>
    <td><input type="text" name="sort<%=oSubList.FItemList(lp).FsubIdx%>" size="3" class="text" value="<%=oSubList.FItemList(lp).FsubSortNo%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oSubList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oSubList.FItemList(lp).FsubIsUsing="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">���</label><input type="radio" name="use<%=oSubList.FItemList(lp).FsubIdx%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oSubList.FItemList(lp).FsubIsUsing="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">����</label>
		</span>
    </td>
    <td><%=oSubList.FItemList(lp).FsubRegUsername%></td>
    <td><%=left(oSubList.FItemList(lp).FsubRegdate,10)%></td>
</tr>
<%	Next %>
</tbody>
</table>
</form>
<%
	set oTemplate = Nothing
	set oSubList = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->