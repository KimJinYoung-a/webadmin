<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/mainWCMSCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : ����Ʈ ���� ����
' History : 2013.03.28 ������ : �ű� ����
'###############################################

'// ���� ����
Dim siteDiv, pageDiv, isusing, tplIdx, selDt, sDt, eDt 
Dim oTemplate, oMainCont, lp, modiTime
Dim page

'// �Ķ���� ����
siteDiv = request("site")
pageDiv = request("pDiv")
isusing = request("isusing")
tplIdx = request("tplIdx")
sDt = request("sDt")
eDt = request("eDt")
page = request("page")
if siteDiv="" then siteDiv="P"		'�⺻�� PC��(P:PC��, M:�����)
if pageDiv="" then pageDiv="10"		'�⺻�� ����Ʈ����(10:����Ʈ����, 20:�̺�Ʈ����...)
if isusing="" then isusing="Y"
if sDt="" then sDt=cStr(date)
if eDt="" then eDt=cStr(date)
if sDt=eDt then selDt=sDt
if page="" then page="1"


'// ���ø� ���
	set oTemplate = new CCMSContent
	oTemplate.FPageSize = 50
	oTemplate.FRectSiteDiv = siteDiv
	oTemplate.FRectPageDiv = pageDiv
    oTemplate.GetTemplateList

'// ���������� ���
	set oMainCont = new CCMSContent
	oMainCont.FPageSize = 20
	oMainCont.FRectTplIdx = tplIdx
	oMainCont.FRectSiteDiv = siteDiv
	oMainCont.FRectPageDiv = pageDiv
	oMainCont.FRectIsUsing = isusing
	oMainCont.FRectStartDate = sDt
	oMainCont.FRectEndDate = eDt
    oMainCont.GetMainPageList
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
$(function() {
  	$("input[type=submit]").button();

  	// ������ư
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	// Ķ����
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
    	onClose: function() {
    		if($("#sDt").datepicker("getDate")>$("#eDt").datepicker("getDate")) {
    			$("#eDt").datepicker("setDate",$("#sDt").datepicker("getDate"));
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
    	onClose: function() {
    		if($("#eDt").datepicker("getDate")<$("#sDt").datepicker("getDate")) {
    			$("#sDt").datepicker("setDate",$("#eDt").datepicker("getDate"));
    		}
    	}
    });

	// �� ����
	$( "#mainList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="30" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
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

function popTemplateManage(){
    var popwin = window.open('/admin/sitemaster/wcms/popTemplateEdit.asp?site=<%=siteDiv%>&pDiv=<%=pageDiv%>','popTemplateManage','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function goMainContent(idx) {
	location.href="/admin/sitemaster/wcms/mainPageManage.asp?mainIdx="+idx+"&site=<%=siteDiv%>&pDiv=<%=pageDiv%>&menupos=<%=request("menupos")%>&sDt=<%=sDt%>&eDt=<%=eDt%>";
}

function goPage(page){
    document.frm.page.value = page;
    document.frm.submit();
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
		alert("�����Ͻ� ���ø��� �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.target="_self";
		document.frmList.action="doListModify.asp";
		document.frmList.submit();
	}
}


//����Ʈ �̸�����
function previewPage() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�̸����� ���ø��� �������ּ���.");
		return;
	}

	if($("select[name='site'], input[name='site']").val()=="M") {
		var url;
		switch($("select[name='pDiv']").val()) {
			case "10" :
				url = "<%=mobileUrl%>/chtml/preview/previewMainIndex.asp?sDt=<%=sDt%>";
				break;
			case "20" :
				url = "<%=mobileUrl%>/chtml/preview/previewEventBanner.asp?sDt=<%=sDt%>";
				break;
		}

		 if(confirm("[<%=sDt%>]���� ���갣�� �̸����⸦ �Ͻðڽ��ϱ�?")) {
			 var popwin = window.open('','refreshFrm_Main','width=350,height=600,scrollbars=yes,resizable=yes');
			 popwin.focus();
			 frmList.target = "refreshFrm_Main";
			 frmList.action = url;
			 frmList.submit();
		}
	} else {
		alert("PC���� ���� �غ����Դϴ�.\n���� ������������ ������ּ���.");
		return;
	}
}


//����Ʈ �������� ���� ����
function assignPage() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� ���ø��� �������ּ���.");
		return;
	}

	if($("select[name='site'], input[name='site']").val()=="M") {
		var msg, url;
		switch($("select[name='pDiv']").val()) {
			case "10" :
				msg = "�����Ͻ� �׸��� \"�����\" ����Ʈ ���� �������� ����Ʈ�� ��� �����Ͻðڽ��ϱ�?\n\n���� �� ����� ������ �ǵ��� �� �����ϴ�.";
				url = "<%=mobileUrl%>/chtml/make_main_xml.asp";
				break;
			case "20" :
				msg = "�����Ͻ� �׸��� \"�����\" �̺�Ʈ ���� �������� ����Ʈ�� ��� �����Ͻðڽ��ϱ�?\n\n�ؿ��ú��� 5�Ϻ��� ����Ǹ�, �� �� ����� ������ �ǵ��� �� �����ϴ�.";
				url = "<%=mobileUrl%>/chtml/make_event_xml.asp";
				break;
		}

		if(confirm(msg)) {
			 var popwin = window.open('','refreshFrm_Main','');
			 popwin.focus();

			if($("input[name='cTrm']").val()!="0") {
				frmList.sTrm.value = $("input[name='cTrm']").val();
			}
			 frmList.chkAll.value="N";
			 frmList.target = "refreshFrm_Main";
			 frmList.action = url;
			 frmList.submit();
		}
	} else {
		if(confirm("���� \"PC��\" ����Ʈ�������� ��� �����Ͻðڽ��ϱ�?\n\n���� �� ����� ������ �ǵ��� �� �����ϴ�.")) {
			alert("PC���� ���� �غ����Դϴ�.\n���� ������������ ������ּ���.");
		}
	}
}

//����Ʈ �������� ��ü ����
function assignPageALL() {
	if($("select[name='site'], input[name='site']").val()=="M") {
		var msg, url;
		switch($("select[name='pDiv']").val()) {
			case "10" :
				msg = "������ 4�ϰ��� ��ü �׸��� \"�����\" ����Ʈ ���� �������� ����Ʈ�� ��� �����Ͻðڽ��ϱ�?\n\n���� �� ����� ������ �ǵ��� �� �����ϴ�.";
				url = "<%=mobileUrl%>/chtml/make_main_xml.asp";
				break;
			case "20" :
				msg = "������ 5�ϰ��� ��ü �׸��� \"�����\" �̺�Ʈ ���� �������� ����Ʈ�� ��� �����Ͻðڽ��ϱ�?\n\n���� �� ����� ������ �ǵ��� �� �����ϴ�.";
				url = "<%=mobileUrl%>/chtml/make_event_xml.asp";
				break;
		}

		if(confirm(msg)) {
			 var popwin = window.open('','refreshFrm_Main','');
			 popwin.focus();

			if($("input[name='cTrm']").val()!="0") {
				frmList.sTrm.value = $("input[name='cTrm']").val();
			}
			 frmList.chkAll.value="Y";
			 frmList.target = "refreshFrm_Main";
			 frmList.action = url;
			 frmList.submit();
		}
	} else {
		if(confirm("���� \"PC��\" ����Ʈ�������� ��� �����Ͻðڽ��ϱ�?\n\n���� �� ����� ������ �ǵ��� �� �����ϴ�.")) {
			alert("PC���� ���� �غ����Դϴ�.\n���� ������������ ������ּ���.");
		}
	}
}

// ���� ���
function fnQuitReg(oTpl) {
	var tplId = oTpl.value;
	var tplNm = oTpl[oTpl.selectedIndex].text;
	var tplDt = document.frm.sDt.value;

	var chk = confirm("���� ["+tplDt+"]�� \""+tplNm+"\"�� ����Ͻðڽ��ϱ�?\n\n�� ���ø� �⺻������ �����Ǹ� ���� ������ �ݵ�� �����ϼžߵ˴ϴ�.");
	if(chk) {
		var frm = document.frmQuitReg;
		frm.tplIdx.value=tplId;
		frm.StartDate.value=tplDt;
		frm.EndDate.value=tplDt;
		frm.mainTitle.value="*** ������� > �������ּ���";
		frm.submit();
	} else {
		return;
	}
}
</script>

<!-- ��� �˻��� ���� -->
<form name="frm" method="get" action="" style="margin:0;">
<input type="hidden" name="page" value="" />
<input type="hidden" name="menupos" value="<%= request("menupos") %>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" rowspan="2" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
	    <% if C_ADMIN_AUTH then %>
	    ����Ʈ:
		<select name="site" class="select">
			<option value="P" <%=chkIIF(siteDiv="P","selected","")%> >PC��</option>
			<option value="M" <%=chkIIF(siteDiv="M","selected","")%> >�����</option>
		</select>
		&nbsp;/&nbsp;
		<% else %>
		<input type="hidden" name="site" value="<%=siteDiv%>" />
		<% end if %>
	    ���ó:
		<select name="pDiv" class="select">
			<option value="10" <%=chkIIF(pageDiv="10","selected","")%> >����Ʈ ����</option>
			<option value="20" <%=chkIIF(pageDiv="20","selected","")%> >�̺�Ʈ ����</option>
		</select>
		&nbsp;/&nbsp;
	    ��뱸��:
		<select name="isusing" class="select">
			<option value="A">��ü</option>
			<option value="Y" <%=chkIIF(isusing="Y","selected","")%> >�����</option>
			<option value="N" <%=chkIIF(isusing="N","selected","")%> >������</option>
		</select>
		&nbsp;/&nbsp;
		���ø�:
		<select name="tplIdx" class="select">
			<option value="">��ü</option>
			<%
				if oTemplate.FResultCount>0 then
					for lp=0 to (oTemplate.FResultCount-1)
						Response.Write "<option value='" & oTemplate.FItemList(lp).FtplIdx & "' " & chkIIF(cStr(oTemplate.FItemList(lp).FtplIdx)=tplIdx,"selected","") & ">" & oTemplate.FItemList(lp).FtplName & "</option>"
					next
				end if
			%>
		</select>
	<td width="80" rowspan="2" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" value="�˻�" />
	</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
	<td>
		��ȸ�Ⱓ:
		<span id="rdoDtPreset">
		<input type="radio" name="selDatePreset" id="rdoDtOpt1" value="<%=dateadd("d",-1,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",-1,date)),"checked","")%> /><label for="rdoDtOpt1">-1</label><input type="radio" name="selDatePreset" id="rdoDtOpt2" value="<%=date%>" <%=chkIIF(selDt=cStr(date),"checked","")%> /><label for="rdoDtOpt2">����</label><input type="radio" name="selDatePreset" id="rdoDtOpt3" value="<%=dateadd("d",1,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",1,date)),"checked","")%> /><label for="rdoDtOpt3">+1</label><input type="radio" name="selDatePreset" id="rdoDtOpt4" value="<%=dateadd("d",2,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",2,date)),"checked","")%> /><label for="rdoDtOpt4">+2</label><input type="radio" name="selDatePreset" id="rdoDtOpt5" value="<%=dateadd("d",3,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",3,date)),"checked","")%> /><label for="rdoDtOpt5">+3</label><input type="radio" name="selDatePreset" id="rdoDtOpt6" value="<%=dateadd("d",4,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",4,date)),"checked","")%> /><label for="rdoDtOpt6">+4</label><input type="radio" name="selDatePreset" id="rdoDtOpt7" value="<%=dateadd("d",5,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",5,date)),"checked","")%> /><label for="rdoDtOpt7">+5</label><input type="radio" name="selDatePreset" id="rdoDtOpt8" value="<%=dateadd("d",6,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",6,date)),"checked","")%> /><label for="rdoDtOpt8">+6</label><input type="radio" name="selDatePreset" id="rdoDtOpt9" value="<%=dateadd("d",7,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",7,date)),"checked","")%> /><label for="rdoDtOpt9">+7</label>
		</span>
		<input type="text" id="sDt" name="sDt" size="10" value="<%=sDt%>" style="height:22px;" /> ~
		<input type="text" id="eDt" name="eDt" size="10" value="<%=eDt%>" style="height:22px;" />
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
    <td align="left">
    	<input type="button" value="��ü����" class="button" onClick="chkAllItem()">
    	<input type="button" value="��������" class="button" onClick="saveList()" title="�켱���� �� �����⿩�θ� �ϰ������մϴ�.">
    	/
		<input type="button" value="�̸�����" class="button" onClick="previewPage()" style="background-color:#F0FFF0" title="���� ����Ʈ�������� �̸����ϴ�.">
    	/
    	<% if C_ADMIN_AUTH then %>
		<input type="text" class="text" name="cTrm" value="0" size="1" style="text-align:center;" title="�������� ��¥�� �����մϴ�.(�����ڿ�)">
		<% else %>
		<input type="hidden" name="cTrm" value="0">
		<% end if %>
    	<% if siteDiv="M" and pageDiv="10" then %><input type="button" value="���� ����" class="button" onClick="assignPage()" style="background-color:#F8F8E8" title="����Ʈ�������� ���� �����մϴ�."><% end if %>
    	<input type="button" value="��ü ����" class="button" onClick="assignPageALL()" style="background-color:#FFF0F0" title="����Ʈ�������� ��ü �����մϴ�.">
    </td>
    <td align="right">
    	<% if C_ADMIN_AUTH then %>
		<input type="button" class="button" value="���ø�����" onClick="popTemplateManage();">&nbsp;
		<% end if %>
		<select class="select" onchange="fnQuitReg(this);">
			<option value="">-- ������� --</option>
			<%
				if oTemplate.FResultCount>0 then
					for lp=0 to (oTemplate.FResultCount-1)
						Response.Write "<option value='" & oTemplate.FItemList(lp).FtplIdx & "' " & chkIIF(cStr(oTemplate.FItemList(lp).FtplIdx)=tplIdx,"selected","") & ">" & oTemplate.FItemList(lp).FtplName & "</option>"
					next
				end if
			%>
		</select>&nbsp;
    	<input type="button" value="������ ���" class="button" onClick="goMainContent('');">
    </td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ��� ���� -->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="main">
<input type="hidden" name="sTrm" value="0">
<input type="hidden" name="chkAll" value="N">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout: fixed;">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		�˻���� : <b><%=oMainCont.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oMainCont.FtotalPage%></b>
	</td>
</tr>
<colgroup>
    <col width="30" />
    <col width="50" />
    <col width="80" />
    <col width="120" />
    <col width="*" />
    <col width="150" />
    <col width="60" />
    <col width="110" />
    <col width="80" />
    <col width="70" />
</colgroup>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>&nbsp;</td>
    <td>��ȣ</td>
    <td>�������</td>
    <td>���ø�</td>
    <td>����</td>
    <td>����Ⱓ</td>
    <td>�켱<br>����</td>
    <td>�����⿩��</td>
    <td>�����</td>
    <td>�۾���</td>
</tr>
<tbody id="mainList">
<%	for lp=0 to oMainCont.FResultCount - 1 %>
<tr align="center" bgcolor="<%=chkIIF(oMainCont.FItemList(lp).IsExpired,"#DDDDDD","#FFFFFF")%>">
    <td><input type="checkbox" name="chkIdx" value="<%=oMainCont.FItemList(lp).FmainIdx%>" /></td>
    <td><a href="javascript:goMainContent(<%=oMainCont.FItemList(lp).FmainIdx%>)"><%=oMainCont.FItemList(lp).FmainIdx%></a></td>
    <td><%=oMainCont.FItemList(lp).getMainStat%></td>
    <td><%=oMainCont.FItemList(lp).FtplName%></td>
    <td align="left"><a href="javascript:goMainContent(<%=oMainCont.FItemList(lp).FmainIdx%>)"><%=oMainCont.FItemList(lp).FmainTitle%></a></td>
    <td>
    <%
    	Response.Write "����: "
    	Response.Write replace(left(oMainCont.FItemList(lp).FmainStartDate,10),"-",".") & " / " & Num2Str(hour(oMainCont.FItemList(lp).FmainStartDate),2,"0","R") & ":" &Num2Str(minute(oMainCont.FItemList(lp).FmainStartDate),2,"0","R")
    	Response.Write "<br />����: "
    	Response.Write replace(left(oMainCont.FItemList(lp).FmainEndDate,10),"-",".") & " / " & Num2Str(hour(oMainCont.FItemList(lp).FmainEndDate),2,"0","R") & ":" & Num2Str(minute(oMainCont.FItemList(lp).FmainEndDate),2,"0","R")
    %>
    </td>
    <td><input type="text" name="sort<%=oMainCont.FItemList(lp).FmainIdx%>" size="3" class="text" value="<%=oMainCont.FItemList(lp).FmainSortNo%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oMainCont.FItemList(lp).FmainIdx%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oMainCont.FItemList(lp).FmainIsPreOpen="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">����</label><input type="radio" name="use<%=oMainCont.FItemList(lp).FmainIdx%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oMainCont.FItemList(lp).FmainIsPreOpen="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">����</label>
		</span>
    </td>
    <td><%=oMainCont.FItemList(lp).FmainRegUsername%></td>
    <td>
    <%
    	modiTime = oMainCont.FItemList(lp).FmainLastModiDate
    	if Not(modiTime="" or isNull(modiTime)) then
	    		Response.Write getStaffUserName(oMainCont.FItemList(lp).FmainLastModiUserid) & "<br />"
	    		Response.Write left(modiTime,10)
	    end if
    %>
    </td>
</tr>
<%	Next %>
</tbody>
<tr bgcolor="#FFFFFF">
    <td colspan="10" align="center">
    <% if oMainCont.HasPreScroll then %>
		<a href="javascript:goPage('<%= oMainCont.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for lp=0 + oMainCont.StartScrollPage to oMainCont.FScrollCount + oMainCont.StartScrollPage - 1 %>
		<% if lp>oMainCont.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(lp) then %>
		<font color="red">[<%= lp %>]</font>
		<% else %>
		<a href="javascript:goPage('<%= lp %>');">[<%= lp %>]</a>
		<% end if %>
	<% next %>

	<% if oMainCont.HasNextScroll then %>
		<a href="javascript:goPage('<%= lp %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</form>
<form name="frmQuitReg" method="POST" action="doQuickMainCont.asp" style="margin:0;">
<input type="hidden" name="site" value="<%=siteDiv%>" />
<input type="hidden" name="pDiv" value="<%=pageDiv%>" />
<input type="hidden" name="tplIdx" value="" />
<input type="hidden" name="StartDate" value="" />
<input type="hidden" name="EndDate" value="" />
<input type="hidden" name="sTm" value="00:00:00" />
<input type="hidden" name="eTm" value="23:59:59" />
<input type="hidden" name="mainTitle" value="" />
</form>
<!-- ��� �� -->
<%
	set oTemplate = Nothing
	set oMainCont = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->