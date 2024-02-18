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
<!-- #include virtual="/lib/classes/appmanage/eventBannerCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : APP ���� �̺�Ʈ ��� ����
' History : 2014.03.27 ������ : �ű� ����
'###############################################

'// ���� ����
Dim appName, bannerType, isUsing, selDt, sDt, eDt 
Dim oEvtBanner, lp
Dim page

'// �Ķ���� ����
appName = request("appName")
isusing = request("isusing")
bannerType = request("bannerType")
sDt = request("sDt")
eDt = request("eDt")
page = request("page")
if appName="" then appName="wishapp"		'�⺻�� wishApp (wishapp:����, hitchhiker:��ġ����Ŀ, calapp:Ķ����)
if isusing="" then isusing="A"				'�⺻�� ��ü
if sDt="" then sDt=cStr(date)
if eDt="" then eDt=cStr(date)
if sDt=eDt then selDt=sDt
if page="" then page="1"

'// ���������� ���
	set oEvtBanner = new CEvtBanner
	oEvtBanner.FPageSize = 20
	oEvtBanner.FRectAppName = appName
	oEvtBanner.FRectIsUsing = isusing
	oEvtBanner.FRectType = bannerType
	oEvtBanner.FRectStartDate = sDt
	oEvtBanner.FRectEndDate = eDt
    oEvtBanner.GetEvtBannerList
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
	/*
	$( "#mainList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="30" colspan="11" style="border:1px solid #F9BD01;">&nbsp;</td>');
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
	*/
});

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
		alert("�����Ͻ� ��ʸ� �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.target="_self";
		document.frmList.action="doListModify.asp";
		document.frmList.submit();
	}
}

function goEvtBannerent(idx) {
    var popwin = window.open('/admin/appmanage/eventbanner/pop_EvtBannerEdit.asp?idx='+idx+'&appName=<%=appName%>','popEvtBanner','width=750,height=420,scrollbars=yes,resizable=yes');
    popwin.focus();
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
	    ���ó:
		<select name="appName" class="select">
			<option value="wishapp" <%=chkIIF(appName="wishapp","selected","")%> >���� APP</option>
			<option value="calapp" <%=chkIIF(appName="calapp","selected","")%> >Ķ���� APP</option>
			<option value="hitchhiker" <%=chkIIF(appName="hitchhiker","selected","")%> >��ġ����Ŀ</option>
		</select>
		&nbsp;/&nbsp;
	    ��뱸��:
		<select name="isusing" class="select">
			<option value="A">��ü</option>
			<option value="Y" <%=chkIIF(isusing="Y","selected","")%> >�����</option>
			<option value="N" <%=chkIIF(isusing="N","selected","")%> >������</option>
		</select>
		&nbsp;/&nbsp;
	    �������:
		<select name="bannerType" class="select">
			<option value="">��ü</option>
			<option value="F" <%=chkIIF(bannerType="F","selected","")%> >Ǯ���</option>
			<option value="H" <%=chkIIF(bannerType="H","selected","")%> >�������</option>
		</select>
	</td>
	<td width="80" rowspan="2" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" value="�˻�" />
	</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
	<td>
		��ȸ�Ⱓ:
		<span id="rdoDtPreset">
		<input type="radio" name="selDatePreset" id="rdoDtOpt1" value="<%=dateadd("d",-1,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",-1,date)),"checked","")%> /><label for="rdoDtOpt1">-1</label><input type="radio" name="selDatePreset" id="rdoDtOpt2" value="<%=date%>" <%=chkIIF(selDt=cStr(date),"checked","")%> /><label for="rdoDtOpt2">����</label><input type="radio" name="selDatePreset" id="rdoDtOpt3" value="<%=dateadd("d",1,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",1,date)),"checked","")%> /><label for="rdoDtOpt3">+1</label><input type="radio" name="selDatePreset" id="rdoDtOpt4" value="<%=dateadd("d",2,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",2,date)),"checked","")%> /><label for="rdoDtOpt4">+2</label><input type="radio" name="selDatePreset" id="rdoDtOpt5" value="<%=dateadd("d",3,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",3,date)),"checked","")%> /><label for="rdoDtOpt5">+3</label><input type="radio" name="selDatePreset" id="rdoDtOpt6" value="<%=dateadd("d",4,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",4,date)),"checked","")%> /><label for="rdoDtOpt6">+4</label><input type="radio" name="selDatePreset" id="rdoDtOpt7" value="<%=dateadd("d",5,date)%>" <%=chkIIF(selDt=cStr(dateadd("d",5,date)),"checked","")%> /><label for="rdoDtOpt7">+5</label>
		</span>
		<input type="text" id="sDt" name="sDt" size="10" value="<%=sDt%>" style="height:22px;" /> ~
		<input type="text" id="eDt" name="eDt" size="10" value="<%=eDt%>" style="height:22px;" />
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
    <td align="left">
    	<input type="button" value="��ü����" class="button" onClick="chkAllItem()">
    	<input type="button" value="��������" class="button" onClick="saveList()" title="�켱���� �� ���⿩�θ� �ϰ������մϴ�.">
    </td>
    <td align="right">
    	<input type="button" value="������ ���" class="button" onClick="goEvtBannerent('');">
    </td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ��� ���� -->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="chkAll" value="N">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout: fixed;">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
		�˻���� : <b><%=oEvtBanner.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oEvtBanner.FtotalPage%></b>
	</td>
</tr>
<colgroup>
    <col width="30" />
    <col width="50" />
    <col width="50" />
    <col width="100" />
    <col width="*" />
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
    <td>Ÿ��</td>
    <td>�̹���</td>
    <td>����</td>
    <td>��ũ</td>
    <td>����Ⱓ</td>
    <td>�켱<br>����</td>
    <td>���⿩��</td>
    <td>�����</td>
    <td>�۾���</td>
</tr>
<tbody id="mainList">
<%	for lp=0 to oEvtBanner.FResultCount - 1 %>
<tr align="center" bgcolor="<%=chkIIF(oEvtBanner.FItemList(lp).IsExpired,"#DDDDDD","#FFFFFF")%>">
    <td><input type="checkbox" name="chkIdx" value="<%=oEvtBanner.FItemList(lp).Fidx%>" /></td>
    <td><a href="javascript:goEvtBannerent(<%=oEvtBanner.FItemList(lp).Fidx%>)"><%=oEvtBanner.FItemList(lp).Fidx%></a></td>
    <td><a href="javascript:goEvtBannerent(<%=oEvtBanner.FItemList(lp).Fidx%>)"><%=oEvtBanner.FItemList(lp).getBannerTypeNm%></a></td>
    <td><a href="javascript:goEvtBannerent(<%=oEvtBanner.FItemList(lp).Fidx%>)"><img src="<%=oEvtBanner.FItemList(lp).FbannerImg%>" alt="���" style="width:94px; border:1px solid #606060;" /></a></td>
    <td align="left"><a href="javascript:goEvtBannerent(<%=oEvtBanner.FItemList(lp).Fidx%>)"><%=oEvtBanner.FItemList(lp).FeventName%></a></td>
    <td align="left"><a href="javascript:goEvtBannerent(<%=oEvtBanner.FItemList(lp).Fidx%>)"><%=oEvtBanner.FItemList(lp).FbannerLink%></a></td>
    <td>
    <%
    	Response.Write "����: "
    	Response.Write replace(left(oEvtBanner.FItemList(lp).FstartDate,10),"-",".") & " / " & Num2Str(hour(oEvtBanner.FItemList(lp).FstartDate),2,"0","R") & ":" &Num2Str(minute(oEvtBanner.FItemList(lp).FstartDate),2,"0","R")
    	Response.Write "<br />����: "
    	Response.Write replace(left(oEvtBanner.FItemList(lp).FendDate,10),"-",".") & " / " & Num2Str(hour(oEvtBanner.FItemList(lp).FendDate),2,"0","R") & ":" & Num2Str(minute(oEvtBanner.FItemList(lp).FendDate),2,"0","R")
    %>
    </td>
    <td><input type="text" name="sort<%=oEvtBanner.FItemList(lp).Fidx%>" size="3" class="text" value="<%=oEvtBanner.FItemList(lp).FsortNo%>" style="text-align:center;" /></td>
    <td>
		<span class="rdoUsing">
		<input type="radio" name="use<%=oEvtBanner.FItemList(lp).Fidx%>" id="rdoUsing<%=lp%>_1" value="Y" <%=chkIIF(oEvtBanner.FItemList(lp).FisUsing="Y","checked","")%> /><label for="rdoUsing<%=lp%>_1">����</label><input type="radio" name="use<%=oEvtBanner.FItemList(lp).Fidx%>" id="rdoUsing<%=lp%>_2" value="N" <%=chkIIF(oEvtBanner.FItemList(lp).FisUsing="N","checked","")%> /><label for="rdoUsing<%=lp%>_2">����</label>
		</span>
    </td>
    <td><%=oEvtBanner.FItemList(lp).FregUsername%></td>
    <td><%=getStaffUserName(oEvtBanner.FItemList(lp).FlastUpdateUser)%>
    </td>
</tr>
<%	Next %>
</tbody>
<tr bgcolor="#FFFFFF">
    <td colspan="11" align="center">
    <% if oEvtBanner.HasPreScroll then %>
		<a href="javascript:goPage('<%= oEvtBanner.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for lp=0 + oEvtBanner.StartScrollPage to oEvtBanner.FScrollCount + oEvtBanner.StartScrollPage - 1 %>
		<% if lp>oEvtBanner.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(lp) then %>
		<font color="red">[<%= lp %>]</font>
		<% else %>
		<a href="javascript:goPage('<%= lp %>');">[<%= lp %>]</a>
		<% end if %>
	<% next %>

	<% if oEvtBanner.HasNextScroll then %>
		<a href="javascript:goPage('<%= lp %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</form>
<div style="text-align:right;">�� [������]���� ������ ��ʴ� ����Ⱓ�� ���� �ڵ����� ���µǸ�, Ǯ��� ���´� ������ʺ��� �켱������ ǥ�õ˴ϴ�.</div>
<!-- ��� �� -->
<%
	set oEvtBanner = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->