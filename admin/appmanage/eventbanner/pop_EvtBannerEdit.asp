<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/appmanage/eventBannerCls.asp" -->
<%
'###############################################
' PageName : pop_EvtBannerEdit.asp
' Discription : �̺�Ʈ �ų� ���/����
' History : 2014.03.28 ������ : �ű� ����
'###############################################

'// ���� ����
Dim i, page
Dim oEvtBanner
Dim Idx, appName, startDate, endDate, eventName, sortNo, bannerType, bannerImg, bannerLink, isUsing, regUserid, regdate, lastUpdateUser, lastUpdate, workComment
Dim startTime, endTime


'// �Ķ���� ����
idx = request("idx")
appName = request("appName")
if appName="" then appName="wishapp"

'// ���ø� ����
	set oEvtBanner = new CEvtBanner
	oEvtBanner.FRectIdx = idx
    if idx<>"" then
    	oEvtBanner.GetOneEvtBanner()
		if oEvtBanner.FResultCount>0 then
            appName			= oEvtBanner.FOneItem.FappName
			startDate		= left(oEvtBanner.FOneItem.FstartDate,10)
            endDate			= left(oEvtBanner.FOneItem.FendDate,10)
            eventName		= oEvtBanner.FOneItem.FeventName
            sortNo			= oEvtBanner.FOneItem.FsortNo
            bannerType		= oEvtBanner.FOneItem.FbannerType
            bannerImg		= oEvtBanner.FOneItem.FbannerImg
            bannerLink		= oEvtBanner.FOneItem.FbannerLink
            isUsing			= oEvtBanner.FOneItem.FisUsing
            regUserid		= oEvtBanner.FOneItem.FregUserid
            regdate			= oEvtBanner.FOneItem.Fregdate
            lastUpdateUser	= oEvtBanner.FOneItem.FlastUpdateUser
            lastUpdate		= oEvtBanner.FOneItem.FlastUpdate
            workComment		= oEvtBanner.FOneItem.FworkComment

            startTime		= Num2Str(hour(oEvtBanner.FOneItem.FstartDate),2,"0","R") & ":" & Num2Str(minute(oEvtBanner.FOneItem.FstartDate),2,"0","R") & ":" & Num2Str(second(oEvtBanner.FOneItem.FstartDate),2,"0","R")
            endTime			= Num2Str(hour(oEvtBanner.FOneItem.FendDate),2,"0","R") & ":" & Num2Str(minute(oEvtBanner.FOneItem.FendDate),2,"0","R") & ":" & Num2Str(second(oEvtBanner.FOneItem.FendDate),2,"0","R")
		end if
    else
    	startDate		= date
    	EndDate			= date
    	bannerType		= "H"
    	startTime		= "00:00:00"
    	endTime			= "23:59:59"
    	regdate = now()
    	sortNo = "50"
    	isUsing="N"
    end if
    set oEvtBanner = Nothing
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
$(function(){
	//���� ��ư
	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");

	// Ķ����
	var arrDayMin = ["��","��","ȭ","��","��","��","��"];
	var arrMonth = ["1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��"];
    $("#startDate").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 1,
      	showOn: "button",
    	onClose: function() {
    		if($("#startDate").datepicker("getDate")>$("#endDate").datepicker("getDate")) {
    			$("#endDate").datepicker("setDate",$("#startDate").datepicker("getDate"));
    		}
    	}
    });
    $("#endDate").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
    	numberOfMonths: 1,
      	showOn: "button",
    	onClose: function() {
    		if($("#endDate").datepicker("getDate")<$("#startDate").datepicker("getDate")) {
    			$("#startDate").datepicker("setDate",$("#endDate").datepicker("getDate"));
    		}
    	}
    });
});

// ���˻�
function SaveEvtBanner(frm) {
	var selChk=true;
	$("select").each(function(){
		if($(this).val()=="") {
			alert($(this).attr("title")+"��(��) �������ּ���");
			$(this).focus();
			selChk=false;
			return false;
		}
	});
	if(!selChk) return;

	if($("input[name='eventName']").val()=="") {
		alert("������ �Է����ּ���.");
		$("input[name='eventName']").focus();
		selChk=false;
	}

	if(selChk) {
		frm.submit();
	} else {
		return;
	}
}

function jsSetImg(sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('pop_evtBanner_upload.asp?yr=<%=Year(regdate)%>&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   $("#"+sName).val('');
	   $("#"+sSpan).fadeOut();
	}
}

function fnDefaulUrl(v) {
	if(v=="e") {
		$("input[name='bannerLink']").val("/event/eventmain.asp?eventid=");
	} else if(v=="i") {
		$("input[name='bannerLink']").val("/category/category_itemPrd.asp?itemid=");
	}
}
</script>
<center>
<form name="frmEvtBanner" method="post" action="doEvtBanner.asp" style="margin:0px;">
<table width="690" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
    <td height="25" colspan="4" bgcolor="#F8F8F8"><b>�̺�Ʈ ��� ���/����</b></td>
</tr>
<% if idx<>"" then %>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">��ȣ</td>
    <td width="610" colspan="3">
        <%=idx %>
        <input type="hidden" name="idx" value="<%=idx %>" />
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">���ó</td>
    <td width="230">
        <select name="appName" class="select" title="���ó">
        	<option value="">::����::</option>
			<option value="wishapp" <%=chkIIF(appName="wishapp","selected","")%> >���� APP</option>
			<option value="calapp" <%=chkIIF(appName="calapp","selected","")%> >Ķ���� APP</option>
			<option value="hitchhiker" <%=chkIIF(appName="hitchhiker","selected","")%> >��ġ����Ŀ</option>
        </select>
    </td>
    <td width="100" bgcolor="#DDDDFF">�������</td>
    <td width="230">
		<select name="bannerType" class="select" title="�������">
			<option value="">��ü</option>
			<option value="F" <%=chkIIF(bannerType="F","selected","")%> >Ǯ���</option>
			<option value="H" <%=chkIIF(bannerType="H","selected","")%> >�������</option>
		</select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">���⿩��</td>
    <td width="230">
		<% if idx<>"" then %>
		<span class="rdoUsing">
		<input type="radio" name="isusing" id="rdoUsing_1" value="Y" <%=chkIIF(isUsing="Y","checked","")%> /><label for="rdoUsing_1">����</label><input type="radio" name="isusing" id="rdoUsing_2" value="N" <%=chkIIF(isUsing="N","checked","")%> /><label for="rdoUsing_2">����</label>
		</span>
		<% else %>
		<input type="hidden" name="isusing" value="N">
		�������<br><span style="color:#D03030;font-size:11px;">�� ���� ��Ͻ� �����Ұ� (��� �� ���� ���)</span>
		<% end if %>
    </td>
    <td width="100" bgcolor="#DDDDFF">���ļ���</td>
    <td width="230">
		<input type="text" name="sortNo" class="text" size="4" value="<%=sortNo%>" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">����</td>
    <td width="610" colspan="3">
        <input type="text" name="eventName" value="<%= eventName %>" maxlength="64" size="64" title="����">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">����Ⱓ</td>
    <td width="610" colspan="3">
		<input type="text" id="startDate" name="startDate" size="10" value="<%=startDate%>" style="height:22px;" />
		<input type="text" name="startTime" size="8" value="<%=startTime%>" style="height:22px;"> ~
		<input type="text" id="endDate" name="endDate" size="10" value="<%=endDate%>" style="height:22px;" />
		<input type="text" name="endTime" size="8" value="<%=endTime%>" style="height:22px;">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">�۾����޻���</td>
    <td width="610" colspan="3">
		<textarea name="workcomment" style="width:100%; height:90px;"><%=workcomment%></textarea>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">����̹���</td>
    <td width="610" colspan="3">
		<input type="hidden" name="bannerImg" id="bannerImg" value="<%=bannerImg%>">
		<input type="button" value="�̹��� ���" onClick="jsSetImg('<%=bannerImg%>','bannerImg','spanBanner')" class="button"> <span style="color:#D03030;font-size:11px;">�� 600��282px</span>
		<div id="spanBanner" style="padding: 5 5 5 5">
			<%IF bannerImg <> "" THEN %>
			<img src="<%=bannerImg%>" width="400" border="0">
			<a href="javascript:jsDelImg('bannerImg','spanBanner');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">��� ��ũ</td>
    <td width="610" colspan="3">
        <input type="text" name="bannerLink" value="<%= bannerLink %>" maxlength="64" size="64"><br>
        <span onclick="fnDefaulUrl('e')" style="color:#606060;font-size:11px;cursor:pointer;">ex #1) /event/eventmain.asp?eventid=�̺�Ʈ��ȣ</span><br>
        <span onclick="fnDefaulUrl('i')" style="color:#606060;font-size:11px;cursor:pointer;">ex #2) /category/category_itemPrd.asp?itemid=��ǰ��ȣ</span><br>
        <span style="color:#D03030;font-size:11px;">(APP�� ��η� ��ȯ�Ǵ� ����� Full URL�� �Է����� ������)</span>
    </td>
</tr>



<tr bgcolor="#FFFFFF">
    <td colspan="4" align="center"><input type="button" value=" �� �� " onClick="SaveEvtBanner(this.form);"></td>
</tr>
</table>
</form>
</center>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->