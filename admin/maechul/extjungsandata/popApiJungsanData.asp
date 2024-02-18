<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� ���굥���� API ���
' Hieditor : 2021.02.04 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim extjdate : extjdate = requestCheckvar(request("extjdate"),8) ''YYYYMMDD
if (extjdate="") then
    extjdate = replace(LEFT(dateAdd("d",-1,now()),10),"-","")
end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function jsBySite(s){
	if((s == "ssg" || s == "coupang" || s == "WMP" || s == "wconcept1010")){
		$("#extMeachulDate_span").show();
	}else{
		$("#extMeachulDate_span").hide();
	}

	if (s == "ezwel" || s == "ezwelNew"){
		$("#extMeachulMonth_span").show();
	}else{
		$("#extMeachulMonth_span").hide();
	}
}

function frmSumbit(){
	var sitename	= $("#extsellsite").val();
    var yyyymmdd	= $("#extMeachulDate").val();
	var yyyymm		= $("#extMeachulMonth").val();

	if (sitename == "") {
		alert('���޸��� �����ϼ���');
		return;
	}

    if (confirm('���� ��� �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		switch (sitename){
			case "ssg" 		: document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/xSiteJungsan_ssg_Process.asp?yyyymmdd="+yyyymmdd; break;
			case "ezwel"	: document.frmSvArr.action = "<%=apiURL%>/outmall/jungsan/xSiteJungsan_Ins_Process.asp?sellsite="+sitename+"&reqdate="+yyyymm; break;
			case "ezwelNew"	: document.frmSvArr.action = "/admin/etc/jungsan/xSiteJungsan_Ins_Process.asp?sellsite=ezwel&reqdate="+yyyymm; break;
			case "coupang"	: document.frmSvArr.action = "<%=apiURL%>/outmall/jungsan/xSiteJungsan_Ins_Process.asp?sellsite="+sitename+"&reqdate="+yyyymmdd; break;
			case "WMP"		: document.frmSvArr.action = "<%=apiURL%>/outmall/jungsan/xSiteJungsan_Ins_Process.asp?sellsite="+sitename+"&reqdate="+yyyymmdd; break;
			case "wconcept1010"	: document.frmSvArr.action = "/admin/etc/jungsan/xSiteJungsan_Ins_Process.asp?sellsite=wconcept1010&reqdate="+yyyymmdd; break;
		}
		document.frmSvArr.submit();
    }
}
</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr><td align="left"><b>���޸� ���굥���� API ���</b></td></tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">���޸�</td>
	<td align="left">
		<select class="select" id="extsellsite" name="extsellsite" onChange="jsBySite(this.value);">
			<option value="">-����-</option>
			<option value="ssg">SSG</option>
			<option value="ezwelNew">���������(��API)</option>
			<option value="coupang">����</option>
			<option value="WMP">������</option>
			<option value="wconcept1010">W����</option>
		</select>
		&nbsp;&nbsp;
		<span id="extMeachulDate_span" style="margin-right:20px;display:none;">
			������ : <input type="text" name="extMeachulDate" id="extMeachulDate" value="<%=extjdate%>" size="10" maxlength="10">
		</span>
		<span id="extMeachulMonth_span" style="margin-right:20px;display:none;">
			����� : <input type="text" name="extMeachulMonth" id="extMeachulMonth" value="<%=Replace(Left(DateAdd("m", -1, NOW()), 7), "-", "")%>" size="10" maxlength="10">
		</span>
		<input type="button" class="button" value="���" onClick="frmSumbit();">
	</td>
</tr>
</table>
<form name="frmSvArr" method="post">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="500"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->