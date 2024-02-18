<%
'###########################################################
' Description : 데이터분석 공통 매뉴. 공통 인크루드 파일
' History : 2016.01.29 한용민 생성
'/http://52.79.73.177:5050/
'###########################################################
%>
<!-- #include virtual="/lib/classes/dataanalysis/dataanalysis_cls.asp"-->

<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%
''<!-- #include virtual="/lib/util/aspJSON1.17.asp"-->
%>
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
dim i, j, goapiurl, kind, measure, dimensiongubun, dimension, pretypegubun, pretype, startdate, enddate, pretypeuse
dim ordtype, ordtypegubun, mainidx, channeltype, comment, orggroupcd, groupcd, shchannelgubun, shchannel, shmakeridgubun, shmakerid
dim orgmeasure, shdate, shdategubun, shdateunit, shdatetermgubun, syyyy, smm, sdd, eyyyy, emm, edd
	kind = requestcheckvar(request("kind"),32)
	measure = requestcheckvar(request("measure"),32)
	dimension = requestcheckvar(request("dimension"),32)
	pretype = requestcheckvar(request("pretype"),32)
	startdate = requestcheckvar(request("startdate"),10)
	enddate = requestcheckvar(request("enddate"),10)
	pretypeuse = requestCheckVar(request("pretypeuse"),2)
	ordtype = requestCheckVar(request("ordtype"),32)
	orggroupcd = requestcheckvar(request("orggroupcd"),32)
	orgmeasure = requestcheckvar(request("orgmeasure"),32)
	groupcd = requestcheckvar(request("groupcd"),32)
	shchannel = requestcheckvar(request("shchannel"),32)
	shmakerid = requestcheckvar(request("shmakerid"),32)
	shdate = requestcheckvar(request("shdate"),32)
	syyyy = requestcheckvar(request("syyyy"),4)
	smm = requestcheckvar(request("smm"),2)
	eyyyy = requestcheckvar(request("eyyyy"),4)
	emm = requestcheckvar(request("emm"),2)
%>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link rel="stylesheet" type="text/css" href="/admin/dataanalysis/js/bootstrap.min.css" />
<link href="https://cdn.datatables.net/1.10.11/css/jquery.dataTables.min.css" rel="stylesheet" type="text/css">
<link href="https://cdn.datatables.net/1.10.11/js/dataTables.bootstrap.min.js" rel="stylesheet" type="text/css">

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/json2/20150503/json2.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script src="https://cdn.datatables.net/1.10.11/js/jquery.dataTables.min.js" type="text/javascript"></script>
<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
<% '<script type="text/javascript" src="/admin/dataanalysis/js/lodash.min.js"></script>		'신버전이긴 한데 익스8에서 안된다고 함 %>
<script type="text/javascript" src="/admin/dataanalysis/js/lodash-3.10.1.min.js"></script>
<script type="text/javascript" src="/admin/dataanalysis/js/chart.js"></script>
<script type="text/javascript" src="/admin/dataanalysis/js/main.js"></script>

<script type="text/javascript">

$(function() {
	var CAL_Start = new Calendar({
		inputField : "startdate", trigger    : "startdate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
	var CAL_End = new Calendar({
		inputField : "enddate", trigger    : "enddate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_Start.args.max = date;
			CAL_Start.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
});

</script>