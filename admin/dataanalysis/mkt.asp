<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �����ͺм� �湮��
' History : 2016.01.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_html5.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<!-- #include virtual="/admin/dataanalysis/dataanalysis_menu.asp"-->
<!-- #include virtual="/lib/classes/dataanalysis/dataanalysis_salesissue_cls.asp"-->
<%
if groupcd="" then groupcd="all"

'/������ ������ ���İ� Ʋ���� ����
if orggroupcd<>groupcd then
	measure=""
	dimension="" : shchannel="" : shmakerid="" : shdate="" : ordtype="" : pretype="" : pretypeuse="" : startdate="" : enddate=""
	syyyy="" : smm="" : eyyyy="" : emm=""
elseif orgmeasure<>measure then
	dimension="" : shchannel="" : shmakerid="" : shdate="" : ordtype="" : pretype="" : pretypeuse="" : startdate="" : enddate=""
	syyyy="" : smm="" : eyyyy="" : emm=""
end if

'/�Ŵ����� ///////////////////////////////////////////////////////////////////////
dim cdata
SET cdata = New cdataanalysis
	cdata.FCurrPage = 1
	cdata.FPageSize = 1000
	cdata.frectgroupcd = groupcd
	cdata.frectisusing="Y"
	cdata.Getdataanalysis_maingroup_list()	'/��񿡼� �Ŵ����� ��ä�� ������

if cdata.FResultCount>0 then
	'/�Ŵ� �ʱⰪ ����
	cdata.getdefaultsetting
	if measure="" then measure=cdata.fdefault_measure
	if dimension="" then
		dimensiongubun=cdata.fdefault_dimensiongubun
		dimension=cdata.fdefault_dimension
	end if
	if pretype="" then
		pretypegubun=cdata.fdefault_pretypegubun
	end if

	for i = 0 to cdata.FResultCount -1
		'/���� �׷쿡�� �������� �������� ���õ� ������ �޾ƿ´�
		if measure = cdata.FItemList(i).fmeasure then
			mainidx = cdata.FItemList(i).fmainidx
			kind = cdata.FItemList(i).fkind
			dimensiongubun = cdata.FItemList(i).fdimensiongubun
			pretypegubun = cdata.FItemList(i).fpretypegubun
			shchannelgubun = cdata.FItemList(i).fshchannelgubun
			shmakeridgubun = cdata.FItemList(i).fshmakeridgubun
			goapiurl = cdata.FItemList(i).fapiurl
				if goapiurl="" then goapiurl=cdata.fdefault_apiurl
			comment = cdata.FItemList(i).fcomment
			shdategubun = cdata.FItemList(i).fshdategubun
			if shdate="" then
				if shdategubun="" or isnull(shdategubun) then
					shdate="defaultdate"
				else
					shdate="ipkumdate"
				end if
			end if
			shdateunit = cdata.FItemList(i).fshdateunit
			shdatetermgubun = cdata.FItemList(i).fshdatetermgubun
			ordtypegubun = cdata.FItemList(i).fordtypegubun
		end if
	next

	'/��¥���� : YYYY
	if shdateunit="1" then
		if syyyy="" then
			syyyy=year(date)
			eyyyy=year(date)
			startdate = syyyy
			enddate = syyyy
		else
			startdate = syyyy
			enddate = syyyy
		end if

	'/��¥���� : YYYY-MM
	elseif shdateunit="2" then
		if syyyy="" then
			syyyy=year(date)
			smm=Format00(2,month(date))
			sdd=Format00(2,day(date))
			eyyyy=year(date)
			emm=Format00(2,month(date))
			edd=Format00(2,day(date))
			startdate = dateserial(syyyy,smm,sdd)
			enddate = dateserial(eyyyy,emm,edd)
		else
			startdate = dateserial(syyyy,smm,sdd)
			enddate = dateserial(eyyyy,emm,edd)
		end if

	'/��¥���� : YYYY-MM-DD
	else
		if startdate="" or enddate="" then
			if shdatetermgubun=1 then	'/����
				startdate = date()
				enddate = date()
			elseif shdatetermgubun=2 then	'/�Ѵ�
				startdate = dateadd("m", -1, date())
				enddate = date()
			else	'/�̵�Ͻ� �⺻�� ������
				startdate = dateadd("d", -7, date())
				enddate = date()
			end if
		end if
	end if
else
	response.write "�����Ͱ� �����ϴ�."
	dbget.close()	:	response.end
end if
'/�Ŵ����� ///////////////////////////////////////////////////////////////////////

'/��Ʈ���� ///////////////////////////////////////////////////////////////////////
dim cchart
SET cchart = New cdataanalysis
	cchart.FCurrPage = 1
	cchart.FPageSize = 1000
	cchart.frectmainidx = mainidx
	cchart.frectisusing="Y"

	if mainidx<>"" then
		cchart.Getdataanalysis_chart_list()		'/��񿡼� ��Ʈ���� ��ä�� ������
	end if

if cchart.FResultCount>0 then
%>
	<script type="text/javascript">
		var chartidx =new Array(<%= cchart.FResultCount -1 %>);
		var channeltype =new Array(<%= cchart.FResultCount -1 %>);
		var charttype =new Array(<%= cchart.FResultCount -1 %>);
		var position =new Array(<%= cchart.FResultCount -1 %>);
		var positionpretype =new Array(<%= cchart.FResultCount -1 %>);
		var option1 =new Array(<%= cchart.FResultCount -1 %>);
		var option2 =new Array(<%= cchart.FResultCount -1 %>);
	
		<%
		for i = 0 to cchart.FResultCount -1
			'/��Ʈ �ʱⰪ ����
			if i=0 then
				cchart.fdefault_channeltype = cchart.FItemList(i).fchanneltype
			end if
		%>
			chartidx[<%= i %>] = '<%= cchart.FItemList(i).fchartidx %>';
		 	channeltype[<%= i %>] = '<%= cchart.FItemList(i).fchanneltype %>';
			charttype[<%= i %>] = '<%= cchart.FItemList(i).fcharttype %>';
	
			<% if cchart.FItemList(i).fposition<>"" then %>
				position[<%= i %>] = <%= cchart.FItemList(i).fposition %>;
			<% else %>
				position[<%= i %>] = "";
			<% end if %>
			<% if cchart.FItemList(i).fpositionpretype<>"" then %>
				positionpretype[<%= i %>] = <%= cchart.FItemList(i).fpositionpretype %>;
			<% else %>
				positionpretype[<%= i %>] = "";
			<% end if %>
			<% if cchart.FItemList(i).foption1<>"" then %>
				option1[<%= i %>] = <%= cchart.FItemList(i).foption1 %>;
			<% else %>
				option1[<%= i %>] = "";
			<% end if %>
			<% if cchart.FItemList(i).foption2<>"" then %>
				option2[<%= i %>] = <%= cchart.FItemList(i).foption2 %>;
			<% else %>
				option2[<%= i %>] = "";
			<% end if %>
		<% next %>
	</script>
<%
	if channeltype="" then channeltype = cchart.fdefault_channeltype
end if
'/��Ʈ���� ///////////////////////////////////////////////////////////////////////

'api ��� ���� & ����///////////////////////////////////////////////////////////////////////
if goapiurl="" or isnull(goapiurl) then
	response.write "api �ּҰ� �����ϴ�."
	dbget.close()	:	response.end
end if
goapiurl = goapiurl & "?kind="&kind
goapiurl = goapiurl & "&shdate="&shdate
if startdate<>"" then
	goapiurl = goapiurl & "&startdate="&startdate
end if
if enddate<>"" then
	goapiurl = goapiurl & "&enddate="&enddate
end if
if dimension<>"" then
	goapiurl = goapiurl & "&dimensions="&dimension
end if
if channeltype<>"" then
	goapiurl = goapiurl & "&channel="&channeltype
end if
if measure<>"" then
	goapiurl = goapiurl & "&param2="&measure
end if
if pretype<>"" then
	goapiurl = goapiurl & "&pretype="&pretype
end if
if ordtype<>"" then
	goapiurl = goapiurl & "&ordtype="&ordtype
end if
if shchannel<>"" then
	goapiurl = goapiurl & "&shchannel="&shchannel
end if
if shmakerid<>"" then
	goapiurl = goapiurl & "&shmakerid="&shmakerid
end if
'api ��� ���� & ����///////////////////////////////////////////////////////////////////////

dim osales
set osales = new cdataanalysis_salesissue
	osales.FPageSize = 5
	osales.FCurrPage = 1
	osales.frectisusing = "Y"
	osales.frectstartdate = startdate
	osales.frectenddate = enddate
	'osales.getdataanalysis_salesissue_top()
%>

<script type="text/javascript">

//��ŵ����� ���� �������� ����
var G_chartdata='';

//�� ��Ʈ ���� �������� ����
var G_simpletablechart=0; var G_datatablechartcnt=0;

$(function() {
    // ��Ʈ�� �������� �ε��ϱ� ������ ready �� ȣ��� ���ĺ��� ��Ʈ�� ����Ҽ� ����.
    readyChart(function() {
        // ���� ���� ��Ʈ �Լ����� ����� �� ����.

		//ù�ε��� �ѹ� ���
		getChartdata($('#goapiurl').val());
    });
});

//��Ʈ ������ ���
function getChartdata(goapiurl) {
    load(goapiurl, function(data, error) {
        if ( data.response == 'error' ) {
            alert(data.errmsg);
            return;
        }
        if ( data ) {
        	//��ŵ����� ���������� ����
			G_chartdata = data
			//ù�ε��� ��Ʈ �׸���
			drawChart('<%= channeltype %>');
        } else {
           	alert( "ajax loader error" + error );
            console.log(error);
        }
    });
}

//��Ʈ�� �׸���.
function drawChart(tmpchanneltype) {
	//�񱳰� ���� �ߴ��� üũ
	var pretypeuse=$("input[name=pretypeuse]:checkbox:checked").length;
	var pretype=$("input[name=pretype]:radio:checked").length;
	var ispretype=false;
	if (pretypeuse==1 && pretype==1) ispretype=true;

    if ( G_chartdata ) {
        //var transfromData = transformWithGroupName(G_chartdata, " ");  //group ��� ��ƮŸ��Ʋ
        var transfromData = G_chartdata
		var container=''; var chartoption1='';

		for (var i=0; i < chartidx.length; i++){
			container = charttype[i] + 'Container_' + chartidx[i]	//�ϴܿ� ��Ʈ�� �׷��� ������ ID

			if (charttype[i] == 'simpletablechart' || charttype[i] == 'datatablechart'){
			}else{
				$("#"+container).empty().html('');	//�ϴܿ� ��Ʈ�� �׷��� ���� ��ü ���
				$("#"+container).height('0px');	//�ϴܿ� ��Ʈ�� �׷��� ���� ���̰� ����
			}
			if (channeltype[i] == tmpchanneltype){	//�ش� �Ǵ� ä�θ� �׸���
				if (charttype[i] == 'linechart'){	//������Ʈ �׸���
					$("#"+container).height('360px');
					$("#"+container).width('100%');
					//,'title': '����������'
					//,gridlines: {color: '#fff', count: 15} // continues �� ������.
				    //,dashsWithIndex: [1]	//�÷��� ��÷� ǥ����.
					//,dashs: [true, true, false]		//true�� �÷��� ��÷� ǥ����.
				    //,defaultLineWidth: 5	//���α���
					//,colors: ["#ff0000"]	//"#ff0000", "#fff", "black"
				    //,colors: ["#rrggbb", "#rgb", "black"]	//3���� �������� �÷����� �� �� ����.
					if (option1[i]!=''){	//����� �ɼ�1�� �������
						chartoption1 = option1[i]
					}else{	//����Ѱ� ������ �⺻��
						chartoption1 = {vAxis:{textStyle:{fontSize:12},gridlines:{count:5}},hAxis:{textStyle:{fontSize:12}},pattern:'yyyy-MM-dd',colors:['#dc3912', '#3366cc', '#ff9900', '#109618', '#990099', '#0099c6', '#dd4477', '#66aa00']}
					}
					if (ispretype){	//�� üũ ���ý�
						drawGoogleChartLine(convertDataForGoogleChartLine(transfromData, positionpretype[i], hookDate()).dataTable, container, chartoption1);
					}else{
						drawGoogleChartLine(convertDataForGoogleChartLine(transfromData, position[i], hookDate()).dataTable, container, chartoption1);
					}
				}
				if (charttype[i] == 'piechart'){	//������Ʈ �׸���
					$("#"+container).width('100%');
					$("#"+container).height('300px');
					if (ispretype){	//�� üũ ���ý�
						drawGoogleChartPie(convertDataForGoogleChartPie(transfromData, positionpretype[i]).dataTable, container);
					}else{
						drawGoogleChartPie(convertDataForGoogleChartPie(transfromData, position[i]).dataTable, container);
					}
				}
				if (charttype[i] == 'sumpiechart'){	//�հ�������Ʈ �׸���
					$("#"+container).width('95%');
					$("#"+container).height('300px');
					if (ispretype){	//�� üũ ���ý�
						drawGoogleChartPie(convertDataForGoogleChartPieWithSum(transfromData, positionpretype[i]).dataTable, container);
					}else{
						drawGoogleChartPie(convertDataForGoogleChartPieWithSum(transfromData, position[i]).dataTable, container);
					}
				}
				if (charttype[i] == 'barchart'){	//����Ʈ �׸���
					$("#"+container).width('95%');
					$("#"+container).height('300px');
            		//orientation=horizontal(���ΰ�),vertical(���ΰ��⺻��), isStacked = true(���� ����) �⺻���� false
					if (option1[i]!=''){	//����� �ɼ�1�� �������
						chartoption1 = option1[i]
					}else{	//����Ѱ� ������ �⺻��
						chartoption1 = {'orientation':'horizontal','isStacked':true}
					}
					if (ispretype){	//�� üũ ���ý�
						drawGoogleChartBar(convertDataForGoogleChartBar(transfromData, positionpretype[i]).dataTable, container, chartoption1);
					}else{
						drawGoogleChartBar(convertDataForGoogleChartBar(transfromData, position[i]).dataTable, container, chartoption1);
					}
				}
				if (G_simpletablechart==0){		//ó��1ȸ�� �׸���
					if (charttype[i] == 'simpletablechart'){	//�������̺���Ʈ �׸���		//�Ⱦ��� ��������
						G_simpletablechart = G_simpletablechart + parseInt(G_simpletablechart + 1)
						$("#"+container).width('95%');
						$("#"+container).height('300px');
						if (option1[i]!=''){	//����� �ɼ�1�� �������
							chartoption1 = option1[i]
						}else{	//����Ѱ� ������ �⺻��
							chartoption1 = {info:false,paging:false,searching:false,group:{type:'sum',fixed:0},order:[]}
						}
						drawGoogleChartTable(convertDataForGoogleChartTable(G_chartdata).dataTable, container, chartoption1);
					}
				}
				if (G_datatablechartcnt==0){	//ó��1ȸ�� �׸���
					if (charttype[i] == 'datatablechart'){	//���������̺���Ʈ �׸���
						G_datatablechartcnt = G_datatablechartcnt + parseInt(G_datatablechartcnt + 1)
						$("#"+container).width('95%');
						//, scrollY: '300px'	//�������̺��� �и��Ǵ� �����߻�, ���Ž� ������� �����߻�. paging: true �� ��ó�ؼ� ���
						//,group: { type: 'avg', fixed: 0 }	//type:avg, sum		//fixed �Ҽ��� �߶����
						if (option1[i]!=''){	//����� �ɼ�1�� �������
							chartoption1 = option1[i]
						}else{	//����Ѱ� ������ �⺻��
							chartoption1 = {info:false,paging:false,searching:false,group:{type:'sum',fixed:0},order:[]}
						}
						drawDataTable(convertDataForDataTable(G_chartdata), container, chartoption1);
					}
				}

			}
		}

        //drawRaw(G_chartdata, 'jsonContainer');	//���̽� ������
    }
}

function frmsubmit(layout){
	if (layout=='pretype'){
		frm.pretypeuse.checked=true;
	}else if (layout=='pretypeuse'){
		if (!frm.pretypeuse.checked){
			$('input:radio[name="pretype"]').prop("checked", false);
		}
	}else if (layout=='dimension'){
		$('input:checkbox[name="pretypeuse"]').prop("checked", false);
		$('input:radio[name="pretype"]').prop("checked", false);
	}

	frm.submit();
}

//���ε����� ����
function jsmaindatareg(groupcd, menupos){
	var jsmaindatareg = window.open('/admin/dataanalysis/manager/maindata_edit.asp?searchgroupcd='+groupcd+'&menupos='+menupos,'jsmaindatareg','width=1480,height=768,scrollbars=yes,resizable=yes');
	jsmaindatareg.focus();
}

//��Ʈ ����
function jschartreg(mainidx, menupos){
	var jschartreg = window.open('/admin/dataanalysis/manager/chart_edit.asp?mainidx='+mainidx+'&menupos='+menupos,'jschartreg','width=1480,height=768,scrollbars=yes,resizable=yes');
	jschartreg.focus();
}

function jsexceldown(){
	var goapiurl = $('#goapiurl').val();
	var jsexceldown = window.open(goapiurl+'&excelyn=Y','jsexceldown','width=1024,height=768,scrollbars=yes,resizable=yes');
	jsexceldown.focus();
}

</script>

<form name="frm" id="frm" method="get" action="" style="margin:0px;">
<input type='hidden' name='menupos' value='<%= menupos %>'>
<input type='hidden' name='orggroupcd' value='<%= groupcd %>'>
<input type='hidden' name='orgmeasure' value='<%= measure %>'>
<input type='hidden' name='kind' value='<%= kind %>'>
<input type='hidden' name='reloadcharton' value=''>
<input type='<% if C_ADMIN_AUTH then %>text<% else %>hidden<% end if %>' name='goapiurl' id='goapiurl' size=180 value='<%= goapiurl %>' readonly>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�׷� : <% drawSelectBoxdataanalysisgroup "groupcd", groupcd, " onchange='frmsubmit("""");'", "", "" %>
		<%'= getdataanalysisgubun(groupcd, "groupcd") %>
		&nbsp;&nbsp;
		<% cdata.drawlayout_measure "measure", measure, " onchange='frmsubmit(""measure"");'", "N" %>
		&nbsp;&nbsp;
		<% cdata.drawlayout_shchannel "shchannel", shchannel, " onchange='frmsubmit(""shchannel"");'", "Y", shchannelgubun %>
		&nbsp;&nbsp;
		<% cdata.drawlayout_shmakerid "shmakerid", shmakerid, " onclick='jsSearchBrandID(this.form.name,""shmakerid"");'", shmakeridgubun %>
		<p>
		<% cdata.drawlayout_shdate "shdate", shdate, " onchange='frmsubmit(""shdate"");'", shdategubun %>
		<%
		'/��¥���� : YYYY
		if shdateunit="1" then
		%>
			<% DrawyearBoxdynamic "syyyy", syyyy, "" %>
		<%
		'/��¥���� : YYYY-MM
		elseif shdateunit="2" then
		%>
			<% DrawYMBoxdynamic "syyyy", syyyy, "smm", smm, "" %>~<% DrawYMBoxdynamic "eyyyy", eyyyy, "emm", emm, "" %>
		<%
		'/��¥���� : YYYY-MM-DD
		else
		%>
			<input id='startdate' name='startdate' value='<%= startdate %>' class='text' size='10' maxlength='10' />
			<img src='https://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
			~<input id='enddate' name='enddate' value='<%= enddate %>' class='text' size='10' maxlength='10' />
			<img src='https://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		<% end if %>
		&nbsp;&nbsp;
		<% cdata.drawlayout_dimension "dimension", dimension, " onclick='frmsubmit(""dimension"");'", "N", dimensiongubun %>
		&nbsp;&nbsp;
		<% cdata.drawlayout_pretype "pretype", pretype, " onclick='frmsubmit(""pretype"");'", "N", pretypegubun, dimension, pretypeuse, " onclick='frmsubmit(""pretypeuse"");'" %>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" onclick='frmsubmit("");' id="btnOk" class="button_s" value="�˻�" >
	</td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" width=250>
		�ؼ��� : <%= nl2br(comment) %>
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="right">
		<% if C_ADMIN_AUTH then %>
			[�����ڸ�� : 
			<input type="button" onClick="jsmaindatareg('<%= groupcd %>', '<%= menupos %>');" value="�����ͼ���" class="button">
			&nbsp;<input type="button" onClick="jschartreg('<%= mainidx %>', '<%= menupos %>');" value="��Ʈ����" class="button">]
			&nbsp;&nbsp;&nbsp;&nbsp;
		<% end if %>

		<% cdata.drawlayout_ordtype "ordtype", ordtype, " onchange='frmsubmit(""ordtype"");'", "Y", ordtypegubun %>
		&nbsp;
		<input type="button" class="button" value="�����ٿ�" onclick="jsexceldown();">
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" width=250>
		<% drawSelectBoxdataanalysischart "channeltype", channeltype, " onclick='drawChart(this.value);'", "Y", "N", mainidx %>
	</td>
</tr>
</table>

<table width="100%" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<%
'/��Ʈ �ѷ��� ����
if cchart.FResultCount>0 then
%>
	<% for i = 0 to cchart.FResultCount -1 %>
		<% if i=0 then 'ù���ϰ�� %>
			<tr bgcolor="#FFFFFF" >
				<!--<td align="left" width="80%">-->
				<td align="left">
					<div id='<%=cchart.FItemList(i).fcharttype %>Container_<%=cchart.FItemList(i).fchartidx %>' ><%= chkiif(i=0,"<img src='https://fiximage.10x10.co.kr/icons/loading16.gif' width=20 height=20>","") %></div>
		
					<% if cchart.FItemList(i).fcharttype="simpletablechart" or cchart.FItemList(i).fcharttype="datatablechart" then %>
						<style type='text/css'>
							div#<%= cchart.FItemList(i).fcharttype %>Container_<%= cchart.FItemList(i).fchartidx %> td, div#<%= cchart.FItemList(i).fcharttype %>Container_<%= cchart.FItemList(i).fchartidx %> th {
								font-size: 12px;
							}
						</style>
					<% end if %>
				</td>
				<!--<td valign="top" width="20%">
					<table width="100%" valign="top" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr bgcolor="#FFFFFF">
						<td align="left" colspan=2>
							<b>�����̽�</b>
						</td>
					</tr>
					<% if osales.FresultCount>0 then %>
						<% for j=0 to osales.FresultCount-1 %>
						<tr bgcolor="#FFFFFF">
							<td align="left">
								<%= FormatDate(osales.FItemList(j).fstartdate,"00.00") %> ~ <%= FormatDate(osales.FItemList(j).fenddate,"00.00") %>
							</td>
							<td align="left">
								<%= chrbyte(osales.FItemList(j).ftitle,30,"Y") %>
							</td>
						</tr>
						<% next %>
					<% else %>
						<tr bgcolor="#FFFFFF">
							<td align="left" colspan=2>
								<b>�����̽� �˻� ����� �����ϴ�.</b>
							</td>
						</tr>
					<% end if %>
					</table>
				</td>-->
			</tr>
		<% else %>
			<tr bgcolor="#FFFFFF">
				<!--<td align="left" colspan=2>-->
				<td align="left">	
					<div id='<%=cchart.FItemList(i).fcharttype %>Container_<%=cchart.FItemList(i).fchartidx %>' ><%= chkiif(i=0,"<img src='https://fiximage.10x10.co.kr/icons/loading16.gif' width=20 height=20>","") %></div>
		
					<% if cchart.FItemList(i).fcharttype="simpletablechart" or cchart.FItemList(i).fcharttype="datatablechart" then %>
						<style type='text/css'>
							div#<%= cchart.FItemList(i).fcharttype %>Container_<%= cchart.FItemList(i).fchartidx %> td, div#<%= cchart.FItemList(i).fcharttype %>Container_<%= cchart.FItemList(i).fchartidx %> th {
								font-size: 12px;
							}
						</style>
					<% end if %>
				</td>
			</tr>
		<% end if %>
	<% next %>
<% end if %>
</table>
<div id='jsonContainer' style='height:300px; width:100%; display:none;'></div>
</form>

<%
set osales=nothing
set cdata=nothing
set cchart=nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->