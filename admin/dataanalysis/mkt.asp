<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 데이터분석 방문수
' History : 2016.01.29 한용민 생성
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

'/폼전송 이전과 이후가 틀리면 리셋
if orggroupcd<>groupcd then
	measure=""
	dimension="" : shchannel="" : shmakerid="" : shdate="" : ordtype="" : pretype="" : pretypeuse="" : startdate="" : enddate=""
	syyyy="" : smm="" : eyyyy="" : emm=""
elseif orgmeasure<>measure then
	dimension="" : shchannel="" : shmakerid="" : shdate="" : ordtype="" : pretype="" : pretypeuse="" : startdate="" : enddate=""
	syyyy="" : smm="" : eyyyy="" : emm=""
end if

'/매뉴셋팅 ///////////////////////////////////////////////////////////////////////
dim cdata
SET cdata = New cdataanalysis
	cdata.FCurrPage = 1
	cdata.FPageSize = 1000
	cdata.frectgroupcd = groupcd
	cdata.frectisusing="Y"
	cdata.Getdataanalysis_maingroup_list()	'/디비에서 매뉴정보 통채로 가져옴

if cdata.FResultCount>0 then
	'/매뉴 초기값 셋팅
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
		'/현재 그룹에서 측정값을 기준으로 선택된 값들을 받아온다
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

	'/날짜단위 : YYYY
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

	'/날짜단위 : YYYY-MM
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

	'/날짜단위 : YYYY-MM-DD
	else
		if startdate="" or enddate="" then
			if shdatetermgubun=1 then	'/당일
				startdate = date()
				enddate = date()
			elseif shdatetermgubun=2 then	'/한달
				startdate = dateadd("m", -1, date())
				enddate = date()
			else	'/미등록시 기본값 일주일
				startdate = dateadd("d", -7, date())
				enddate = date()
			end if
		end if
	end if
else
	response.write "데이터가 없습니다."
	dbget.close()	:	response.end
end if
'/매뉴셋팅 ///////////////////////////////////////////////////////////////////////

'/차트셋팅 ///////////////////////////////////////////////////////////////////////
dim cchart
SET cchart = New cdataanalysis
	cchart.FCurrPage = 1
	cchart.FPageSize = 1000
	cchart.frectmainidx = mainidx
	cchart.frectisusing="Y"

	if mainidx<>"" then
		cchart.Getdataanalysis_chart_list()		'/디비에서 차트정보 통채로 가져옴
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
			'/차트 초기값 셋팅
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
'/차트셋팅 ///////////////////////////////////////////////////////////////////////

'api 경로 지정 & 셋팅///////////////////////////////////////////////////////////////////////
if goapiurl="" or isnull(goapiurl) then
	response.write "api 주소가 없습니다."
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
'api 경로 지정 & 셋팅///////////////////////////////////////////////////////////////////////

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

//통신데이터 넣을 전역변수 지정
var G_chartdata='';

//각 차트 수량 전역변수 지정
var G_simpletablechart=0; var G_datatablechartcnt=0;

$(function() {
    // 차트를 동적으로 로딩하기 때문에 ready 가 호출된 이후부터 차트를 사용할수 있음.
    readyChart(function() {
        // 여기 부터 차트 함수들을 사용할 수 있음.

		//첫로딩시 한번 통신
		getChartdata($('#goapiurl').val());
    });
});

//차트 데이터 통신
function getChartdata(goapiurl) {
    load(goapiurl, function(data, error) {
        if ( data.response == 'error' ) {
            alert(data.errmsg);
            return;
        }
        if ( data ) {
        	//통신데이터 전역변수에 넣음
			G_chartdata = data
			//첫로딩시 차트 그린다
			drawChart('<%= channeltype %>');
        } else {
           	alert( "ajax loader error" + error );
            console.log(error);
        }
    });
}

//차트를 그린다.
function drawChart(tmpchanneltype) {
	//비교값 선택 했는지 체크
	var pretypeuse=$("input[name=pretypeuse]:checkbox:checked").length;
	var pretype=$("input[name=pretype]:radio:checked").length;
	var ispretype=false;
	if (pretypeuse==1 && pretype==1) ispretype=true;

    if ( G_chartdata ) {
        //var transfromData = transformWithGroupName(G_chartdata, " ");  //group 헤더 차트타이틀
        var transfromData = G_chartdata
		var container=''; var chartoption1='';

		for (var i=0; i < chartidx.length; i++){
			container = charttype[i] + 'Container_' + chartidx[i]	//하단에 차트가 그려질 영역의 ID

			if (charttype[i] == 'simpletablechart' || charttype[i] == 'datatablechart'){
			}else{
				$("#"+container).empty().html('');	//하단에 차트가 그려질 영역 전체 비움
				$("#"+container).height('0px');	//하단에 차트가 그려질 영역 높이값 리셋
			}
			if (channeltype[i] == tmpchanneltype){	//해당 되는 채널를 그린다
				if (charttype[i] == 'linechart'){	//라인차트 그리기
					$("#"+container).height('360px');
					$("#"+container).width('100%');
					//,'title': '가로축제목'
					//,gridlines: {color: '#fff', count: 15} // continues 만 가능함.
				    //,dashsWithIndex: [1]	//컬럼을 대시로 표현함.
					//,dashs: [true, true, false]		//true인 컬럼을 대시로 표현함.
				    //,defaultLineWidth: 5	//라인굵기
					//,colors: ["#ff0000"]	//"#ff0000", "#fff", "black"
				    //,colors: ["#rrggbb", "#rgb", "black"]	//3가지 포맷으로 컬러값을 줄 수 있음.
					if (option1[i]!=''){	//등록한 옵션1이 있을경우
						chartoption1 = option1[i]
					}else{	//등록한게 없으면 기본값
						chartoption1 = {vAxis:{textStyle:{fontSize:12},gridlines:{count:5}},hAxis:{textStyle:{fontSize:12}},pattern:'yyyy-MM-dd',colors:['#dc3912', '#3366cc', '#ff9900', '#109618', '#990099', '#0099c6', '#dd4477', '#66aa00']}
					}
					if (ispretype){	//비교 체크 선택시
						drawGoogleChartLine(convertDataForGoogleChartLine(transfromData, positionpretype[i], hookDate()).dataTable, container, chartoption1);
					}else{
						drawGoogleChartLine(convertDataForGoogleChartLine(transfromData, position[i], hookDate()).dataTable, container, chartoption1);
					}
				}
				if (charttype[i] == 'piechart'){	//파이차트 그리기
					$("#"+container).width('100%');
					$("#"+container).height('300px');
					if (ispretype){	//비교 체크 선택시
						drawGoogleChartPie(convertDataForGoogleChartPie(transfromData, positionpretype[i]).dataTable, container);
					}else{
						drawGoogleChartPie(convertDataForGoogleChartPie(transfromData, position[i]).dataTable, container);
					}
				}
				if (charttype[i] == 'sumpiechart'){	//합계파이차트 그리기
					$("#"+container).width('95%');
					$("#"+container).height('300px');
					if (ispretype){	//비교 체크 선택시
						drawGoogleChartPie(convertDataForGoogleChartPieWithSum(transfromData, positionpretype[i]).dataTable, container);
					}else{
						drawGoogleChartPie(convertDataForGoogleChartPieWithSum(transfromData, position[i]).dataTable, container);
					}
				}
				if (charttype[i] == 'barchart'){	//바차트 그리기
					$("#"+container).width('95%');
					$("#"+container).height('300px');
            		//orientation=horizontal(세로값),vertical(가로값기본값), isStacked = true(스택 형태) 기본값은 false
					if (option1[i]!=''){	//등록한 옵션1이 있을경우
						chartoption1 = option1[i]
					}else{	//등록한게 없으면 기본값
						chartoption1 = {'orientation':'horizontal','isStacked':true}
					}
					if (ispretype){	//비교 체크 선택시
						drawGoogleChartBar(convertDataForGoogleChartBar(transfromData, positionpretype[i]).dataTable, container, chartoption1);
					}else{
						drawGoogleChartBar(convertDataForGoogleChartBar(transfromData, position[i]).dataTable, container, chartoption1);
					}
				}
				if (G_simpletablechart==0){		//처음1회만 그린다
					if (charttype[i] == 'simpletablechart'){	//심플테이블차트 그리기		//안쓰는 방향으로
						G_simpletablechart = G_simpletablechart + parseInt(G_simpletablechart + 1)
						$("#"+container).width('95%');
						$("#"+container).height('300px');
						if (option1[i]!=''){	//등록한 옵션1이 있을경우
							chartoption1 = option1[i]
						}else{	//등록한게 없으면 기본값
							chartoption1 = {info:false,paging:false,searching:false,group:{type:'sum',fixed:0},order:[]}
						}
						drawGoogleChartTable(convertDataForGoogleChartTable(G_chartdata).dataTable, container, chartoption1);
					}
				}
				if (G_datatablechartcnt==0){	//처음1회만 그린다
					if (charttype[i] == 'datatablechart'){	//데이터테이블차트 그리기
						G_datatablechartcnt = G_datatablechartcnt + parseInt(G_datatablechartcnt + 1)
						$("#"+container).width('95%');
						//, scrollY: '300px'	//사용시테이블이 분리되는 문제발생, 제거시 길어지는 문제발생. paging: true 로 대처해서 사용
						//,group: { type: 'avg', fixed: 0 }	//type:avg, sum		//fixed 소수점 잘라버림
						if (option1[i]!=''){	//등록한 옵션1이 있을경우
							chartoption1 = option1[i]
						}else{	//등록한게 없으면 기본값
							chartoption1 = {info:false,paging:false,searching:false,group:{type:'sum',fixed:0},order:[]}
						}
						drawDataTable(convertDataForDataTable(G_chartdata), container, chartoption1);
					}
				}

			}
		}

        //drawRaw(G_chartdata, 'jsonContainer');	//제이슨 데이터
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

//메인데이터 수정
function jsmaindatareg(groupcd, menupos){
	var jsmaindatareg = window.open('/admin/dataanalysis/manager/maindata_edit.asp?searchgroupcd='+groupcd+'&menupos='+menupos,'jsmaindatareg','width=1480,height=768,scrollbars=yes,resizable=yes');
	jsmaindatareg.focus();
}

//차트 수정
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
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		그룹 : <% drawSelectBoxdataanalysisgroup "groupcd", groupcd, " onchange='frmsubmit("""");'", "", "" %>
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
		'/날짜단위 : YYYY
		if shdateunit="1" then
		%>
			<% DrawyearBoxdynamic "syyyy", syyyy, "" %>
		<%
		'/날짜단위 : YYYY-MM
		elseif shdateunit="2" then
		%>
			<% DrawYMBoxdynamic "syyyy", syyyy, "smm", smm, "" %>~<% DrawYMBoxdynamic "eyyyy", eyyyy, "emm", emm, "" %>
		<%
		'/날짜단위 : YYYY-MM-DD
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
		<input type="button" onclick='frmsubmit("");' id="btnOk" class="button_s" value="검색" >
	</td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" width=250>
		※설명 : <%= nl2br(comment) %>
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="right">
		<% if C_ADMIN_AUTH then %>
			[관리자모드 : 
			<input type="button" onClick="jsmaindatareg('<%= groupcd %>', '<%= menupos %>');" value="데이터수정" class="button">
			&nbsp;<input type="button" onClick="jschartreg('<%= mainidx %>', '<%= menupos %>');" value="차트수정" class="button">]
			&nbsp;&nbsp;&nbsp;&nbsp;
		<% end if %>

		<% cdata.drawlayout_ordtype "ordtype", ordtype, " onchange='frmsubmit(""ordtype"");'", "Y", ordtypegubun %>
		&nbsp;
		<input type="button" class="button" value="엑셀다운" onclick="jsexceldown();">
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
'/차트 뿌려질 영역
if cchart.FResultCount>0 then
%>
	<% for i = 0 to cchart.FResultCount -1 %>
		<% if i=0 then '첫줄일경우 %>
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
							<b>영업이슈</b>
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
								<b>영업이슈 검색 결과가 없습니다.</b>
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