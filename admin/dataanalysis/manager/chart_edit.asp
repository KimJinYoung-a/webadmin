<%@ language=vbscript %>
<% option explicit %>
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
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/dataanalysis/dataanalysis_cls.asp"-->

<%
dim i, menupos, mainidx
	menupos = requestCheckVar(request("menupos"),10)
	mainidx = requestCheckVar(request("mainidx"),10)

if mainidx="" or isnull(mainidx) then
	response.write "측정값 번호가 없습니다."
	dbget.close() : response.end
end if

dim cchart
SET cchart = New cdataanalysis
	cchart.FCurrPage = 1
	cchart.FPageSize = 1000
	cchart.frectmainidx = mainidx
	'cchart.frectisusing="Y"

	if mainidx<>"" then
		cchart.Getdataanalysis_chart_list()		'/디비에서 차트정보 통채로 가져옴
	end if

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">

function chkAllchartItem() {
	if($("input[name='chartidx']:first").attr("checked")=="checked") {
		$("input[name='chartidx']").attr("checked",false);
	} else {
		$("input[name='chartidx']").attr("checked","checked");
	}
}

function savechartList() {
	var chk=0;
	$("form[name='frmchartlist']").find("input[name='chartidx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("수정하실 항목을 선택해주세요.");
		return;
	}

	if(confirm("지정하신 차트 정보를 저장 하시겠습니까?")) {
		//frmchartlist.target="ifproc";
		frmchartlist.mode.value="chartlistedit";
		frmchartlist.action="/admin/dataanalysis/manager/manager_process.asp";
		frmchartlist.submit();
	}
}

function savechartone() {
	var tmpisusing='';
	for (var i=0; i < frmchartone.isusing.length; i++){
		if (frmchartone.isusing[i].checked){
			tmpisusing = frmchartone.isusing[i].value;
		}
	}
	if (tmpisusing==''){
		alert('사용여부를 선택해 주세요.');
		return false;
	}
	if(frmchartone.chartsortno.value!=''){
		if (!IsDouble(frmchartone.chartsortno.value)){
			alert('정렬은 숫자만 입력 가능 합니다.');
			frmchartone.chartsortno.focus();
			return;
		}
	}else{
		alert('정렬값을 입력해주세요.');
		return false;
	}

	if(confirm("차트 정보를 신규 저장 하시겠습니까?")) {
		//frmchartone.target="ifproc";
		frmchartone.mode.value="inchartreg";
		frmchartone.action="/admin/dataanalysis/manager/manager_process.asp";
		frmchartone.submit();
	}
}

function gochartreg(){
	$("#gochartreg").show();
}
function gochartregclose(){
	$("#gochartreg").hide();
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="FFFFFF">
	<td>
		※측정값 번호 : <font color="red"><%= mainidx %></font> 차트구성을 보고 계십니다.
		<form name="frmchartone" method="POST" action="" style="margin:0;">
		<input type="hidden" name="mode" value="">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="mainidx" value="<%= mainidx %>">
		<table width="100%" align="left" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" id="gochartreg" style="display:<% if cchart.FtotalCount<1 then %><% else %>none<% end if %>;">
		<tr align="center" bgcolor="FFFFFF">
			<td colspan="2">
				차트 신규등록
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>채널</td>
			<td bgcolor="FFFFFF" align="left">
				<% drawSelectBoxdataanalysisgubun "channeltype", "", "", "N", "channeltype" %>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>차트</td>
			<td bgcolor="FFFFFF" align="left">
				<% drawSelectBoxdataanalysisgubun "charttype", "", "", "N", "chart" %>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>위치값</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="position" size=32 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>위치값<br>(비교시)</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="positionpretype" size=32 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>옵션1</td>
			<td bgcolor="FFFFFF" align="left">
				<textarea name="option1" cols=30 rows=3></textarea>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>옵션2</td>
			<td bgcolor="FFFFFF" align="left">
				<textarea name="option2" cols=30 rows=3></textarea>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>정렬</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="chartsortno" size=2 class="text" value="100" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>사용여부</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="radio" name="isusing" value="Y" checked />Y
				<input type="radio" name="isusing" value="N" />N
			</td>
		</tr>
		<tr align="center" bgcolor="FFFFFF">
			<td colspan="2">
				<input type="button" onClick="savechartone();" value="신규저장" class="button">
				&nbsp;
				<input type="button" onClick="gochartregclose();" value="닫기" class="button">
			</td>
		</tr>
		</table>
		</form>
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td>
		<br>
		<form name="frmchartlist" method="POST" action="" style="margin:0;">
		<input type="hidden" name="chkAll" value="N">
		<input type="hidden" name="mode" value="">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="mainidx" value="<%= mainidx %>">
		<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="20" align="right">
				<input type="button" onClick="savechartList();" value="선택저장" class="button">
				&nbsp;
				<input type="button" onClick="gochartreg();" value="신규등록" class="button">
			</td>
		</tr>
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="20">
				검색결과 : <b><%= cchart.FtotalCount %></b>
			</td>
		</tr>
		<% if cchart.FtotalCount>0 then %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		    <td width=30><input type="button" value="전체" class="button" onClick="chkAllchartItem();"></td>
		    <td width=60>차트idx</td>
		    <td width=100>채널</td>
		    <td width=130>차트</td>
		    <td>위치값</td>
		    <td>위치값<br>(비교시)</td>
		    <td>옵션1</td>
		    <td>옵션2</td>
		    <td width=40>정렬</td>
		    <td width=60>사용여부</td>
		</tr>
		<%	for i=0 to cchart.FResultCount - 1 %>
		<tr bgcolor="<%=chkIIF(cchart.FItemList(i).fisusing="Y","#FFFFFF","#DDDDDD")%>" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='<%=chkIIF(cchart.FItemList(i).fisusing="Y","#FFFFFF","#DDDDDD")%>';>
		    <td><input type="checkbox" name="chartidx" value="<%= cchart.FItemList(i).fchartidx %>" /></td>
		    <td><%= cchart.FItemList(i).fchartidx %></td>
		    <td><% drawSelectBoxdataanalysisgubun "channeltype_" & cchart.FItemList(i).fchartidx, cchart.FItemList(i).fchanneltype, "", "N", "channeltype" %></td>
		    <td><% drawSelectBoxdataanalysisgubun "charttype_" & cchart.FItemList(i).fchartidx, cchart.FItemList(i).fcharttype, "", "N", "chart" %></td>
		    <td>
		    	<input type="text" name="position_<%= cchart.FItemList(i).fchartidx %>" size=15 class="text" value="<%= cchart.FItemList(i).fposition %>" />
		    </td>
		    <td>
		    	<input type="text" name="positionpretype_<%= cchart.FItemList(i).fchartidx %>" size=15 class="text" value="<%= cchart.FItemList(i).fpositionpretype %>" />
		    </td>
		    <td>
		    	<textarea name="option1_<%= cchart.FItemList(i).fchartidx %>" cols=30 rows=3><%= cchart.FItemList(i).foption1 %></textarea>
		    </td>
		    <td>
		    	<textarea name="option2_<%= cchart.FItemList(i).fchartidx %>" cols=30 rows=3><%= cchart.FItemList(i).foption2 %></textarea>
		    </td>
		    <td>
		    	<input type="text" name="chartsortno_<%= cchart.FItemList(i).fchartidx %>" size=2 class="text" value="<%= cchart.FItemList(i).fchartsortno %>" />
		    </td>
		    <td>
				<input type="radio" name="isusing_<%= cchart.FItemList(i).fchartidx %>" value="Y" <%=chkIIF(cchart.FItemList(i).fisusing="Y" or isnull(cchart.FItemList(i).fisusing) or cchart.FItemList(i).fisusing="","checked","")%> />Y
				<input type="radio" name="isusing_<%= cchart.FItemList(i).fchartidx %>" value="N" <%=chkIIF(cchart.FItemList(i).fisusing="N","checked","")%> />N
		    </td>
		</tr>
		<%	Next %>
		<% else %>
			<tr bgcolor="#FFFFFF">
				<td colspan="20" align="center">검색결과가 없습니다.</td>
			</tr>
		<% end if %>
		</table>
		</form>
	</td>
</tr>
</table>

<iframe id="ifproc" name="ifproc" width="0" height="0"></iframe>

<%
set cchart=nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->