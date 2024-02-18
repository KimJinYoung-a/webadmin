<%@ language=vbscript %>
<% option explicit %>
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
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/dataanalysis/dataanalysis_cls.asp"-->

<%
dim i, menupos, mainidx
	menupos = requestCheckVar(request("menupos"),10)
	mainidx = requestCheckVar(request("mainidx"),10)

if mainidx="" or isnull(mainidx) then
	response.write "������ ��ȣ�� �����ϴ�."
	dbget.close() : response.end
end if

dim cchart
SET cchart = New cdataanalysis
	cchart.FCurrPage = 1
	cchart.FPageSize = 1000
	cchart.frectmainidx = mainidx
	'cchart.frectisusing="Y"

	if mainidx<>"" then
		cchart.Getdataanalysis_chart_list()		'/��񿡼� ��Ʈ���� ��ä�� ������
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
		alert("�����Ͻ� �׸��� �������ּ���.");
		return;
	}

	if(confirm("�����Ͻ� ��Ʈ ������ ���� �Ͻðڽ��ϱ�?")) {
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
		alert('��뿩�θ� ������ �ּ���.');
		return false;
	}
	if(frmchartone.chartsortno.value!=''){
		if (!IsDouble(frmchartone.chartsortno.value)){
			alert('������ ���ڸ� �Է� ���� �մϴ�.');
			frmchartone.chartsortno.focus();
			return;
		}
	}else{
		alert('���İ��� �Է����ּ���.');
		return false;
	}

	if(confirm("��Ʈ ������ �ű� ���� �Ͻðڽ��ϱ�?")) {
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
		�������� ��ȣ : <font color="red"><%= mainidx %></font> ��Ʈ������ ���� ��ʴϴ�.
		<form name="frmchartone" method="POST" action="" style="margin:0;">
		<input type="hidden" name="mode" value="">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="mainidx" value="<%= mainidx %>">
		<table width="100%" align="left" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" id="gochartreg" style="display:<% if cchart.FtotalCount<1 then %><% else %>none<% end if %>;">
		<tr align="center" bgcolor="FFFFFF">
			<td colspan="2">
				��Ʈ �űԵ��
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>ä��</td>
			<td bgcolor="FFFFFF" align="left">
				<% drawSelectBoxdataanalysisgubun "channeltype", "", "", "N", "channeltype" %>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>��Ʈ</td>
			<td bgcolor="FFFFFF" align="left">
				<% drawSelectBoxdataanalysisgubun "charttype", "", "", "N", "chart" %>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>��ġ��</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="position" size=32 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>��ġ��<br>(�񱳽�)</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="positionpretype" size=32 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>�ɼ�1</td>
			<td bgcolor="FFFFFF" align="left">
				<textarea name="option1" cols=30 rows=3></textarea>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>�ɼ�2</td>
			<td bgcolor="FFFFFF" align="left">
				<textarea name="option2" cols=30 rows=3></textarea>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>����</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="chartsortno" size=2 class="text" value="100" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>��뿩��</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="radio" name="isusing" value="Y" checked />Y
				<input type="radio" name="isusing" value="N" />N
			</td>
		</tr>
		<tr align="center" bgcolor="FFFFFF">
			<td colspan="2">
				<input type="button" onClick="savechartone();" value="�ű�����" class="button">
				&nbsp;
				<input type="button" onClick="gochartregclose();" value="�ݱ�" class="button">
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
				<input type="button" onClick="savechartList();" value="��������" class="button">
				&nbsp;
				<input type="button" onClick="gochartreg();" value="�űԵ��" class="button">
			</td>
		</tr>
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="20">
				�˻���� : <b><%= cchart.FtotalCount %></b>
			</td>
		</tr>
		<% if cchart.FtotalCount>0 then %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		    <td width=30><input type="button" value="��ü" class="button" onClick="chkAllchartItem();"></td>
		    <td width=60>��Ʈidx</td>
		    <td width=100>ä��</td>
		    <td width=130>��Ʈ</td>
		    <td>��ġ��</td>
		    <td>��ġ��<br>(�񱳽�)</td>
		    <td>�ɼ�1</td>
		    <td>�ɼ�2</td>
		    <td width=40>����</td>
		    <td width=60>��뿩��</td>
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
				<td colspan="20" align="center">�˻������ �����ϴ�.</td>
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