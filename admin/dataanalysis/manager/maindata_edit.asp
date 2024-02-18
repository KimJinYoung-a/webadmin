<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �����ͺм�
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
dim i, menupos, searchgroupcd
	menupos = requestCheckVar(request("menupos"),10)
	searchgroupcd = requestCheckVar(request("searchgroupcd"),32)

dim cdata
SET cdata = New cdataanalysis
	cdata.FCurrPage = 1
	cdata.FPageSize = 1000
	cdata.frectgroupcd = searchgroupcd
	'cdata.frectisusing="Y"
	cdata.Getdataanalysis_maingroup_list()

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">

function chkAllchartItem() {
	if($("input[name='mainidx']:first").attr("checked")=="checked") {
		$("input[name='mainidx']").attr("checked",false);
	} else {
		$("input[name='mainidx']").attr("checked","checked");
	}
}

function savemaindataList() {
	var chk=0;
	$("form[name='frmmaindatalist']").find("input[name='mainidx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� �׸��� �������ּ���.");
		return;
	}

	if(confirm("�����Ͻ� ��Ʈ ������ ���� �Ͻðڽ��ϱ�?")) {
		//frmmaindatalist.target="ifproc";
		frmmaindatalist.mode.value="maindatalistedit";
		frmmaindatalist.action="/admin/dataanalysis/manager/manager_process.asp";
		frmmaindatalist.submit();
	}
}

function savemaindataone() {
	if (frmmaindataone.groupcd.value=='NEWREG'){
		if (frmmaindataone.groupcdnewreg.value==''){
			alert('�׷��ڵ带 �Է��� �ּ���.');
			frmmaindataone.groupcdnewreg.focus();
			return;
		}
		if (frmmaindataone.groupcdnamenewreg.value==''){
			alert('�׷��ڵ���� �Է��� �ּ���.');
			frmmaindataone.groupcdnamenewreg.focus();
			return;
		}
	}
	if(frmmaindataone.groupsortno.value!=''){
		if (!IsDouble(frmmaindataone.groupsortno.value)){
			alert('�׷������� ���ڸ� �Է� ���� �մϴ�.');
			frmmaindataone.groupsortno.focus();
			return;
		}
	}else{
		alert('���İ��� �Է����ּ���.');
		return false;
	}
	var tmpisusing='';
	for (var i=0; i < frmmaindataone.isusing.length; i++){
		if (frmmaindataone.isusing[i].checked){
			tmpisusing = frmmaindataone.isusing[i].value;
		}
	}
	if (tmpisusing==''){
		alert('��뿩�θ� ������ �ּ���.');
		return false;
	}
	if (frmmaindataone.measure.value==''){
		alert('�������ڵ带 �Է��� �ּ���.');
		frmmaindataone.measure.focus();
		return;
	}
	if (frmmaindataone.measurename.value==''){
		alert('���������� �Է��� �ּ���.');
		frmmaindataone.measurename.focus();
		return;
	}
	if (frmmaindataone.dimensiongubun.value!=''){
		if (!IsDouble(frmmaindataone.dimensiongubun.value)){
			alert('����Ÿ���� ���ڸ� �Է� ���� �մϴ�.');
			frmmaindataone.dimensiongubun.focus();
			return;
		}
	}
	if (frmmaindataone.pretypegubun.value!=''){
		if (!IsDouble(frmmaindataone.pretypegubun.value)){
			alert('��Ÿ���� ���ڸ� �Է� ���� �մϴ�.');
			frmmaindataone.pretypegubun.focus();
			return;
		}
	}
	if (frmmaindataone.shchannelgubun.value!=''){
		if (!IsDouble(frmmaindataone.shchannelgubun.value)){
			alert('ä��Ÿ���� ���ڸ� �Է� ���� �մϴ�.');
			frmmaindataone.shchannelgubun.focus();
			return;
		}
	}
	if (frmmaindataone.shmakeridgubun.value!=''){
		if (!IsDouble(frmmaindataone.shmakeridgubun.value)){
			alert('�귣��Ÿ���� ���ڸ� �Է� ���� �մϴ�.');
			frmmaindataone.shmakeridgubun.focus();
			return;
		}
	}
	if (frmmaindataone.shdategubun.value!=''){
		if (!IsDouble(frmmaindataone.shdategubun.value)){
			alert('�귣��Ÿ���� ���ڸ� �Է� ���� �մϴ�.');
			frmmaindataone.shdategubun.focus();
			return;
		}
	}
	if (frmmaindataone.shdatetermgubun.value!=''){
		if (!IsDouble(frmmaindataone.shdatetermgubun.value)){
			alert('�귣��Ÿ���� ���ڸ� �Է� ���� �մϴ�.');
			frmmaindataone.shdatetermgubun.focus();
			return;
		}
	}
	if(confirm("��Ʈ ������ �ű� ���� �Ͻðڽ��ϱ�?")) {
		//frmmaindataone.target="ifproc";
		frmmaindataone.mode.value="maindatareg";
		frmmaindataone.action="/admin/dataanalysis/manager/manager_process.asp";
		frmmaindataone.submit();
	}
}

function gochartreg(){
	$("#gochartreg").show();
}
function gochartregclose(){
	$("#gochartreg").hide();
}

function gosearch(groupcd){
	location.replace('/admin/dataanalysis/manager/maindata_edit.asp?searchgroupcd='+ groupcd +'&menupos=<%= menupos %>');
}

function chgroupcdnewreg(groupcd){
	if (groupcd=='NEWREG'){
		$("#groupcdnewreg").show();
	}else{
		$("#groupcdnewreg").hide();
	}
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="FFFFFF">
	<td>
		<form name="frmmaindataone" method="POST" action="" style="margin:0;">
		<input type="hidden" name="mode" value="">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<table width="100%" align="left" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" id="gochartreg" style="display:<% if cdata.FtotalCount<1 then %><% else %>none<% end if %>;">
		<tr align="center" bgcolor="FFFFFF">
			<td colspan="2">
				������ �űԵ��
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>�׷�</td>
			<td bgcolor="FFFFFF" align="left">
				<% drawSelectBoxdataanalysisgroup "groupcd", searchgroupcd, " onchange='chgroupcdnewreg(this.value);'", "NEW", "" %> 
				�űԵ���� ���Ͻø� �űԵ���� �����ϼ���.<br>
				<div id="groupcdnewreg" style="display:none;">
					�׷��ڵ� : <input type="text" name="groupcdnewreg" size=32 class="text" value="" /> ex) mkt
					<br>�׷��ڵ�� : <input type="text" name="groupcdnamenewreg" size=32 class="text" value="" /> ex) ������
				</div>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>�׷�����</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="groupsortno" size=3 class="text" value="100" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>����</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="kind" size=32 class="text" value="" /> ex) mktall
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>�������ڵ�</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="measure" size=32 class="text" value="" /> ex) gavisit
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>��������</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="measurename" size=32 class="text" value="" /> ex) �湮��(GA)
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>APiURL</td>
			<td bgcolor="FFFFFF" align="left">
				<input type="text" name="goapiurl" size=32 class="text" value="" /> ex) ���Է½� �⺻�� http://wapi.10x10.co.kr/anal/getque.asp
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>����Ÿ��</td>
			<td bgcolor="FFFFFF" align="left">
				1 : ��,�ð�,��,��,��,����
				<br>2 : ��,��,��,��,����
				<br>-�Է¾��� : ������
				<br><input type="text" name="dimensiongubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>��Ÿ��</td>
			<td bgcolor="FFFFFF" align="left">
				1 : 
				<br>&nbsp;&nbsp;�� ���ý� : ����,����,����,���⵿����
				<br>&nbsp;&nbsp;�ð� ���ý� : ����,����,����,����,���⵿����
				<br>&nbsp;&nbsp;�� ���ý� : ����,����,���⵿����
				<br>&nbsp;&nbsp;�� ���ý� : ����,���⵿����
				<br>&nbsp;&nbsp;�� ���ý� : ����
				<br>&nbsp;&nbsp;���� ���ý� : ����, ����, ����, ���⵿����"
				<br>-�Է¾��� : ������
				<br><input type="text" name="pretypegubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>ä��Ÿ��</td>
			<td bgcolor="FFFFFF" align="left">
				1 : WWW,�����,�����_����,APP,APP_����,���޸�,3PL
				<br>-�Է¾��� : ������
				<br><input type="text" name="shchannelgubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>�귣��Ÿ��</td>
			<td bgcolor="FFFFFF" align="left">
				1 : �귣��˻������
				<br>-�Է¾��� : ������
				<br><input type="text" name="shmakeridgubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>��¥����Ÿ��</td>
			<td bgcolor="FFFFFF" align="left">
				1 : �ֹ���
				<br>2 : �ֹ���,������
				<br>3 : �ֹ���,������,�����
				<br>-�Է¾��� : �Ⱓ
				<br><input type="text" name="shdategubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>��¥����</td>
			<td bgcolor="FFFFFF" align="left">
				1 : YYYY
				<br>2 : YYYY-MM
				<br>-�Է¾��� : YYYY-MM-DD
				<br><input type="text" name="shdateunit" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>��¥�ⰣŸ��</td>
			<td bgcolor="FFFFFF" align="left">
				1 : ����
				<br>2 : �Ѵ�
				<br>-�Է¾��� : ������
				<br><input type="text" name="shdatetermgubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>����Ÿ��</td>
			<td bgcolor="FFFFFF" align="left">
				1 : ��������,��������
				<br>2 : ī�װ����м�,�����,����޼���
				<br>-�Է¾��� : ������
				<br><input type="text" name="ordtypegubun" size=1 class="text" value="" />
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>�ڸ�Ʈ</td>
			<td bgcolor="FFFFFF" align="left">
				<textarea name="comment" cols=30 rows=2></textarea>
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
				<input type="button" onClick="savemaindataone();" value="�ű�����" class="button">
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
		<form name="frmmaindatalist" method="POST" action="" style="margin:0;">
		<input type="hidden" name="chkAll" value="N">
		<input type="hidden" name="mode" value="">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="20" align="right">
				<input type="button" onClick="savemaindataList();" value="��������" class="button">
				&nbsp;
				<input type="button" onClick="gochartreg();" value="�űԵ��" class="button">
			</td>
		</tr>
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="20">
				�˻���� : <b><%= cdata.FtotalCount %></b>
				<br>
				�ذ˻���
				&nbsp;&nbsp;&nbsp;&nbsp;
				�׷� : <% drawSelectBoxdataanalysisgroup "searchgroupcd", searchgroupcd, " onchange='gosearch(this.value);'", "", "" %>
			</td>
		</tr>
		<% if cdata.FtotalCount>0 then %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		    <td width=30><input type="button" value="��ü" class="button" onClick="chkAllchartItem();"></td>
		    <td width=50>����idx</td>
		    <td width=40>�׷�<br>����</td>
		    <td>����</td>
		    <td>�������ڵ�</td>
		    <td>��������</td>
		    <td width=160>APiURL</td>
		    <td width=30>����<br>Ÿ��</td>
		    <td width=30>��<br>Ÿ��</td>
		    <td width=30>ä��<br>Ÿ��</td>
		    <td width=40>�귣��<br>Ÿ��</td>
		    <td width=30>��¥<br>����<br>Ÿ��</td>
		    <td width=30>��¥<br>����</td>
		    <td width=30>��¥<br>�Ⱓ<br>Ÿ��</td>
		    <td width=30>����<br>Ÿ��</td>
		    <td width=160>�ڸ�Ʈ</td>
		    <td width=60>��뿩��</td>
		</tr>
		<%	for i=0 to cdata.FResultCount - 1 %>
		<input type="hidden" name="groupcd_<%= cdata.FItemList(i).fmainidx %>" value="<%= cdata.FItemList(i).fgroupcd %>" />
		<tr bgcolor="<%=chkIIF(cdata.FItemList(i).fisusing="Y","#FFFFFF","#DDDDDD")%>" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='<%=chkIIF(cdata.FItemList(i).fisusing="Y","#FFFFFF","#DDDDDD")%>';>
		    <td><input type="checkbox" name="mainidx" value="<%= cdata.FItemList(i).fmainidx %>" /></td>
		    <td><%= cdata.FItemList(i).fmainidx %></td>
		    <td>
		    	<input type="text" name="groupsortno_<%= cdata.FItemList(i).fmainidx %>" size=2 class="text" value="<%= cdata.FItemList(i).fgroupsortno %>" />
		    </td>
		    <td>
		    	<input type="text" name="kind_<%= cdata.FItemList(i).fmainidx %>" size=15 class="text" value="<%= cdata.FItemList(i).fkind %>" />
		    </td>
		    <td>
		    	<input type="text" name="measure_<%= cdata.FItemList(i).fmainidx %>" size=15 class="text" value="<%= cdata.FItemList(i).fmeasure %>" />
		    </td>
		    <td>
		    	<input type="text" name="measurename_<%= cdata.FItemList(i).fmainidx %>" size=15 class="text" value="<%= cdata.FItemList(i).fmeasurename %>" />
		    </td>
		    <td>
		    	<textarea name="goapiurl_<%= cdata.FItemList(i).fmainidx %>" cols=20 rows=3><%= cdata.FItemList(i).fapiurl %></textarea>
		    </td>
		    <td>
		    	<input type="text" name="dimensiongubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fdimensiongubun %>" />
		    </td>
		    <td>
		    	<input type="text" name="pretypegubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fpretypegubun %>" />
		    </td>
		    <td>
		    	<input type="text" name="shchannelgubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fshchannelgubun %>" />
		    </td>
		    <td>
		    	<input type="text" name="shmakeridgubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fshmakeridgubun %>" />
		    </td>
		    <td>
		    	<input type="text" name="shdategubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fshdategubun %>" />
		    </td>
		    <td>
		    	<input type="text" name="shdateunit_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fshdateunit %>" />
		    </td>
		    <td>
		    	<input type="text" name="shdatetermgubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fshdatetermgubun %>" />
		    </td>
		    <td>
		    	<input type="text" name="ordtypegubun_<%= cdata.FItemList(i).fmainidx %>" size=1 class="text" value="<%= cdata.FItemList(i).fordtypegubun %>" />
		    </td>
		    <td>
		    	<textarea name="comment_<%= cdata.FItemList(i).fmainidx %>" cols=20 rows=3><%= cdata.FItemList(i).fcomment %></textarea>
		    </td>
		    <td align="center">
				<input type="radio" name="isusing_<%= cdata.FItemList(i).fmainidx %>" value="Y" <%=chkIIF(cdata.FItemList(i).fisusing="Y" or isnull(cdata.FItemList(i).fisusing) or cdata.FItemList(i).fisusing="","checked","")%> />Y
				<input type="radio" name="isusing_<%= cdata.FItemList(i).fmainidx %>" value="N" <%=chkIIF(cdata.FItemList(i).fisusing="N","checked","")%> />N
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

<iframe id="ifproc" name="ifproc" width=0 height=0></iframe>

<%
set cdata=nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->