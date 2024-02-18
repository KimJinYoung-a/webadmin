<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/exhibitionCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : ����� ���� ���� ��ȹ�� ��ũ
' History : 2016.04.07 ����ȭ
'###############################################
	
	Dim isusing , dispcate
	dim page 
	Dim i
	dim exhibitionList
	Dim sDt , modiTime

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")

	if page="" then page=1
	If isusing = "" Then isusing ="Y"

	set exhibitionList = new Cexhibition
	exhibitionList.FPageSize		= 20
	exhibitionList.FCurrPage		= page
	exhibitionList.Fisusing			= isusing
	exhibitionList.Fsdt				= sDt
	exhibitionList.GetContentsList()

%>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script type='text/javascript'>
<!--
//����
function jsmodify(v){
	location.href = "exhibition_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
}
$(function() {
  	// ������ư
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");
	
});

function RefreshCaFavKeyWordRec(term){
	if(confirm("�����- exhibition�� �����Ͻðڽ��ϱ�?")) {
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.target = "refreshFrm";
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_new_exhibition_xml.asp?term=" + term;
			refreshFrm.submit();
	}
}

function jsquickadd(v){
	if(confirm("�Ϻ� ��������� ���� �Ͻðڽ��ϱ�?")) {
	location.href = "doexhibition.asp?menupos=<%=menupos%>&mode=quickadd&prevDate="+v;
	}
}

function jssearch(){
	document.frm.submit();
}
function addContents(){	
	var dateOptionParam
	var frm = document.frm
	dateOptionParam = frm.prevDate.value

	document.location.href="exhibition_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>&dateoption="+dateOptionParam
}
-->
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			* ��뿩�� :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			�������� <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			<!-- ����� -->
			<% If sDt <> "" Then %>
			��<input type="button" onclick="jsquickadd(document.all.prevDate.value)" value="�������"/>
			<% End If %>
			</div>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onclick="jssearch();">
		</td>
	</tr>
</form>	
</table>
<!-- �˻� �� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<!-- <td>������ �����Ͽ� <input type="text" name="vTerm" value="1" size="1" class="text" style="text-align:right;">�ϰ�<a href="javascript:RefreshCaFavKeyWordRec(document.all.vTerm.value);"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html�����"></a>XML Real ����(����)</a></td> -->
    <td align="right">
		<!-- �űԵ�� -->
    	<a href="javascript:addContents()"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  ����Ʈ -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		�� ��ϼ� : <b><%=exhibitionList.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=exhibitionList.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
    <!-- <td width="10%">������ real ����ð�</td> -->
	<td width="22%">����</td>	 
    <td width="18%">������/������</td>
    <td width="10%">�����</td>
    <td width="10%">�����</td>
    <td width="10%">����������</td>
    <td width="10%">��뿩��</td>
</tr>
<% 
	Dim bgcol : bgcol = ""
	for i=0 to exhibitionList.FResultCount-1 
		
		If exhibitionList.FItemList(i).Fisusing="Y" Then
			If exhibitionList.FItemList(i).Ftopview="Y" Then '��޳���
				bgcol = "#FFB6C1"
			Else
				bgcol = "#FFFFFF"
			End If 
		Else
			bgcol = "#F0F0F0"
		End if
%>
<tr  height="30" align="center" bgcolor="<%=bgcol%>">
    <td onclick="jsmodify('<%=exhibitionList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=exhibitionList.FItemList(i).Fidx%><%=chkiif(exhibitionList.FItemList(i).Ftopview="Y","<br/>(���)","")%></td>
	<!-- <td>
		<%
			If exhibitionList.FItemList(i).Fxmlregdate <> "" then
			Response.Write replace(left(exhibitionList.FItemList(i).Fxmlregdate,10),"-",".") & " / " & Num2Str(hour(exhibitionList.FItemList(i).Fxmlregdate),2,"0","R") & ":" &Num2Str(minute(exhibitionList.FItemList(i).Fxmlregdate),2,"0","R")
			End If 
		%>
	</td> -->
    <td onclick="jsmodify('<%=exhibitionList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=exhibitionList.FItemList(i).Fexhibitiontitle%></td>
	<td onclick="jsmodify('<%=exhibitionList.FItemList(i).Fidx%>');" style="cursor:pointer;" align="left">
		<% 
			If exhibitionList.FItemList(i).Fstartdate <> "" And exhibitionList.FItemList(i).Fenddate Then 
				Response.Write "����: "
				Response.Write replace(left(exhibitionList.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(exhibitionList.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(exhibitionList.FItemList(i).Fstartdate),2,"0","R")
				Response.Write "<br />����: "
				Response.Write replace(left(exhibitionList.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(exhibitionList.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(exhibitionList.FItemList(i).Fenddate),2,"0","R")
				If cInt(datediff("d", now() , exhibitionList.FItemList(i).Fenddate)) < 0 Or cInt(datediff("h", now() , exhibitionList.FItemList(i).Fenddate )) < 0  Then 
					Response.write " <span style=""color:red"">(����)</span>"
				ElseIf cInt(datediff("d", exhibitionList.FItemList(i).Fenddate , now())) = 0  Then '�������� ���ó�¥
					If cInt(datediff("h", now() , exhibitionList.FItemList(i).Fenddate )) >= 0 And cInt(datediff("h", now() , exhibitionList.FItemList(i).Fenddate )) < 24 Then ' ����
						Response.write " <span style=""color:red"">(�� "& cInt(datediff("h", now() , exhibitionList.FItemList(i).Fenddate )) &" �ð��� ����)</span>"
					Else  ' ������
						Response.write " <span style=""color:red"">(������)</span>"					
					End If 
				'// �������� ���ó�¥�̰� �������� ������ �ƴϸ�
				ElseIf cInt(datediff("d", exhibitionList.FItemList(i).Fstartdate , now()))>=0 And cInt(datediff("d", exhibitionList.FItemList(i).Fenddate , now())) < 0 Then
					Response.write " <span style=""color:red"">(�� "&cInt(datediff("d", now() , exhibitionList.FItemList(i).Fenddate ))&"�� " &cInt(datediff("h", now() , exhibitionList.FItemList(i).Fenddate ))-cInt(datediff("d", now() , exhibitionList.FItemList(i).Fenddate ))*24 &"�ð��� ����)</span>"
				ElseIf cInt(datediff("d", exhibitionList.FItemList(i).Fenddate , now())) < 0 Then
					If cInt(datediff("h", now() , exhibitionList.FItemList(i).Fenddate )) >= 0 And cInt(datediff("h", now() , exhibitionList.FItemList(i).Fenddate )) < 24 Then ' ����
						Response.write " <span style=""color:red"">(�� "& cInt(datediff("h", now() , exhibitionList.FItemList(i).Fenddate )) &" �ð��� ����)</span>"
					Else  ' ������
						Response.write " <span style=""color:red"">(������)</span>"
					End If 
				End If
			End If  
		%>
	</td>
	<td><%=left(exhibitionList.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(exhibitionList.FItemList(i).Fadminid)%></td>
	<td>
		<%
			modiTime = exhibitionList.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write getStaffUserName(exhibitionList.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
    <td><%=chkiif(exhibitionList.FItemList(i).Fisusing="N","������","�����")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if exhibitionList.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= exhibitionList.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + exhibitionList.StartScrollPage to exhibitionList.StartScrollPage + exhibitionList.FScrollCount - 1 %>
				<% if (i > exhibitionList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(exhibitionList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if exhibitionList.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set exhibitionList = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->