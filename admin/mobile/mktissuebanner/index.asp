<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/mktevtbannerCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : ����� �̺�Ʈ ������ ���
' History : 2015-01-07 ����ȭ
'###############################################
	
	Dim isusing , dispcate , validdate , research
	dim page 
	Dim i
	dim oTodaydealList
	Dim sDt , modiTime , gubun

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")
	validdate= request("validdate")
	research= request("research")
	gubun= request("gubun")

	if ((research="") and (isusing="")) then
	    isusing = "Y"
	    validdate = "on"
	end if
	
	if page="" then page=1

	set oTodaydealList = new CEvtMktbanner
	oTodaydealList.FPageSize		= 20
	oTodaydealList.FCurrPage		= page
	oTodaydealList.Fisusing			= isusing
	oTodaydealList.Fsdt				= sDt
	oTodaydealList.FRectvaliddate	= validdate
	oTodaydealList.FRectgubun		= gubun
	oTodaydealList.GetContentsList()

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
	location.href = "deal_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
}
$(function() {
  	$("input[type=submit]").button();

  	// ������ư
  	$("#rdoDtPreset").buttonset();
	$("input[name='selDatePreset']").click(function(){
		$("#sDt").val($(this).val());
		$("#eDt").val($(this).val());
	}).next().attr("style","font-size:11px;");

});

function NextPage(page){
    frm.page.value = page;
    frm.submit();
}


function Addnewcontents(val){
    var popwin = window.open('mktban_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>&idx='+val,'mainposcodeedit','width=800,height=550,scrollbars=yes,resizable=yes');
    popwin.focus();
}
-->
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="<%=page%>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������&nbsp;
			* ��뿩�� :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			�������� <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			&nbsp;&nbsp;&nbsp;
			���� : <select name="gubun" onchange="onchgbox(this.value);" width="100">
						<option value="">���м���</option>
						<option value="1" <%=chkiif(gubun="1","selected","")%>>Mobile & Apps</option>
						<option value="2" <%=chkiif(gubun="2","selected","")%>>Mobile</option>
						<option value="3" <%=chkiif(gubun="3","selected","")%>>Apps</option>
					</select>&nbsp;&nbsp;
			</div>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:submit();">
		</td>
	</tr>
</form>	
</table>
<!-- �˻� �� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
    <td align="right">
		<!-- �űԵ�� -->
    	<a href="" onclick="Addnewcontents(''); return false;" target="_blank"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  ����Ʈ -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
		�� ��ϼ� : <b><%=oTodaydealList.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oTodaydealList.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
	<td width="10%">����</td>
	<td width="10%">�̺�Ʈ ����</td>
	<td width="7%">����̹���</td>	 
    <td width="20%">������/������</td>
    <td width="10%">�����<br/>����������</td>
    <td width="10%">��������</td>
    <td width="5%">�켱����</td>
    <td width="5%">��뿩��</td>
</tr>
<% 
	for i=0 to oTodaydealList.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oTodaydealList.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="Addnewcontents('<%=oTodaydealList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=oTodaydealList.FItemList(i).Fidx%></td>
	<td><%=getGubun(oTodaydealList.FItemList(i).Fgubun)%></td>
	<td><%=chkiif(oTodaydealList.FItemList(i).Fevtgubun = "1","��ȹ��","������")%></td>
	<td>
		<img src="<%=oTodaydealList.FItemList(i).Fmktimg%>" width="300" alt="<%=oTodaydealList.FItemList(i).Faltname%>"/>
	</td>
	<td align="left">
		<% 
			Response.Write "����� : "
			Response.Write left(oTodaydealList.FItemList(i).Fregdate,10) &"</br>"
			Response.Write "����: "
			Response.Write replace(left(oTodaydealList.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(oTodaydealList.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(oTodaydealList.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />����: "
			Response.Write replace(left(oTodaydealList.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(oTodaydealList.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(oTodaydealList.FItemList(i).Fenddate),2,"0","R")
			
			If cInt(datediff("d", now() , oTodaydealList.FItemList(i).Fenddate)) < 0 Or cInt(datediff("h", now() , oTodaydealList.FItemList(i).Fenddate )) < 0  Then 
					Response.write " <span style=""color:red"">(����)</span>"
			ElseIf cInt(datediff("d", oTodaydealList.FItemList(i).Fenddate , now())) < 1  Then '���ó�¥

				If cInt(datediff("h", now() , oTodaydealList.FItemList(i).Fenddate )) >= 0 Then ' ����
				Response.write " <span style=""color:red"">(�� "& cInt(datediff("h", now() , oTodaydealList.FItemList(i).Fenddate )) &" �ð��� ����)</span>"
				Else  ' ������
				Response.write " <span style=""color:red"">(������)</span>"					
				End If 

			End If 
		%>
	</td>
	<td>
		<%=getStaffUserName(oTodaydealList.FItemList(i).Fadminid)%><br/>
		<%
			modiTime = oTodaydealList.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write "(���� : " & getStaffUserName(oTodaydealList.FItemList(i).Flastadminid) & " " & left(modiTime,10) & ")"
			end if
		%>
	</td>
	<td><%=chkiif(oTodaydealList.FItemList(i).Ftopfixed = "Y","����","�����")%></td>
    <td><%=oTodaydealList.FItemList(i).Fsortnum%></td>
    <td><%=chkiif(oTodaydealList.FItemList(i).Fisusing="N","������","�����")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td colspan="11" align="center">
		<% if oTodaydealList.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oTodaydealList.StarScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i = 0 + oTodaydealList.StartScrollPage to oTodaydealList.StartScrollPage + oTodaydealList.FScrollCount - 1 %>
			<% if (i > oTodaydealList.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oTodaydealList.FCurrPage) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oTodaydealList.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<%
set oTodaydealList = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->