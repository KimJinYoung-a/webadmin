<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : index.asp
' Discription : ����� ��ǰ���
' History : 2014.06.23 ����ȭ ����
'			2021.02.22 �ѿ�� ����(�ҽ� ǥ�ر԰����� ����. ����üũ �߰�. �������� ����)
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/today_twinitemsCls.asp" -->
<%
Dim isusing, validdate , research, page, i, oEnjoyeventlist, sDt , modiTime , sedatechk, addtype, dispcate
Dim L_img , R_img , ii
	page = requestCheckVar(getNumeric(request("page")),10)
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = requestCheckVar(request("prevDate"),10)
	validdate= requestCheckVar(request("validdate"),2)
	research= requestCheckVar(request("research"),2)
	sedatechk = requestCheckVar(request("sedatechk"),2)
	dispcate = request("disp")
	addtype = request("addtype")

if ((research="") and (isusing="")) then
	isusing = "Y"
	validdate = "on"
end if

if page="" then page=1

set oEnjoyeventlist = new CMainbanner
	oEnjoyeventlist.FPageSize			= 20
	oEnjoyeventlist.FCurrPage			= page
	oEnjoyeventlist.Fisusing			= isusing
	oEnjoyeventlist.Fsdt				= sDt
	oEnjoyeventlist.FRectvaliddate		= validdate
	oEnjoyeventlist.FRectsedatechk		= sedatechk '//������ ���� üũ
	oEnjoyeventlist.GetContentsList()

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
	location.href = "twinitems_insert.asp?menupos=<%=menupos%>&idx="+v;
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

function addContents(){
	var dateOptionParam
	var frm = document.frm
	dateOptionParam = frm.prevDate.value

	document.location.href="twinitems_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>&dateoption="+dateOptionParam
}
-->

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<div style="padding-bottom:10px;">
		* ��뿩�� : <% DrawSelectBoxUsingYN "isusing",isusing %>
		&nbsp;
		* �����ϱ��� <input type="checkbox" name="sedatechk" <% if sedatechk="on" then response.write "checked" %> />
		&nbsp;
		�������� <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script type="text/javascript">
			var CAL_Start = new Calendar({
				inputField : "prevDate", trigger    : "prevDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		&nbsp;
		<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������
		</div>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:submit();">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<Br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;" >
<tr>
	<td align="left"></td>
	<td align="right">	
		<input type="button" class="button" value="�űԵ��" onclick="addContents();">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!--  ����Ʈ -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		�� ��ϼ� : <b><%=oEnjoyeventlist.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oEnjoyeventlist.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
	<td width="20%">����̹���</td>	 
    <td width="15%">������/������</td>
    <td width="10%">�����</td>
    <td width="10%">�����</td>
    <td width="10%">����������</td>
    <td width="10%">��뿩��</td>
</tr>
<% if oEnjoyeventlist.FResultCount>0 then %>
<% 
	
	for i=0 to oEnjoyeventlist.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oEnjoyeventlist.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=oEnjoyeventlist.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=oEnjoyeventlist.FItemList(i).Fidx%></td>
    <td>
		<%
			L_img =  oEnjoyeventlist.FItemList(i).FL_img
			R_img =  oEnjoyeventlist.FItemList(i).FR_img
			if not isnull(oEnjoyeventlist.FItemList(i).Fiteminfo) then 
				If ubound(Split(oEnjoyeventlist.FItemList(i).Fiteminfo,"^^")) > 0 Then ' �̹��� 3�� ����
					For ii = 0 To ubound(Split(oEnjoyeventlist.FItemList(i).Fiteminfo,","))
						If CStr(oEnjoyeventlist.FItemList(i).FL_itemid) = CStr(Split(Split(oEnjoyeventlist.FItemList(i).Fiteminfo,",")(ii),"|")(0)) And oEnjoyeventlist.FItemList(i).FL_img = (staticImgUrl & "/mobile/twinitems") Then
							L_img =  webImgUrl & "/image/icon1/" & GetImageSubFolderByItemid(oEnjoyeventlist.FItemList(i).FL_itemid) & "/" & Split(Split(oEnjoyeventlist.FItemList(i).Fiteminfo,",")(ii),"|")(2)
						End If

						If CStr(oEnjoyeventlist.FItemList(i).FR_itemid) = CStr(Split(Split(oEnjoyeventlist.FItemList(i).Fiteminfo,",")(ii),"|")(0)) And oEnjoyeventlist.FItemList(i).FR_img = (staticImgUrl & "/mobile/twinitems") Then
							R_img =  webImgUrl & "/image/icon1/" & GetImageSubFolderByItemid(oEnjoyeventlist.FItemList(i).FR_itemid) & "/" & Split(Split(oEnjoyeventlist.FItemList(i).Fiteminfo,",")(ii),"|")(2)
						End If
					Next 
				End If
			end if
		%>
		<img src="<%=L_img%>" width="100" height="100" alt="<%=oEnjoyeventlist.FItemList(i).FL_itemname%>"/>
		<img src="<%=R_img%>" width="100" height="100" alt="<%=oEnjoyeventlist.FItemList(i).FR_itemname%>"/>
	</td>
	<td>
		<% 
			Response.Write "����: "
			Response.Write replace(left(oEnjoyeventlist.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(oEnjoyeventlist.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(oEnjoyeventlist.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />����: "
			Response.Write replace(left(oEnjoyeventlist.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(oEnjoyeventlist.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(oEnjoyeventlist.FItemList(i).Fenddate),2,"0","R")
		%>
	</td>
	<td><%=left(oEnjoyeventlist.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(oEnjoyeventlist.FItemList(i).Fadminid)%></td>
	<td>
		<%
			modiTime = oEnjoyeventlist.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write getStaffUserName(oEnjoyeventlist.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
    <td><%=chkiif(oEnjoyeventlist.FItemList(i).Fisusing="N","������","�����")%></td>
</tr>
<% Next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if oEnjoyeventlist.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oEnjoyeventlist.StartScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i = 0 + oEnjoyeventlist.StartScrollPage to oEnjoyeventlist.StartScrollPage + oEnjoyeventlist.FScrollCount - 1 %>
			<% if (i > oEnjoyeventlist.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oEnjoyeventlist.FCurrPage) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oEnjoyeventlist.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set oEnjoyeventlist = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->