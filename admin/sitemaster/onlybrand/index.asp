<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : PC���ΰ��� �´Ϻ귣��
' History : ������ ����
'			2022.07.01 �ѿ�� ����(isms�������ġ)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/pcmain/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/onlybrandCls.asp" -->
<%
	Dim isusing , dispcate
	dim page 
	Dim i
	dim onlyBrandList
	Dim sDt , modiTime

	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")

	if page="" then page=1
	If isusing = "" Then isusing ="Y"

	set onlyBrandList = new ConlyBrand
	onlyBrandList.FPageSize		= 20
	onlyBrandList.FCurrPage		= page
	onlyBrandList.Fisusing			= isusing
	onlyBrandList.Fsdt					= sDt
	onlyBrandList.GetContentsList()

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
	location.href = "onlyBrand_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>&paramisusing=<%=isusing%>";
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
-->
</script>
<!-- �˻� ���� -->
<form name="frm" method="post" style="margin:0px;" action="/admin/sitemaster/onlybrand/index.asp">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
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
			</div>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="submit" class="button_s" value="�� ��">
		</td>
	</tr>
</table>
</form>
<!-- �˻� �� -->
<br>
<div style="float:right;clear:both;"><a href="onlybrand_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>&paramisusing=<%=isusing%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a></div>
<br><br>
<!--  ����Ʈ -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		�� ��ϼ� : <b><%=onlyBrandList.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=onlyBrandList.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
	<td width="22%">����̹���</td>	 
    <td width="18%">����ī��</td>
    <td width="18%">������/������</td>
    <td width="5%">�����</td>
    <td width="10%">�����</td>
    <td width="10%">����������</td>
    <td width="5%">���Ĺ�ȣ</td>
    <td width="10%">��뿩��</td>
</tr>
<% 
	for i=0 to onlyBrandList.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(onlyBrandList.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=onlyBrandList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=onlyBrandList.FItemList(i).Fidx%></td>
    <td onclick="jsmodify('<%=onlyBrandList.FItemList(i).Fidx%>');" style="cursor:pointer;"><img src="<%=onlyBrandList.FItemList(i).FbannerImg%>" width="300"></td>
    <td onclick="jsmodify('<%=onlyBrandList.FItemList(i).Fidx%>');" style="cursor:pointer;"><%= ReplaceBracket(onlyBrandList.FItemList(i).Fmaincopy) %></td>
	<td onclick="jsmodify('<%=onlyBrandList.FItemList(i).Fidx%>');" style="cursor:pointer;" align="left">
		<% 
			If onlyBrandList.FItemList(i).Fstartdate <> "" And onlyBrandList.FItemList(i).Fenddate Then 
				Response.Write "����: "
				Response.Write replace(left(onlyBrandList.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(onlyBrandList.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(onlyBrandList.FItemList(i).Fstartdate),2,"0","R")
				Response.Write "<br />����: "
				Response.Write replace(left(onlyBrandList.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(onlyBrandList.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(onlyBrandList.FItemList(i).Fenddate),2,"0","R")
				If clng(datediff("d", now() , onlyBrandList.FItemList(i).Fenddate)) < 0 Or clng(datediff("h", now() , onlyBrandList.FItemList(i).Fenddate )) < 0  Then 
					Response.write " <span style=""color:red"">(����)</span>"
				ElseIf clng(datediff("d", onlyBrandList.FItemList(i).Fenddate , now())) = 0  Then '�������� ���ó�¥
					If clng(datediff("h", now() , onlyBrandList.FItemList(i).Fenddate )) >= 0 And clng(datediff("h", now() , onlyBrandList.FItemList(i).Fenddate )) < 24 Then ' ����
						Response.write " <span style=""color:red"">(�� "& clng(datediff("h", now() , onlyBrandList.FItemList(i).Fenddate )) &" �ð��� ����)</span>"
					Else  ' ������
						Response.write " <span style=""color:red"">(������)</span>"					
					End If 
				'// �������� ���ó�¥�̰� �������� ������ �ƴϸ�
				ElseIf clng(datediff("d", onlyBrandList.FItemList(i).Fstartdate , now()))>=0 And clng(datediff("d", onlyBrandList.FItemList(i).Fenddate , now())) < 0 Then
					Response.write " <span style=""color:red"">(�� "&clng(datediff("d", now() , onlyBrandList.FItemList(i).Fenddate ))&"�� " &clng(datediff("h", now() , onlyBrandList.FItemList(i).Fenddate ))-clng(datediff("d", now() , onlyBrandList.FItemList(i).Fenddate ))*24 &"�ð��� ����)</span>"
				ElseIf clng(datediff("d", onlyBrandList.FItemList(i).Fenddate , now())) < 0 Then
					If clng(datediff("h", now() , onlyBrandList.FItemList(i).Fenddate )) >= 0 And clng(datediff("h", now() , onlyBrandList.FItemList(i).Fenddate )) < 24 Then ' ����
						Response.write " <span style=""color:red"">(�� "& clng(datediff("h", now() , onlyBrandList.FItemList(i).Fenddate )) &" �ð��� ����)</span>"
					Else  ' ������
						Response.write " <span style=""color:red"">(������)</span>"
					End If 
				End If
			End If 
		%>
	</td>
	<td><%=left(onlyBrandList.FItemList(i).Fregdate,10)%></td>
	<td><%=onlyBrandList.FItemList(i).Fusername%></td>
	<td>
		<%
			modiTime = onlyBrandList.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write onlyBrandList.FItemList(i).Fusername2 & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
    <td><%=onlyBrandList.FItemList(i).Forderby%></td>
    <td><%=chkiif(onlyBrandList.FItemList(i).Fisusing="N","������","�����")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if onlyBrandList.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= onlyBrandList.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + onlyBrandList.StartScrollPage to onlyBrandList.StartScrollPage + onlyBrandList.FScrollCount - 1 %>
				<% if (i > onlyBrandList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(onlyBrandList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if onlyBrandList.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set onlyBrandList = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->