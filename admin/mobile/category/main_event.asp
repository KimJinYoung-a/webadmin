<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/category/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/mo_catetoryMainManageCls.asp" -->
<%
'###########################################################
' Description :  ����� ī�װ� ���� �̺�Ʈ
' History : 2020.12.02 ������ ����
'###########################################################
	
	Dim view_yn , dispcate , validdate , research
	dim page 
	Dim i
	dim oEventinfo
	Dim sDt , modiTime , sedatechk
	Dim addtype, catecode

	page = request("page")
	dispcate = request("disp")
	view_yn = RequestCheckVar(request("view_yn"),13)
	sDt = request("prevDate")
	validdate= request("validdate")
	research= request("research")
	sedatechk = request("sedatechk")
	addtype = request("addtype")
    catecode = request("catecode")

	if ((research="") and (view_yn="")) then
	    view_yn = "1"
	    validdate = "on"
	end if
	
	if page="" then page=1

	set oEventinfo = new CMainContents
	oEventinfo.FPageSize = 20
	oEventinfo.FCurrPage = page
	oEventinfo.FRectIsusing = view_yn
	oEventinfo.Fsdt = sDt
	oEventinfo.FRectvaliddate = validdate
    oEventinfo.FRectCatecode = catecode
	oEventinfo.FRectsedatechk = sedatechk '//������ ���� üũ
	oEventinfo.GetEventContentsList()
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
function addContents(idx){
	var dateOptionParam
	var frm = document.frm
	dateOptionParam = frm.prevDate.value

    var popwin = window.open('eventinfo_insert.asp?idx=' + idx + '&dateoption=' + dateOptionParam,'mainposcodeedit','width=800,height=500,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function tnCheckAll(bool, comp){
    var frm = comp.form;
    if (!comp.length){
        comp.checked = bool;
    }else{
        for (var i=0;i<comp.length;i++){
            comp[i].checked = bool;
        }
    }
}

function fnOrderidxEdit(){
var itemcount = 0;
var frm;
var ck=0;
frm = document.frmArrupdate;

	if(typeof(frm.cksel) !="undefined"){
		if(!frm.cksel.length){
			if(!frm.cksel.checked){
				alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
				return;
			}
				frm.orderidxarr.value = frm.cksel.value;
		}else{
			for(i=0;i<frm.cksel.length;i++){
				if(frm.cksel[i].checked) {
					ck=ck+1;	   	    			
					if (frm.orderidxarr.value==""){
						frm.idxarr.value =  frm.cksel[i].value;
						frm.orderidxarr.value =  frm.orderidx[i].value;
					}else{
						frm.idxarr.value = frm.idxarr.value + "," +frm.cksel[i].value;
						frm.orderidxarr.value = frm.orderidxarr.value + "," +frm.orderidx[i].value;
					} 
				}	
			}
			
			if (frm.orderidxarr.value == ""){
				alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
				return;
			}
		}
	}else{
		alert("�߰��� ��ǰ�� �����ϴ�.");
		return;
	}
	frm.target = "FrameCKP";
	frm.action = "doEventOrderidx.asp";
	frm.submit();
}
-->
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<div>
			<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������&nbsp;
			&nbsp;ī�װ�&nbsp;
            <% DrawSelectBoxDispCateLarge "catecode", catecode, "" %>
            &nbsp;* ��뿩�� :&nbsp;
                <select name="view_yn" class="select">
                <option value="">��ü
                <option value="1" <% if view_yn="1" then response.write "selected" %> >�����
                <option value="0" <% if view_yn="0" then response.write "selected" %> >������
                </select>&nbsp;&nbsp;&nbsp;
			�����ϱ��� <input type="checkbox" name="sedatechk" <% if sedatechk="on" then response.write "checked" %> />
			�������� <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			&nbsp;
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
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
	<td align="left">
    	<input type="button" value="���� ���� ����" onClick="fnOrderidxEdit();"/>
    </td>
    <td align="right">
    	<a href="javascript:addContents(0);"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  ����Ʈ -->
<form name="frmArrupdate" method="post">
<input type="hidden" name="idxarr">
<input type="hidden" name="orderidxarr">
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		�� ��ϼ� : <b><%=oEventinfo.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oEventinfo.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="validdate" onClick="tnCheckAll(this.checked,frmArrupdate.cksel);" ></td>
    <td width="5%">idx</td>
	<td width="10%">ī�װ�</td>
    <td width="20%">�̺�Ʈ�ڵ�</td>	 
    <td width="20%">�̺�Ʈ��</td>
    <td width="15%">������/������</td>
    <td width="10%">�����</td>
    <td width="10%">�����</td>
	<td width="5%">�켱����</td>
    <td width="10%">��뿩��</td>
</tr>
<% 
	for i=0 to oEventinfo.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oEventinfo.FItemList(i).Fview_yn="1","#FFFFFF","#F0F0F0")%>">
	<td align="center"><input type="checkbox" name="cksel" id="cksel" value="<%= oEventinfo.FItemList(i).Fidx %>"></td>
    <td onclick="addContents('<%=oEventinfo.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=oEventinfo.FItemList(i).Fidx%></td>
    <td onclick="addContents('<%=oEventinfo.FItemList(i).Fidx%>');" style="cursor:pointer;"><a href="javascript:AddNewMainContents('<%= oEventinfo.FItemList(i).Fidx %>');"><%= oEventinfo.FItemList(i).Fcatename %></a></td>
    <td onclick="addContents('<%=oEventinfo.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=oEventinfo.FItemList(i).Fevt_code%></td>
    <td onclick="addContents('<%=oEventinfo.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=oEventinfo.FItemList(i).Fevt_name%></td>
	<td>
		<% 
			Response.Write "����: "
			Response.Write replace(left(oEventinfo.FItemList(i).Fstart_date,10),"-",".") & " / " & Num2Str(hour(oEventinfo.FItemList(i).Fstart_date),2,"0","R") & ":" &Num2Str(minute(oEventinfo.FItemList(i).Fstart_date),2,"0","R")
			Response.Write "<br />����: "
			Response.Write replace(left(oEventinfo.FItemList(i).Fend_date,10),"-",".") & " / " & Num2Str(hour(oEventinfo.FItemList(i).Fend_date),2,"0","R") & ":" & Num2Str(minute(oEventinfo.FItemList(i).Fend_date),2,"0","R")
		%>
	</td>
	<td><%=left(oEventinfo.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(oEventinfo.FItemList(i).Freguserid)%></td>
	<td align="center">
		<input type="text" name="orderidx" class="formTxt" size=5 value="<%= oEventinfo.FItemList(i).fview_order %>">
    </td>
    <td><%=chkiif(oEventinfo.FItemList(i).Fview_yn="0","������","�����")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td colspan="11" align="center">
		<% if oEventinfo.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oEventinfo.StarScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i = 0 + oEventinfo.StarScrollPage to oEventinfo.StarScrollPage + oEventinfo.FScrollCount - 1 %>
			<% if (i > oEventinfo.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oEventinfo.FCurrPage) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oEventinfo.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
</form>
<%
set oEventinfo = Nothing
%>
<iframe name="FrameCKP" src="" frameborder="0" width="600" height="400"></iframe>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->