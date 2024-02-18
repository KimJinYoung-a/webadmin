<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : ����� keywordbanner
' History : 2013.12.16 �ѿ��
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/mobile/main/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/keywordbanner_cls.asp" -->
<%
Dim isusing, page, i, okeyword, reload ,ndate
	page = request("page")
	reload = request("reload")
	ndate = request("prevDate")
	isusing = RequestCheckVar(request("isusing"),1)

if page="" then page=1
if reload="" and isusing="" then isusing="Y"

set okeyword = new ckeywordbanner
	okeyword.FPageSize			= 20
	okeyword.FCurrPage		= page
	okeyword.Frectisusing		= isusing
	okeyword.Frectdate			= ndate
	okeyword.getkeywordbanner_list()

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>
$(function(){
	 $(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
	<% if ndate <> "" then %>
	$( "#subList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).parent().find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).parent().find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
	<% end if %>
});

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

function keywordbanneredit(idx){
	var keywordbanneredit = window.open('/admin/mobile/keywordbanner/keywordbanner_edit.asp?idx='+idx+'&menupos=<%=menupos%>','keywordbanneredit','width=1024,height=768,scrollbars=yes,resizable=yes');
	keywordbanneredit.focus();
}

function RefreshCaFavKeyWordRec(term){
	if(confirm("�����- KEYWORDBANNER�� �����Ͻðڽ��ϱ�?")) {
			var popwin = window.open('','refreshFrm','');
			popwin.focus();
			refreshFrm.target = "frm";
			refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_main_KeyWordBanner_new_xml.asp?term=" + term;
			refreshFrm.submit();
	}
}

// �̹��� Ŭ���� ���� ũ��� �˾� ����
function doImgPop(img){
	img1= new Image();
	img1.src=(img);
	imgControll(img);
}
function imgControll(img){
	if((img1.width!=0)&&(img1.height!=0)){
		viewImage(img);
	}else{
		controller="imgControll('"+img+"')";
		intervalID=setTimeout(controller,20);
	}
}
function viewImage(img){
	W=img1.width;
	H=img1.height;
	O="width="+W+",height="+H+",scrollbars=yes";
	imgWin=window.open("","",O);
	imgWin.document.write("<html><head><title>:*:*:*: �̹����󼼺��� :*:*:*:*:*:*:</title></head>");
	imgWin.document.write("<body topmargin=0 leftmargin=0>");
	imgWin.document.write("<img src="+img+" onclick='self.close()' style='cursor:pointer;' title ='Ŭ���Ͻø� â�� �����ϴ�.'>");
	imgWin.document.close();
}

function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� ���縦 �������ּ���.");
		return;
	}
	if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
		document.frmList.action="doListModify.asp";
		document.frmList.submit();
	}
}
</script>

<img src="/images/icon_arrow_link.gif"> <b>KEYWORDBANNER</b>
<p>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="ON">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ��뿩�� : <% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		�������� <input id="prevDate" name="prevDate" value="<%=ndate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script type="text/javascript">
			var CAL_Start = new Calendar({
				inputField : "prevDate", trigger    : "prevDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:frmsubmit('');">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td>
		������ �����Ͽ� <input type="text" name="vTerm" value="1" size="1" class="text" style="text-align:right;">�ϰ�<a href="javascript:RefreshCaFavKeyWordRec(document.all.vTerm.value);"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html�����"></a>XML Real ����(����)</a>&nbsp;&nbsp;&nbsp;
		<a onclick="doImgPop('/admin/mobile/keywordbanner/cc.JPG')" style="cursor:pointer;"><font color="RED">���ú���</font></a>
	</td>
    <td align="right">
    	<input type="button" onclick="keywordbanneredit('')" value="�űԵ��" class="button">
    </td>
</tr>
</table>
<!--  ����Ʈ -->
<form name="frmList" method="POST" style="margin:0;">
<input type="hidden" name="mode" value="sub">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�� ��ϼ� : <b><%=okeyword.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=okeyword.FtotalPage%></b>
		&nbsp;&nbsp;&nbsp;
		<% If ndate <> "" Then %>
		<input type="button" value="��ü����" class="button" onClick="chkAllItem()">
		<input type="button" value="��������" class="button" onClick="saveList()" title="ǥ�ü��� �� ��뿩�θ� �ϰ������մϴ�.">
		<% End If %>
	</td>
</tr>
<col width="30" />
<col width="80" />
<col span="4" width="0*" />
<col width="150" />
<col width="100" />
<col width="80" />
<col width="80" />
<col width="80" />
<tr align="center" bgcolor="#DDDDFF">
    <td>&nbsp;</td>
    <td>��ʱ���</td>	 
	<td>������ <br/>real ����ð�</td>
    <td>Ű���� �� �̹���</td>
	<td>������/������</td>
	<td>�����</td>
    <td>���ļ���</td>
    <td>��뿩��</td>
    <td>���</td>
</tr>
<tbody id="subList">
<%
if okeyword.FResultCount>0 then
	
for i=0 to okeyword.FResultCount - 1 
%>

<tr height="30" align="center" bgcolor="<%=chkIIF(okeyword.FItemList(i).fisusing="Y","#FFFFFF","#F0F0F0")%>">
	<td><input type="checkbox" name="chkIdx" value="<%=okeyword.FItemList(i).Fidx%>" /></td>
    <td><%= okeyword.FItemList(i).fkeywordtypename %>(<%= okeyword.FItemList(i).fkeywordtype %>)</td>
	<td>
		<%
			If okeyword.FItemList(i).Fxmlregdate <> "" then
			Response.Write replace(left(okeyword.FItemList(i).Fxmlregdate,10),"-",".") & " <br/> " & Num2Str(hour(okeyword.FItemList(i).Fxmlregdate),2,"0","R") & ":" &Num2Str(minute(okeyword.FItemList(i).Fxmlregdate),2,"0","R")
			End If 
		%>
	</td>
    <td>
    	<% If okeyword.FItemList(i).fkeywordtype = "1" Then %>
    	<img src="<%= okeyword.FItemList(i).fimagepath %>" width=50 height=50 />
    	<%
    	   ElseIf okeyword.FItemList(i).fkeywordtype = "2" Then 
    		response.write okeyword.FItemList(i).fkeyword 
    	   End If
    	%>
    </td>
	<td>
		<%
			If okeyword.FItemList(i).Fstartdate <> "" And okeyword.FItemList(i).Fenddate <> "" Then 
				Response.Write "����: "
				Response.Write replace(left(okeyword.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(okeyword.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(okeyword.FItemList(i).Fstartdate),2,"0","R")
				Response.Write "<br />����: "
				Response.Write replace(left(okeyword.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(okeyword.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(okeyword.FItemList(i).Fenddate),2,"0","R")
			End If 
		%>
	</td>
	<td><%=Left(okeyword.FItemList(i).fregdate,10)%></td>
	<td>
		<input type="text" name="sort<%=okeyword.FItemList(i).Fidx%>" size="3" class="text" value="<%=okeyword.FItemList(i).forderno%>" style="text-align:center;" />
	</td>
	<td>
		<span class="rdoUsing">
		<input type="radio" name="isusing<%=okeyword.FItemList(i).Fidx%>" id="rdoUsing<%=i%>_1" value="Y" <%=chkIIF(okeyword.FItemList(i).fisusing="Y","checked","")%> /><label for="rdoUsing<%=i%>_1">���</label><input type="radio" name="isusing<%=okeyword.FItemList(i).Fidx%>" id="rdoUsing<%=i%>_2" value="N" <%=chkIIF(okeyword.FItemList(i).fisusing="N","checked","")%> /><label for="rdoUsing<%=i%>_2">����</label>
		</span>
	</td>

	<td>
		<input type="button" onclick="keywordbanneredit('<%=okeyword.FItemList(i).Fidx%>')" value="����" class="button">
	</td>
</tr>
<% Next %>

<tr bgcolor="#FFFFFF">
	<td align="center" colspan="20">
		<% if okeyword.HasPreScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= okeyword.StartScrollPage-1 %>')">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + okeyword.StartScrollPage to okeyword.StartScrollPage + okeyword.FScrollCount - 1 %>
			<% if (i > okeyword.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(okeyword.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:frmsubmit('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if okeyword.HasNextScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= i %>')">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</form>
</table>
</tbody>
<%
set okeyword = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->