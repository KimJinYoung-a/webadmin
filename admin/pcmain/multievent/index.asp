<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/pcmain/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/pcmain_multieventCls.asp" -->
<%
'###############################################
' PageName : index.asp
' Discription : PC ���� enjoybanner
' History : 2018-03-14 ����ȭ
'		  : 2018-11-26 pc���� ��� ��ȹ�� ���� ���� �߰�
'###############################################

	Dim isusing , dispcate , validdate , research, mode
	dim page
	Dim i
	dim oMultiEventList
	Dim sDt , modiTime , sedatechk , prevTime
	dim dispOption	' "" : ����, 1 : ���� ��ܱ�ȹ��

	dispOption = request("dispOption")
	page = request("page")
	dispcate = request("disp")
	isusing = RequestCheckVar(request("isusing"),13)
	sDt = request("prevDate")
	validdate= request("validdate")
	research= request("research")
	sedatechk = request("sedatechk")
	prevTime = request("prevTime")
	mode = RequestCheckVar(request("mode"),5)
	
	if ((research="") and (isusing="")) then
	    isusing = "Y"
	    validdate = "on"
	end If

	if dispOption = "" then dispOption = 1
	if prevTime = "" then prevTime = "00"
	if page="" then page=1

	set oMultiEventList = new CMainbanner
	oMultiEventList.FPageSize			= 20
	oMultiEventList.FCurrPage			= page
	oMultiEventList.Fisusing			= isusing
	oMultiEventList.Fsdt				= sDt
	oMultiEventList.FRectvaliddate		= validdate
	oMultiEventList.FRectsedatechk		= sedatechk '//������ ���� üũ
	oMultiEventList.FRectSelDateTime	= prevTime 
	oMultiEventList.FRectDispOption		= dispOption
	oMultiEventList.GetContentsList()
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
function jsmodify(v, contentType){
	if(contentType == 2){
		location.href = "item_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
	}else{
		location.href = "event_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
	}
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
function controlExhibition(){			
	var popwin; 		
	popwin = window.open("exhibition_ctrl.asp", "popup_item", "width=500,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}
function addContents(){
	var dispOption = document.frm.dispOption.value;	
	if(dispOption == "2"){
		document.location.href="event_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>&dispOption="+dispOption		
	}else{
		document.location.href="event_insert.asp?menupos=<%=menupos%>&prevDate=<%=sDt%>"
	}
}

//��ü����
var ichk;
ichk = 1;
	
function jsChkAll(){			
	    var frm, blnChk;
		frm = document.fitem;
		if(!frm.chkI) return;
		if ( ichk == 1 ){
			blnChk = true;
			ichk = 0;
		}else{
			blnChk = false;
			ichk = 1;
		}
		
 		for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA		
		if ((e.type=="checkbox")) {				
			e.checked = blnChk ;
		}
	}
}
//�ϰ� ����
function fnTrendEventCopy() {
	var frm;
	var sValue, sSort, sImgSize,sUsing, sSort_mo, sImgSize_mo,sUsing_mo,sDisp;
	frm = document.fitem;
	sValue = "";
	sSort = ""; 
	sDisp = ""
	var itemid;	

		for (var i=0;i<frm.chkI.length;i++){ 
			if (frm.chkI[i].checked){
				itemid = frm.chkI[i].value;		
				if (sValue==""){
					sValue = frm.chkI[i].value;		
				}else{
					sValue =sValue+","+frm.chkI[i].value;		
				}
			}
		}
		if (sValue == "") {
			alert('���� �������� �����ϴ�.');
			return;
		}
		frm.idxarr.value = sValue;
		frm.submit();

}
-->

function popTodayEasyReg(){
    let popTodayEasyReg = window.open('/admin/mobile/popTodayEasyReg.asp?type=event','mainposcodeedit','width=800,height=400,scrollbars=yes,resizable=yes');
    popTodayEasyReg.focus();
}

</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="mode" value="<%=mode%>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������&nbsp;
			* ��뿩�� :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			�����ϱ��� <input type="checkbox" name="sedatechk" <% if sedatechk="on" then response.write "checked" %> />
			�������� <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<% if sDt <> "" then %>
			�ð� <input type="input" name="prevTime" value="<%=prevTime%>" class="text" size="2" maxlength="2" /> ��~
			<% end if %>
			&nbsp;
			&nbsp; ���� ��ġ : 
			<select name="dispOption" class="select" onchange="javascript:submit();">				
				<option value="1" <%=chkiif(dispOption="1"," selected","")%>>�⺻</option>
				<option value="2" <%=chkiif(dispOption="2"," selected","")%>>���λ�ܱ�ȹ��</option>
			</select>			
			<script type="text/javascript">
				var CAL_Start = new Calendar({
					inputField : "prevDate", trigger    : "prevDate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
			
			</div>
		</td>		
		<td width="120" bgcolor="<%= adminColor("gray") %>">
			<button sytle="float:left" type="button" onclick="controlExhibition();">���λ�ܱ�ȹ������</button>
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
	<% if mode="copy" then %>
	<td align="left"><button onClick="fnTrendEventCopy();">���� ����</button>&nbsp;&nbsp;</td>
	<% end if %>
    <td align="right">
        <input type="button" class="button" value="������" onClick="popTodayEasyReg();" />
		<!-- �űԵ�� -->
    	<a href="javascript:void(0)" onclick="addContents()"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  ����Ʈ -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		�� ��ϼ� : <b><%=oMultiEventList.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oMultiEventList.FtotalPage%></b>
	</td>
</tr>
<% if mode="copy" then %>
<form name="fitem" method="post" action="docopytrendevent.asp">
<input type="hidden" name="idxarr" value="">
<% end if %>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if mode="copy" then %>
	<td width="5%"><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
	<td width="5%">idx</td>
	<td width="5%">���� ��ġ</td>
	<% else %>
	<td width="5%">idx</td>
	<td width="10%">���� ��ġ</td>
	<% end if %>
	<td width="20%">����̹���</td>
    <td width="15%">������/������</td>
    <td width="10%">�����</td>
    <td width="10%">�����</td>
    <td width="10%">����������</td>
    <td width="10%">�켱����</td>
    <td width="10%">��뿩��</td>	
</tr>
<%
	for i=0 to oMultiEventList.FResultCount-1
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oMultiEventList.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
	<% if mode="copy" then %>
	<td><input type="checkbox" name="chkI" value="<%=oMultiEventList.FItemList(i).Fidx%>"></td>
	<% end if %>
    <td style="cursor:pointer;" onclick="jsmodify('<%=oMultiEventList.FItemList(i).Fidx%>','<%=oMultiEventList.FItemList(i).FcontentType%>');return false;"><%=oMultiEventList.FItemList(i).Fidx%><p>&nbsp;</p>
		<!--<a href="" onclick="window.open('enjoy_preview.asp?idx=<%=oMultiEventList.FItemList(i).Fidx%>','enjoypreview', 'width=733, height=900');return false;">[�̸�����]</a>-->
	</td>
	<td>
		<%	
			select case oMultiEventList.FItemList(i).FdispOption
				case 1
					response.write "�⺻"
				case 2
					response.write "���λ�ܱ�ȹ��"
			end select				
		%>
	</td>
    <td>
		<% if oMultiEventList.FItemList(i).FcontentType = 2 then %>
			<img src="<%=oMultiEventList.FItemList(i).FcontentImg%>" width="200" height="90" alt="<%=oMultiEventList.FItemList(i).Fmaincopy%>"/>
		<% else %>
			<img src="<%=oMultiEventList.FItemList(i).Fevtmolistbanner%>" width="200" height="90" alt="<%=oMultiEventList.FItemList(i).Fmaincopy%>"/>
		<% end if %>
		
	</td>
	<td>
		<%
			Response.Write "����: "
			Response.Write replace(left(oMultiEventList.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(oMultiEventList.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(oMultiEventList.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />����: "
			Response.Write replace(left(oMultiEventList.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(oMultiEventList.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(oMultiEventList.FItemList(i).Fenddate),2,"0","R")
		%>
	</td>
	<td><%=left(oMultiEventList.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(oMultiEventList.FItemList(i).Fadminid)%></td>
	<td>
		<%
			modiTime = oMultiEventList.FItemList(i).Flastupdate
			if Not(modiTime="" or isNull(modiTime)) then
					Response.Write getStaffUserName(oMultiEventList.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(modiTime,10)
			end if
		%>
	</td>
    <td><%=oMultiEventList.FItemList(i).Fsortnum%></td>
    <td><%=chkiif(oMultiEventList.FItemList(i).Fisusing="N","������","�����")%></td>
</tr>
<% Next %>
<% if mode="copy" then %>
</form>
<% end if %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="center">
		<% if oMultiEventList.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oMultiEventList.StarScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i = 0 + oMultiEventList.StartScrollPage to oMultiEventList.StartScrollPage + oMultiEventList.FScrollCount - 1 %>
			<% if (i > oMultiEventList.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oMultiEventList.FCurrPage) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oMultiEventList.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<%
set oMultiEventList = Nothing
%>
<form name="refreshFrm" method="post">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->