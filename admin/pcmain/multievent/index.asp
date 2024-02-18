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
' Discription : PC 메인 enjoybanner
' History : 2018-03-14 이종화
'		  : 2018-11-26 pc메인 상단 기획전 구좌 노출 추가
'###############################################

	Dim isusing , dispcate , validdate , research, mode
	dim page
	Dim i
	dim oMultiEventList
	Dim sDt , modiTime , sedatechk , prevTime
	dim dispOption	' "" : 기존, 1 : 메인 상단기획전

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
	oMultiEventList.FRectsedatechk		= sedatechk '//시작일 기준 체크
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
//수정
function jsmodify(v, contentType){
	if(contentType == 2){
		location.href = "item_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
	}else{
		location.href = "event_insert.asp?menupos=<%=menupos%>&idx="+v+"&prevDate=<%=sDt%>";
	}
}
$(function() {
  	$("input[type=submit]").button();

  	// 라디오버튼
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

//전체선택
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
//일괄 복사
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
			alert('선택 컨텐츠가 없습니다.');
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
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="mode" value="<%=mode%>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전&nbsp;
			* 사용여부 :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			시작일기준 <input type="checkbox" name="sedatechk" <% if sedatechk="on" then response.write "checked" %> />
			지정일자 <input id="prevDate" name="prevDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<% if sDt <> "" then %>
			시간 <input type="input" name="prevTime" value="<%=prevTime%>" class="text" size="2" maxlength="2" /> 시~
			<% end if %>
			&nbsp;
			&nbsp; 노출 위치 : 
			<select name="dispOption" class="select" onchange="javascript:submit();">				
				<option value="1" <%=chkiif(dispOption="1"," selected","")%>>기본</option>
				<option value="2" <%=chkiif(dispOption="2"," selected","")%>>메인상단기획전</option>
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
			<button sytle="float:left" type="button" onclick="controlExhibition();">메인상단기획전관리</button>
		</td>		
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:submit();">
		</td>
	</tr>
</form>
</table>
<!-- 검색 끝 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<% if mode="copy" then %>
	<td align="left"><button onClick="fnTrendEventCopy();">선택 복사</button>&nbsp;&nbsp;</td>
	<% end if %>
    <td align="right">
        <input type="button" class="button" value="간편등록" onClick="popTodayEasyReg();" />
		<!-- 신규등록 -->
    	<a href="javascript:void(0)" onclick="addContents()"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  리스트 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		총 등록수 : <b><%=oMultiEventList.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oMultiEventList.FtotalPage%></b>
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
	<td width="5%">노출 위치</td>
	<% else %>
	<td width="5%">idx</td>
	<td width="10%">노출 위치</td>
	<% end if %>
	<td width="20%">등록이미지</td>
    <td width="15%">시작일/종료일</td>
    <td width="10%">등록일</td>
    <td width="10%">등록자</td>
    <td width="10%">최종수정자</td>
    <td width="10%">우선순위</td>
    <td width="10%">사용여부</td>	
</tr>
<%
	for i=0 to oMultiEventList.FResultCount-1
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oMultiEventList.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
	<% if mode="copy" then %>
	<td><input type="checkbox" name="chkI" value="<%=oMultiEventList.FItemList(i).Fidx%>"></td>
	<% end if %>
    <td style="cursor:pointer;" onclick="jsmodify('<%=oMultiEventList.FItemList(i).Fidx%>','<%=oMultiEventList.FItemList(i).FcontentType%>');return false;"><%=oMultiEventList.FItemList(i).Fidx%><p>&nbsp;</p>
		<!--<a href="" onclick="window.open('enjoy_preview.asp?idx=<%=oMultiEventList.FItemList(i).Fidx%>','enjoypreview', 'width=733, height=900');return false;">[미리보기]</a>-->
	</td>
	<td>
		<%	
			select case oMultiEventList.FItemList(i).FdispOption
				case 1
					response.write "기본"
				case 2
					response.write "메인상단기획전"
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
			Response.Write "시작: "
			Response.Write replace(left(oMultiEventList.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(oMultiEventList.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(oMultiEventList.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />종료: "
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
    <td><%=chkiif(oMultiEventList.FItemList(i).Fisusing="N","사용안함","사용함")%></td>
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