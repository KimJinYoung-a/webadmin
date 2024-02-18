<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : main_manager.asp
' Discription : 사이트 메인 관리
' History : 2008.04.11 허진원 : 실서버에서 이전
'			2009.04.19 한용민 2009에 맞게 수정
'			2012.02.14 허진원 : 미니달력 교체
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/offshop/sitemaster/offshopmain_ContentsManageCls.asp" -->

<%
dim idx, poscode, reload
	idx = requestCheckVar(request("idx"),10)
	poscode = requestCheckVar(request("poscode"),10)
	reload = requestCheckVar(request("reload"),2)

if idx="" then idx=0

if reload="on" then
    response.write "<script type='text/javascript'>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End    
end if

dim oMainContents
	set oMainContents = new CMainContents
	oMainContents.FRectIdx = idx
	oMainContents.GetOneMainContents

dim oposcode, defaultMapStr
	set oposcode = new CMainContentsCode
	oposcode.FRectPosCode = poscode
	if poscode<>"" then
	    oposcode.GetOneContentsCode
	    
	    defaultMapStr = "<map name='Map_" +oposcode.FOneItem.FPosvarName + "'>" + VbCrlf
	    defaultMapStr = defaultMapStr + VbCrlf
	    defaultMapStr = defaultMapStr + "</map>"
	end if

dim orderidx
	if oMainContents.FOneItem.forderidx = "" then
		orderidx = 99
	else
		orderidx = oMainContents.FOneItem.forderidx
	end if
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>

	function SaveMainContents(frm){
	    if (frm.poscode.value.length<1){
	        alert('구분을 먼저 선택 하세요.');
	        frm.poscode.focus();
	        return;
	    }
	    
	    if (frm.linkurl.value.length<1){
	        alert('링크 값을 입력 하세요.');
	        frm.linkurl.focus();
	        return;
	    }
	    
	    if (frm.startdate.value.length!=10){
	        alert('시작일을 입력  하세요.');
	        frm.startdate.focus();
	        return;
	    }
	    
	    if (frm.enddate.value.length!=10){
	        alert('종료일을 입력  하세요.');
	        frm.enddate.focus();
	        return;
	    }
	    
	    var vstartdate = new Date(frm.startdate.value.substr(0,4), (1*frm.startdate.value.substr(5,2))-1, frm.startdate.value.substr(8,2));
	    var venddate = new Date(frm.enddate.value.substr(0,4), (1*frm.enddate.value.substr(5,2))-1, frm.enddate.value.substr(8,2));
	    
	    if (vstartdate>venddate){
	        alert('종료일이 시작일보다 빠르면 안됩니다.');
	        frm.enddate.focus();
	        return;
	    }
	
	    if ((frm.fixtype.value=="D")&&(frm.startdate.value!=frm.enddate.value)){
	        alert('반영주기 일별인 경우 시작일과 종료일을 같게 입력하세요.');
	        frm.enddate.focus();
	        return;
	    }
	    
	    if (confirm('저장 하시겠습니까?')){
	        frm.submit();
	    }
	}
	
	function ChangeLinktype(comp){
	    if (comp.value=="M"){
	       document.all.link_M.style.display = "";
	       document.all.link_L.style.display = "none";
	    }else{
	       document.all.link_M.style.display = "none";
	       document.all.link_L.style.display = "";
	    }
	}
	
	//function getOnLoad(){
	//    ChangeLinktype(frmcontents.linktype.value);
	//}
	
	//window.onload = getOnLoad;
	
	function ChangeGubun(comp){
	    location.href = "?poscode=" + comp.value;
	    // nothing;
	}
	
</script>

<table width="660" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/doOffshopMainContentsReg.asp" onsubmit="return false;" enctype="multipart/form-data">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">구분명</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fposname %> (<%= oMainContents.FOneItem.Fposcode %>)
        <input type="hidden" name="poscode" value="<%= oMainContents.FOneItem.Fposcode %>">
        <% else %>
        <% call DrawPoint1010PosCodeCombo("poscode", poscode, "onChange='ChangeGubun(this);'") %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">링크구분</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.getlinktypeName %>
        <% else %>
            <% if poscode<>"" then %>
            <%= oposcode.FOneItem.getlinktypeName %>
            <% else %>
            <font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">적용구분(반영주기)</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.getfixtypeName %>
        <input type="hidden" name="fixtype" value="<%= oMainContents.FOneItem.Ffixtype %>">
        <% else %>
            <% if poscode<>"" then %>
            <%= oposcode.FOneItem.getfixtypeName %>
            <input type="hidden" name="fixtype" value="<%= oposcode.FOneItem.Ffixtype %>">
            <% else %>
            <font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
        <% end if %>
        
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">우선순위</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        	<% if oMainContents.FOneItem.Flinktype="F" then %>
            <input type="text" name="orderidx" size=5 value="<%= orderidx %>">
        	<% end if %>
        <% else %>
            <% if poscode<>"" then %>
            	<% if oposcode.FOneItem.Flinktype = "F" then %>
            	<input type="text" name="orderidx" size=5 value="<%= orderidx %>">
           		<% end if %>
            <% else %>
            	<font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
        <% end if %>
        
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">이미지</td>
  <td><input type="file" name="file1" value="" size="32" maxlength="32" class="file">
  <% if oMainContents.FOneItem.Fidx<>"" then %>
  <br>
  <img src="<%= oMainContents.FOneItem.GetImageUrl %>" >
  <br> <%= oMainContents.FOneItem.GetImageUrl %>
  <% end if %>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">이미지Width</td>
  <td>
  <% if oMainContents.FOneItem.Fidx<>"" then %>
  <input type="text" name="imagewidth" value="<%= oMainContents.FOneItem.Fimagewidth %>" size="8" maxlength="16"> 
  <% else %>
        <% if poscode<>"" then %>
        <%= oposcode.FOneItem.Fimagewidth %>
        <% else %>
        <font color="red">구분을 먼저 선택하세요</font>
        <% end if %>
  <% end if %>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">이미지Height</td>
  <td>
  <% if oMainContents.FOneItem.Fidx<>"" then %>
  <input type="text" name="imageheight" value="<%= oMainContents.FOneItem.Fimageheight %>" size="8" maxlength="16"> 
  <% else %>
        <% if poscode<>"" then %>
        <%= oposcode.FOneItem.Fimageheight %>
        <% else %>
        <font color="red">구분을 먼저 선택하세요</font>
        <% end if %>
  <% end if %>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">링크값</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
            <% if oMainContents.FOneItem.FLinkType="M" then %>
            <textarea name="linkurl" cols="60" rows="6"><%= oMainContents.FOneItem.Flinkurl %></textarea>
            <% else %>
            <input type="text" name="linkurl" value="<%= oMainContents.FOneItem.Flinkurl %>" maxlength="128" size="40">
            <% end if %>
        <% else %>
            <% if poscode<>"" then %>
                <% if oposcode.FOneItem.FLinkType="M" then %>
                    <textarea name="linkurl" cols="60" rows="6"><%= defaultMapStr %></textarea>
                    <br>(이미지맵 변수값 변경 금지)
            	<% elseif oposcode.FOneItem.FLinkType="B" then %>
            		<input type="text" class="text_ro" name="linkurl" value="/" maxlength="128" size="40" readonly>
                <% else %>
                    <input type="text" name="linkurl" value="" maxlength="128" size="40">
                    <br>(상대경로로 표시해 주세요  ex: /event/eventmain.asp?eventid=6263)
                <% end if %>
            <% else %>
            <font color="red">구분을 먼저 선택하세요</font>
            <% end if %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">반영시작일</td>
    <td>
        <input id="startdate" name="startdate" value="<%=Left(oMainContents.FOneItem.Fstartdate,10)%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
        <% if FALSE and oMainContents.FOneItem.Ffixtype="R" then %> <!-- 실시간인경우 / 걍 일단위로 돌림 (나중에 시간단위로 돌릴때 False 제거)-->
        <input type="text" name="startdatetime" size="2" maxlength="2" value="<%= Format00(2,Hour(oMainContents.FOneItem.Fstartdate)) %>">(시 00~23)
        <input type="text" name="dummy0" value="00:00" size="6" readonly style="background:'#EEEEEE'">
        <% else %>
        <input type="text" name="dummy0" value="00:00:00" size="8" readonly style="background:'#EEEEEE'">
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">반영종료일</td>
    <td>
        <input id="enddate" name="enddate" value="<%=Left(oMainContents.FOneItem.Fenddate,10)%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
        <% if FALSE and oMainContents.FOneItem.Ffixtype="R" then %> <!-- 실시간인경우 -->
        <input type="text" name="enddatetime" size="2" maxlength="2" value="<%= ChkIIF(oMainContents.FOneItem.Fenddate="","23",Format00(2,Hour(oMainContents.FOneItem.Fenddate))) %>">(시 00~23)
        <input type="text" name="dummy1" value="59:59" size="6" readonly style="background:'#EEEEEE'">
        <% else %>
        <input type="text" name="dummy1" value="23:59:59" size="8" readonly style="background:'#EEEEEE'">
        <% end if %>
		<script type="text/javascript">
			var CAL_Start = new Calendar({
				inputField : "startdate", trigger    : "startdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "enddate", trigger    : "enddate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">등록일</td>
    <td>
        <%= oMainContents.FOneItem.Fregdate %> (<%= oMainContents.FOneItem.Freguserid %>)
    </td>
</tr>

<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">사용여부</td>
    <td>
        <% if oMainContents.FOneItem.Fisusing="N" then %>
        <input type="radio" name="isusing" value="Y">사용함
        <input type="radio" name="isusing" value="N" checked >사용안함
        <% else %>
        <input type="radio" name="isusing" value="Y" checked >사용함
        <input type="radio" name="isusing" value="N">사용안함
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);">
    <br>※ 링크구분이 링크일때에는 <b><u>사용여부</u></b>에 따라 <u><b>실시간 적용</u></b>이 됩니다.<br>따로 실시간 적용하기가 없습니다. 주의하시기 바랍니다.</td>
</tr>
</form>
</table>
<%
set oposcode = Nothing
set oMainContents = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->