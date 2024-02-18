<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2007.11.09 한용민 수정
'           2008.01.31 허진원 수정 (저장페이지에 userid전달)
'           2012.02.15 허진원 : 미니달력 교체
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/othermall/othermall_main_contents_managecls.asp" -->

<%
dim idx, poscode,reload
	idx = request("idx")
	poscode = request("poscode")
	reload = request("reload")
	if idx="" then idx=0

if reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
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
    
    var vstartdate = new Date(frm.startdate.value.substr(0,4), frm.startdate.value.substr(5,2), frm.startdate.value.substr(8,2));
    var venddate = new Date(frm.enddate.value.substr(0,4), frm.enddate.value.substr(5,2), frm.enddate.value.substr(8,2));
    
    if (vstartdate>venddate){
        alert('종료일이 시작일보다 클 수 없습니다.');
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

function ChangeGubun(comp){
    location.href = "?poscode=" + comp.value;
    // nothing;
}

</script>

<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>외부몰 메인페이지 입력</strong></font>
			</td>		
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>
</table>
<!--표 헤드끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/othermall_doMainContentsReg.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="ckUserId" value="<%=request.Cookies("partner")("userid")%>">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF" align="center">Idx :</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF" align="center">구분명 :</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fposname %> (<%= oMainContents.FOneItem.Fposcode %>)
        <input type="hidden" name="poscode" value="<%= oMainContents.FOneItem.Fposcode %>">
        <% else %>
        <% call DrawMainPosCodeCombo("poscode", poscode, "onChange='ChangeGubun(this);'") %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF" align="center">링크구분 :</td>
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
    <td width="150" bgcolor="#DDDDFF" align="center">적용구분(반영주기):</td>
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
  <td width="150" bgcolor="#DDDDFF" align="center">이미지 :</td>
  <td><input type="file" name="file1" value="" size="32" maxlength="32" class="file">
  <% if oMainContents.FOneItem.Fidx<>"" then %>
  <br>
  <img src="<%= oMainContents.FOneItem.GetImageUrl %>" >
  <br> <%= oMainContents.FOneItem.GetImageUrl %>
  <% end if %>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF" align="center">이미지Width :</td>
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
  <td width="150" bgcolor="#DDDDFF" align="center">이미지Height :</td>
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
    <td width="150" bgcolor="#DDDDFF" align="center">링크값 :</td>
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
    <td width="150" bgcolor="#DDDDFF" align="center">반영시작일 :</td>
    <td>
		<input id="startdate" name="startdate" value="<%= oMainContents.FOneItem.Fstartdate %>" class="text" size="10" maxlength="10" />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF" align="center">반영종료일 :</td>
    <td>
		<input id="enddate" name="enddate" value="<%= Left(oMainContents.FOneItem.Fenddate,10) %>" class="text" size="10" maxlength="10" />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
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
        <input type="text" name="dummy1" value="23:59:59" size="8" readonly style="background:'#EEEEEE'">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF" align="center">등록일 :</td>
    <td>
        <%= oMainContents.FOneItem.Fregdate %> (<%= oMainContents.FOneItem.Freguserid %>)
    </td>
</tr>

<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF" align="center">사용여부 :</td>
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
</form>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="left">
			<input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
<div id=minical OnClick="this.style.display='none';" oncontextmenu='return false' ondragstart='return false' onselectstart='return false' style="background : buttonface; margin: 5; margin-top: 2;border-top: 1 solid buttonhighlight;border-left: 1 solid buttonhighlight;border-right: 1 solid buttonshadow;border-bottom: 1 solid buttonshadow;width:155;display:none;position: absolute; z-index: 99"></div>

<%
set oposcode = Nothing
set oMainContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->