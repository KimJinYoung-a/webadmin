<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/only_sys/only_sys_cls.asp"-->

<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<%
	Dim clist, vPage, iTotCnt, i, vItemID, vS_Gubun, vS_iSD, vS_iED, vS_PKIdx, vS_PKIdx_txt
	vPage = NullFillWith(requestCheckVar(Request("page"),10),1)
	vS_Gubun = requestCheckVar(Request("s_gubun"),20)
	vS_iSD = requestCheckVar(Request("iSD"),10)
	vS_iED = requestCheckVar(Request("iED"),10)
	vS_PKIdx = requestCheckVar(Request("s_pkidx"),20)
	vS_PKIdx_txt = Trim(requestCheckVar(Request("s_pkidx_txt"),100))
	
	Set clist = New cOnlySys
	 	clist.FCurrPage = vPage
	 	clist.FPageSize = 30
	 	clist.FGubun = vS_Gubun
	 	clist.FEvtSDate = vS_iSD
	 	clist.FEvtEDate = vS_iED
	 	clist.FPK_Idx = vS_PKIdx
	 	clist.FPK_Idx_txt = vS_PKIdx_txt
		clist.fnSCMChangeList
		iTotCnt = clist.ftotalcount
%>

<style type="text/css">
.dummy1 {}
</style>
<script>
function searchFrm(p){
	frm.page.value = p;
	frm.submit();
}
function jsViewContents(i){
	if($("#span_contents"+i+"").is(":hidden")){
		$("#span_contents"+i+"").show();
	}else{
		$("#span_contents"+i+"").hide();
	}	
}
function jsAllView(g){
	if(g == "c"){
		$(".dummy1").hide();
	}else{
		$(".dummy1").show();
	}
}
function jsSearchChange(v){
	if(v == "catecode"){
		$("#s_gubun > option[value='dispcate']").attr("selected", "true");
	}else if(v == "itemid"){
		$("#s_gubun > option[value='item']").attr("selected", "true");
	}else if(v == "itemoption"){
		$("#s_gubun > option[value='itemoption']").attr("selected", "true");
	}else if(v == "event"){
		$("#s_gubun > option[value='event']").attr("selected", "true");
	}else if(v == ""){
		$("#s_gubun > option[value='']").attr("selected", "true");
	}
}
</script>
<br>
<h2>* Admin 작업 변경 로그</h1>
<form name="frm" action="<%=CurrURL()%>" method="get" style="margin:0px;">
<input type="hidden" name="page" value="">
<table width="860" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
		<table class="a">
		<tr>
			<td>
				구분 : 
				<select name="s_gubun" id="s_gubun" class="select" onChange="frm.submit();">
					<option value="">-선택-</option>
					<option value="item" <%=CHKIIF(vS_Gubun="item","selected","")%>>item</option>
					<option value="itemoption" <%=CHKIIF(vS_Gubun="itemoption","selected","")%>>itemoption</option>
					<option value="dispcate" <%=CHKIIF(vS_Gubun="dispcate","selected","")%>>dispcate</option>
					<option value="event" <%=CHKIIF(vS_Gubun="event","selected","")%>>event</option>
				</select>
				&nbsp;&nbsp;&nbsp;
				등록일 : 
		        <input id="iSD" name="iSD" value="<%=vS_iSD%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		        <input id="iED" name="iED" value="<%=vS_iED%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "iSD", trigger    : "iSD_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "iED", trigger    : "iED_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
			</td>
			<td rowspan="3" style="padding:0 0 0 30px;" valign="top"><input type="submit" value=" 검  색 " class="button" style="width:70px;height:50px;" onfocus="this.blur();"></td>
		</tr>
		<tr>
			<td>
				키값검색 : 
				<select name="s_pkidx" class="select" onChange="jsSearchChange(this.value);">
					<option value="">-선택-</option>
					<option value="itemid" <%=CHKIIF(vS_PKIdx="itemid","selected","")%>>상품코드</option>
					<option value="itemoption" <%=CHKIIF(vS_PKIdx="itemoption","selected","")%>>상품코드(옵션검색)</option>
					<option value="catecode" <%=CHKIIF(vS_PKIdx="catecode","selected","")%>>1뎁스카테고리코드</option>
					<option value="event" <%=CHKIIF(vS_PKIdx="event","selected","")%>>이벤트코드</option>
				</select>
				<input type="text" name="s_pkidx_txt" value="<%=vS_PKIdx_txt%>" size="60">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<br>
<input type="button" value="상세내용전부보기" onClick="jsAllView('o');">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="button" value="상세내용전부닫기" onClick="jsAllView('c');">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<strong>총 Count : <%=clist.ftotalcount%> 개, 총 Page : <%=clist.FTotalPage%> page</strong>
<br>
<table border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50" nowrap>idx</td>
  	<td width="180" nowrap>등록일</td>
  	<td width="150" nowrap>작업자</td>
	<td width="110" nowrap>구분</td>
	<td width="450" nowrap>작업메뉴</td>
  	<td width="110" nowrap>접근IP</td>
</tr>
<%
If clist.FResultCount > 0 Then
	For i=0 To clist.FResultCount-1
%>
	<tr height="25" bgcolor="#FFFFFF" onClick="jsViewContents('<%=clist.FItemList(i).Fidx%>');" style="cursor:pointer;" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F0F0F0'">
		<td width="50" align="center"><%=clist.FItemList(i).Fidx%></td>
		<td width="180" align="center"><%=clist.FItemList(i).Fregdate%></td>
		<td width="150" align="center"><%=clist.FItemList(i).Fusername%>(<%=clist.FItemList(i).Fuserid%>)</td>
		<td width="110" align="center"><%=clist.FItemList(i).Fgubun%></td>
		<td width="450" style="padding-left:5px;">
			<%=clist.FItemList(i).Fmenuname%>(menupos:<%=clist.FItemList(i).Fmenupos%>)
			&nbsp;
			<% If clist.FItemList(i).Fmenupos <> "0" Then %>
			[<a href="<%=clist.FItemList(i).Fmenulink%>" target="_blank">링 크</a>]
			<% End If %>
		</td>
		<td width="110" align="center"><%=clist.FItemList(i).Frefip%></td>
	</tr>
	<tr>
		<td colspan="6" width="1150" bgcolor="#FFFFFF" id="td_contents" style="word-break:break-all;">
			<span class="dummy1" id="span_contents<%=clist.FItemList(i).Fidx%>" style="display:none;">
			<% If clist.FItemList(i).Fgubun = "item" OR clist.FItemList(i).Fgubun = "dispcate" Then %>
				<% If Len(clist.FItemList(i).Fpk_idx) < 4 Then %>
				<% Else %>
				- 상품코드 : <%=clist.FItemList(i).Fpk_idx%> [<a href="http://www.10x10.co.kr/<%=clist.FItemList(i).Fpk_idx%>" target="_blank">링 크</a>]<br />
				<% End If %>
			<% End If %>
			<%=Replace(clist.FItemList(i).Fcontents,vbCrLf,"<br />")%>
			</span>
		</td>
	</tr>
<%
	Next
End If
%>
</table>

<table width="1100" border="0" class="a">
<tr height="50" bgcolor="FFFFFF">
	<td colspan="20" align="center">
		<% if clist.HasPreScroll then %>
		<a href="javascript:searchFrm('<%= clist.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + clist.StartScrollPage to clist.FScrollCount + clist.StartScrollPage - 1 %>
			<% if i>clist.FTotalpage then Exit for %>
			<% if CStr(vPage)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if clist.HasNextScroll then %>
			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
<% Set clist = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->