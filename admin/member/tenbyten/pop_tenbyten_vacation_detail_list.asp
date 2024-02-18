<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<%
	Dim page, masteridx
	dim i
	dim part_sn
	dim userid, vTmpIdx, vTmpUserid

	page = Request("page")
	masteridx = Request("masteridx")

	if page="" then page=1

	userid = session("ssBctId")

	'// 로그인정보(등급)에 따라 기본 부서 설정(마스터 이상:2 및 시스템팀:7 제외)
	'if Not (session("ssAdminLsn")<=2 or session("ssAdminPsn")=7) then
	'	part_sn = session("ssAdminPsn")
	'end if

	dim oVacation
	Set oVacation = new CTenByTenVacation

	oVacation.FRectMasterIdx = masteridx
	oVacation.FRectpart_sn = part_sn
	oVacation.FRectsearchKey = " t.userid "
	oVacation.FRectsearchString = session("ssBctId")
	oVacation.FRectIsDelete = "N"

	oVacation.GetMasterOne

	oVacation.GetDetailList



%>
<!-- 검색 시작 -->
<script language="javascript">
<!--
	function AddItem()
	{
<% if (oVacation.FItemOne.Fuserid = userid) then %>
	 window.open("pop_vacation_detail_modify.asp?masteridx=<%= masteridx %>","popAddIem","width=500,height=600,scrollbars=yes"); 
<% else %>
		alert("휴가는 본인만 신청할 수 있습니다.");
<% end if %>
	}

	function ViewList(masteridx)
	{
		location.href = "/admin/member/tenbyten/pop_tenbyten_vacation_list.asp?masteridx=" + masteridx;
	}

	function SubmitAllow(masteridx, detailidx)
	{
		var frm = document.frmmodify;

		if (confirm("승인하시겠습니까?") == true) {
			frm.mode.value = "allowdetail";
			frm.masteridx.value = masteridx;
			frm.detailidx.value = detailidx;

			frm.submit();
		}
	}

	function SubmitDeny(masteridx, detailidx)
	{
		var frm = document.frmmodify;

		if (confirm("거절하시겠습니까?") == true) {
			frm.mode.value = "denydetail";
			frm.masteridx.value = masteridx;
			frm.detailidx.value = detailidx;

			frm.submit();
		}
	}

	function SubmitDelete(masteridx, detailidx)
	{
		var frm = document.frmmodify;

		if (confirm("삭제하시겠습니까?") == true) {
			frm.mode.value = "deletedetail";
			frm.masteridx.value = masteridx;
			frm.detailidx.value = detailidx;

			frm.submit();
		}
	}
   
	// 페이지 이동
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}
	
	//전자결재 품의서 등록
	function jsRegEapp(scmidx, adate, aday){ 
		var winEapp = window.open("/admin/approval/eapp/regeapp.asp","popE","width=1000,height=600,scrollbars=yes");
		document.frmEapp.iSL.value = scmidx;
		document.all.divSL.innerHTML = scmidx; 
		document.all.divDate.innerHTML = adate + " ("+aday+"일)"; 
		document.frmEapp.tC.value = document.all.divEapp.innerHTML.replace(/\r|\n/g,"");
		document.frmEapp.target = "popE";
		document.frmEapp.submit();
		winEapp.focus();
	}
	
	//전자결재 품의서 내용보기
	function jsViewEapp(reportidx,reportstate){ 
		var winEapp = window.open("/admin/approval/eapp/popIndex.asp?iRM=M01"+reportstate+"&iridx="+reportidx,"popE","");
		winEapp.focus();
	}
	
	//상세 내용보기
	function jsDetailView(idx){
		var winDetail = window.open("pop_vacation_detail_view.asp?detailidx="+idx,"popDetail","width=500,height=300,scrollbars=yes");
		winDetail.focus();
	}

//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">이름(아이디)</td>
		<td align="left">	
			<%= oVacation.FItemOne.Fusername %>(<%= oVacation.FItemOne.Fuserid %>)
		</td>
		<!--
		<td width="100" bgcolor="<%= adminColor("gray") %>">부서 / 직급</td>
		<td align="left">
			<%= oVacation.FItemOne.Fpart_name %> / <%= oVacation.FItemOne.Fposit_name %>
		</td>
		-->
		<td width="100" bgcolor="<%= adminColor("gray") %>">부서</td>
		<td align="left">
			<%= oVacation.FItemOne.Fpart_name %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">구분</td>
		<td align="left">
			<%= oVacation.FItemOne.GetDivCDStr %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">총일수</td>
		<td align="left">
			<%= oVacation.FItemOne.Ftotalvacationday %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">사용가능기간</td>
		<td align="left">
			<%= Left(oVacation.FItemOne.Fstartday,10) %> - <%= Left(oVacation.FItemOne.Fendday,10) %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">사용일수</td>
		<td align="left">
			<%= oVacation.FItemOne.Fusedvacationday %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">승인대기</td>
		<td align="left">
			<%= oVacation.FItemOne.Frequestedday %>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">잔여일수</td>
		<td align="left">
			<b><%= (oVacation.FItemOne.Ftotalvacationday - (oVacation.FItemOne.Fusedvacationday + oVacation.FItemOne.Frequestedday)) %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td height="25" width="100" bgcolor="<%= adminColor("gray") %>">사용가능</td>
		<td align="left">
			<b><%= oVacation.FItemOne.IsAvailableVacation %></b>
		</td>
		<td width="100" bgcolor="<%= adminColor("gray") %>">삭제여부</td>
		<td align="left">
			<%= oVacation.FItemOne.Fdeleteyn %>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="휴가신청" onClick="javascript:AddItem('');">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 상단 띠 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%=oVacation.FtotalCount%></b>
			&nbsp;
			페이지 : <b><%= page %> / <%=oVacation.FtotalPage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">idx</td>
		<td width="50">상태</td>
		<td width="150">기간</td>
		<td width="60">신청일수</td>
		<td width="100">등록자</td>
		<td width="100">처리자</td>
		<td>비고</td>
    </tr>
	<% if oVacation.FResultCount=0 then %>
	<tr height=30>
		<td colspan="15" align="center" bgcolor="#FFFFFF">등록(검색)된 내용이 없습니다.</td>
	</tr>
	<% else %>
		<% for i=0 to oVacation.FResultCount - 1 %>
	<tr height=30 align="center" bgcolor="<% if (oVacation.FitemList(i).Fdeleteyn="N") then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
		<td><a href="javascript:jsDetailView('<%=oVacation.FitemList(i).Fidx%>')"><%=oVacation.FitemList(i).Fidx%></a></td>
		<td><%= oVacation.FitemList(i).GetStateDivCDStr %></td>
		<td><%= Left(oVacation.FitemList(i).Fstartday,10) %> - <%= Left(oVacation.FitemList(i).Fendday,10) %></td>
		<td><%= oVacation.FitemList(i).Ftotalday %><% If oVacation.FitemList(i).Ftotalday = "0.5" Then Response.Write CHKIIF(oVacation.FItemList(i).Fhalfgubun="am","[오전]","[오후]") End If %></td>
		<td><%= oVacation.FitemList(i).Fregisterid %><% If oVacation.FitemList(i).Ftotalday = "0.5" Then Response.Write CHKIIF(userid=oVacation.FitemList(i).Fregisterid,"<br>[<a href='javascript:jsDetailView("&oVacation.FitemList(i).Fidx&");'>반차수정</a>]","") End If %></td>
		<td><%= oVacation.FitemList(i).Fapproverid %></td>
		<td>
			<% if (oVacation.FitemList(i).Fdeleteyn="N") and (oVacation.FitemList(i).Fstatedivcd="R") then %>
			<input type=button value=" 삭 제 " class="button" onclick="SubmitDelete(<%= masteridx %>, <%=oVacation.FitemList(i).Fidx%>)">
			<% end if %>
		 
			<% if isNull(oVacation.FitemList(i).Freportidx) then %>
			<input type="button" class="button"  value="품의서 작성" onClick="jsRegEapp('<%=oVacation.FitemList(i).Fidx%>','<%= Left(oVacation.FitemList(i).Fstartday,10) %> - <%= Left(oVacation.FitemList(i).Fendday,10) %>','<%= oVacation.FitemList(i).Ftotalday %>');">
			<% else %>
			<input type="button" class="button"  value="품의서 보기" onClick="jsViewEapp('<%=oVacation.FitemList(i).Freportidx%>','<%= oVacation.FitemList(i).Freportstate %>');">
			<% end if%>
		 
		</td>
	</tr>
		<% next %>

	<% end if %>
<!-- 메인 목록 끝 -->

<!-- 페이지 시작 -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<%
				if oVacation.HasPreScroll then
					Response.Write "<a href='javascript:goPage(" & oVacation.StartScrollPage-1 & ")'>[pre]</a>"
				else
					Response.Write "[pre]"
				end if

				for i=0 + oVacation.StartScrollPage to oVacation.FScrollCount + oVacation.StartScrollPage - 1

					if i>oVacation.FTotalpage then Exit for

					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
					end if

				next

				if oVacation.HasNextScroll then
					Response.Write "<a href='javascript:goPage(" & i & ")'>[next]</a>"
				else
					Response.Write "[next]"
				end if
			%>
		</td>
	</tr>
</table>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="리스트" onClick="ViewList(<%= masteridx %>);">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- 액션 끝 -->



<!--전자결재-->
	<form name="frmEapp" method="post" action="/admin/approval/eapp/regeapp.asp">
	<input type="hidden" name="tC" value=""> 
	<input type="hidden" name="ieidx" value="1"> <!-- 문서번호 지정!! -->
	<input type="hidden" name="iSL" value="">
	</form> 
	<div id="divEapp" style="display:none;"> 
	<table width="500" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">idx</td>
		<td bgcolor="#FFFFFF" width="300">
			<div id="divSL"></div>
		</td>
	</tr>
	<tr height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">SCM 아이디</td>
		<td bgcolor="#FFFFFF">
			<%= oVacation.FItemOne.Fuserid %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">구분</td>
		<td bgcolor="#FFFFFF">
			<%= oVacation.FItemOne.GetDivCDStr %>
		</td>
	</tr>
	<tr height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">사용일수/승인대기/총일수 </td>
    	<td bgcolor="#FFFFFF">
    		<%=oVacation.FItemOne.Fusedvacationday %> / <%=oVacation.FItemOne.Frequestedday%> / <%=  oVacation.FItemOne.Ftotalvacationday%>
    	</td>
    </tr>
	<tr height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">신청기간</td>
    	<td bgcolor="#FFFFFF">
    		<div id="divDate"></div>
    	</td>
    </tr> 
	</table> 
	</div>
	<!--/전자결재-->

<form name=frmmodify method=post action="domodifyvacation.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="masteridx" value="">
	<input type="hidden" name="detailidx" value="">
	<input type="hidden" name="userid" value="<%=userid%>">
</form>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->