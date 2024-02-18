<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : 테스터상품후기관리
' History	:  최초생성자 모름
'              2017.07.05 한용민 수정
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/board/lib/classes/testitemGoodUsingCls.asp" -->
<%
Dim page, SearchKey1, SearchKey2, selStatus, lp, lp2, sDt, eDt, cdl, cdm, cds, selPoint, chkTerm, srtMethod,blnPhotomode
Dim strDel
	page 			= requestCheckvar(Request("page"),10)
	SearchKey1 = requestCheckvar(Request("SearchKey1"),32)
	SearchKey2 = requestCheckvar(Request("SearchKey2"),10)
	selStatus = requestCheckvar(Request("selStatus"),1)
	chkTerm 	= requestCheckvar(Request("chkTerm"),10)
	srtMethod = requestCheckvar(Request("srtMethod"),10)
	sDt = requestCheckvar(Request("sDt"),10)
	eDt = requestCheckvar(Request("eDt"),10)
	cdl = requestCheckvar(Request("cdl"),3)
	cdm = requestCheckvar(Request("cdm"),3)
	cds = requestCheckvar(Request("cds"),3)
	selPoint = requestCheckvar(Request("selPoint"),10)
	blnPhotomode = requestCheckvar(Request("photomode"),5)

'기본값 지정
if page="" then page=1
if selStatus="" then selStatus="Y"
if srtMethod="" then srtMethod="idxDcd"
if sDt="" and chkTerm="" then sDt = DateAdd("m",-1,date())
if eDt="" and chkTerm="" then eDt = date()

'// 상품 후기 목록
dim oGoodUsing
Set oGoodUsing = new CGoodUsing
	oGoodUsing.FPagesize = 15
	oGoodUsing.FCurrPage = page
	oGoodUsing.FRectSearchKey1 = SearchKey1
	oGoodUsing.FRectSearchKey2 = SearchKey2
	oGoodUsing.FRectselStatus = selStatus
	oGoodUsing.FRectStartDt = sDt
	oGoodUsing.FRectEndDt = eDt
	oGoodUsing.FRectCDL = cdl
	oGoodUsing.FRectCDM = cdm
	oGoodUsing.FRectCDS = cds
	oGoodUsing.FRectPoint = selPoint
	oGoodUsing.FRectPhotoMode = blnPhotomode
	oGoodUsing.FRectSort = srtMethod
	oGoodUsing.GetGoodUsingList
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}

	// 상태 보기 변경
	function chgStatus(v)
	{
		document.frm.selStatus.value=v;
		document.frm.submit();
	}

	// 상품상세 팝업
	function viewItemInfo(iid)
	{
		var PpUp = window.open("<%=wwwurl%>/common/PopZoomItem.asp?itemid="+ iid +"&pop=pop","itemInfo","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0,width=720,height=444");
		PpUp.focus();
	}

	// 정렬방법 변경
	function ChangeSort(smtd)	{
		document.frm.srtMethod.value=smtd;
		document.frm.submit();
	}

	// 전체 선택,취소
	function chgSel_on_off()
	{
		var frm = document.frm_list;
		if (frm.lineSel.length)
		{
			for(var i=0;i<frm.lineSel.length;i++)
			{
				frm.lineSel[i].checked=frm.tt_sel.checked;
			}
		}
		else
		{
			frm.lineSel.checked=frm.tt_sel.checked;
		}
	}

	// 전체기간 설정
	function swChkTerm(ckt)	{
		if(ckt.checked) {
			frm.sDt.value="";
			frm.eDt.value="";
		}
	}

	// 선택된 항목 삭제/복구
	function doSubmit(md)
	{
		var i, chk=0, strMd;
		var frm = document.frm_list;
		if (md=='restore')
			strMd = "복구";
		else
			strMd = "삭제";

		if (frm.lineSel.length)
		{
			for(i=0;i<frm.lineSel.length;i++)
			{
				if(frm.lineSel[i].checked)
				{
					chk++;
				}
			}
		}
		else
		{
				if(frm.lineSel.checked)
				{
					chk++;
				}
		}

		if(chk==0)
		{
			alert(strMd + "할 상품후기를 적어도 한개이상 선택해주십시오.");
			return;
		}
		else
		{
			if(confirm("선택하신 " + chk + "개의  항목을 모두 " + strMd + "하시겠습니까?"))
			{
				frm.mode.value=md;
				frm.action="dotestItemGoodUsing.asp";
				frm.submit();
			}
			else
			{
				return;
			}
		}
	}
	//이미지 보기
	function showimage(img){
		var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
	}

// 카테고리 변경시 명령
function changecontent(){
}

function evalv(a)
{
	if(document.getElementById("evalview"+a+"").style.display == "block")
	{
		document.getElementById("evalview"+a+"").style.display = "none";
	}
	else
	{
		document.getElementById("evalview"+a+"").style.display = "block";
	}
}
//-->
</script>
<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="srtMethod" value="<%=srtMethod%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		<!-- #include virtual="/common/module/categoryselectbox.asp"--><br>
		아이디 <input type="text" name="SearchKey1" size="12" value="<%=SearchKey1%>" class="text">
		/ 상품번호 <input type="text" name="SearchKey2" size="12" value="<%=SearchKey2%>" class="text">
		/ 상태 <select name="selStatus" onchange="chgStatus(this.value)" class="select">
			<option value="N">삭제</option>
			<option value="Y">일반</option>
		</select>
		/ 점수 <select name="selPoint" class="select">
			<option value="">전체</option>
			<option value="1">★</option>
			<option value="2">★★</option>
			<option value="3">★★★</option>
			<option value="4">★★★★</option>
		</select>
		<br>
		검색기간
        <input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<input type="checkbox" name="chkTerm" value="Check" <% if chkTerm="Check" then Response.Write "checked" %> onClick="swChkTerm(this)">기간전체
		<script language="javascript">
			document.frm.selStatus.value="<%=selStatus%>";
			document.frm.selPoint.value="<%=selPoint%>";
			var CAL_Start = new Calendar({
				inputField : "sDt", trigger    : "sDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "eDt", trigger    : "eDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<form name="frm_list" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="selStatus" value="<%=selStatus%>">
<input type="hidden" name="SearchKey1" value="<%=SearchKey1%>">
<input type="hidden" name="SearchKey2" value="<%=SearchKey2%>">
<input type="hidden" name="sDt" value="<%=sDt%>">
<input type="hidden" name="eDt" value="<%=eDt%>">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="chkTerm" value="<%=chkTerm%>">
<input type="hidden" name="srtMethod" value="<%=srtMethod%>">
<tr>
	<td align="left">
		<% if selStatus="N" then %>
		<input type="button" value="선택복구" onClick="doSubmit('restore')" class="button">
		<% else %>
		<input type="button" value="선택삭제" onClick="doSubmit('delete')" class="button">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->
<!-- 메인 목록 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		검색결과 : <b><%=FormatNumber(oGoodUsing.FTotalCount,0)%></b>
		&nbsp;
		페이지 : <b><%= page %>/<%=FormatNumber(oGoodUsing.FtotalPage,0)%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="tt_sel" onclick="chgSel_on_off()"></td>
	<td width="40">이벤트</td>
	<td>카테고리</td>
	<td><%
		if srtMethod="iidDcd" then
			Response.Write "<a href=javascript:ChangeSort('iidAcd') title='상품코드 올림차순'>상품명▼</a>"
		elseif srtMethod="iidAcd" then
			Response.Write "<a href=javascript:ChangeSort('iidDcd') title='상품코드 내림차순'>상품명▲</a>"
		else
			Response.Write "<a href=javascript:ChangeSort('iidDcd') title='상품코드 내림차순'>상품명</a>"
		end if
	%></td>
	<td>작성자</td>
	<td width="44"><%
		if srtMethod="pntDcd" then
			Response.Write "<a href=javascript:ChangeSort('pntAcd') title='점수 올림차순'>점수▼</a>"
		elseif srtMethod="pntAcd" then
			Response.Write "<a href=javascript:ChangeSort('pntDcd') title='점수 내림차순'>점수▲</a>"
		else
			Response.Write "<a href=javascript:ChangeSort('pntDcd') title='점수 내림차순'>점수</a>"
		end if
	%></td>
	<td width="60">후기내용</td>
	<td width="80"><%
		if srtMethod="idxDcd" then
			Response.Write "<a href=javascript:ChangeSort('idxAcd') title='작성일 올림차순'>작성일▼</a>"
		elseif srtMethod="idxAcd" then
			Response.Write "<a href=javascript:ChangeSort('idxDcd') title='작성일 내림차순'>작성일▲</a>"
		else
			Response.Write "<a href=javascript:ChangeSort('idxDcd') title='작성일 내림차순'>작성일</a>"
		end if
	%></td>
	<td width="40">상태</td>
</tr>
<%
	if oGoodUsing.FResultCount=0 then
%>
<tr>
	<td colspan="9" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 상품후기가 없습니다.</td>
</tr>
<%
	else
		for lp=0 to oGoodUsing.FResultCount - 1
			if oGoodUsing.FitemList(lp).FisUsing="Y" then
				strDel = "<font color=darkblue>일반</font>"
			else
				strDel = "<font color=darkred>삭제</font>"
			end if
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="lineSel" value="<%=oGoodUsing.FitemList(lp).FUsingId%>"></td>
	<td><a href="<%=wwwUrl%>/event/eventmain.asp?eventid=<%=oGoodUsing.FitemList(lp).FEvtCode%>" target="_blank"><%=oGoodUsing.FitemList(lp).FEvtCode%></a></td>
	<td><%=oGoodUsing.FitemList(lp).FCDL & "》" & oGoodUsing.FitemList(lp).FCDM & "》" & oGoodUsing.FitemList(lp).FCDS%></td>
	<td>
		<a href="javascript:viewItemInfo(<%=oGoodUsing.FitemList(lp).Fitemid%>)" title="상품정보 보기">
		<%="[" & oGoodUsing.FitemList(lp).Fitemid & "] " & db2html(oGoodUsing.FitemList(lp).Fitemname)%></a>
	</td>
	<td>
		<%= printUserId(oGoodUsing.FitemList(lp).Fuserid, 2, "*") %>
	</td>
	<td align="left"><% for lp2=1 to oGoodUsing.FitemList(lp).FtotalPoint: Response.Write "<img src=http://fiximage.10x10.co.kr/web2008/category/icon_star.gif>":next%></td>
	<td style="cursor:pointer" onClick="evalv('<%=lp%>');">[보기]</td>
	<td><%=left(oGoodUsing.FitemList(lp).Fregdate,10)%></td>
	<td><%=strDel%></td>
</tr>
<tr id="evalview<%=lp%>" style="display:none" bgcolor="#FFFFFF">
	<td colspan="9">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td width="140" valign="top">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td style="padding-left:10px;">
						<table width="115" border="0" cellpadding="0" cellspacing="0" bgcolor="#fafafa">
						<tr>
							<td height="24" style="border-bottom:1px solid #dddddd;"><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/sub13_star.gif" width="19" height="11" hspace="7"></td>
							<td style="border-bottom:1px solid #dddddd;"><img src="http://fiximage.10x10.co.kr/web2010/category/review_star0<%=oGoodUsing.FItemList(lp).FTotalPoint%>.gif" width="56" height="14"></td>
						</tr>
						<tr>
							<td width="50" height="24" style="border-bottom:1px solid #dddddd;"><img src="http://fiximage.10x10.co.kr/web2009/category/review_text01.gif" width="19" height="11" hspace="7"></td>
							<td style="border-bottom:1px solid #dddddd;"><img src="http://fiximage.10x10.co.kr/web2010/category/review_star0<%=oGoodUsing.FItemList(lp).FPoint_fun%>grey.gif"></td>
						</tr>
						<tr>
							<td height="24" style="border-bottom:1px solid #dddddd;"><img src="http://fiximage.10x10.co.kr/web2009/category/review_text02.gif" width="28" height="11" hspace="7"></td>
							<td style="border-bottom:1px solid #dddddd;"><img src="http://fiximage.10x10.co.kr/web2010/category/review_star0<%=oGoodUsing.FItemList(lp).FPoint_dgn%>grey.gif"></td>
						</tr>
						<tr>
							<td height="24" style="border-bottom:1px solid #dddddd;"><img src="http://fiximage.10x10.co.kr/web2009/category/review_text03.gif" width="19" height="11" hspace="7"></td>
							<td style="border-bottom:1px solid #dddddd;"><img src="http://fiximage.10x10.co.kr/web2010/category/review_star0<%=oGoodUsing.FItemList(lp).FPoint_prc%>grey.gif"></td>
						</tr>
						<tr>
							<td height="24"><img src="http://fiximage.10x10.co.kr/web2009/category/review_text04.gif" width="29" height="11" hspace="7"></td>
							<td><img src="http://fiximage.10x10.co.kr/web2010/category/review_star0<%=oGoodUsing.FItemList(lp).FPoint_stf%>grey.gif"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
			<td valign="top" style="padding-top:6px;">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="bp5px"><img src="http://fiximage.10x10.co.kr/web2010/mytenbyten/tester_review_01.gif"></td>
				</tr>
				<tr>
					<td class="bp10px" style="padding-bottom:10;"><% = Replace(db2html(oGoodUsing.FItemList(lp).FContents),vbCrLf,"<br>") %></td>
				</tr>
				<tr>
					<td class="bp5px"><img src="http://fiximage.10x10.co.kr/web2010/mytenbyten/tester_review_02.gif"></td>
				</tr>
				<tr>
					<td class="bp10px" style="padding-bottom:10;"><% = Replace(db2html(oGoodUsing.FItemList(lp).FUseGood),vbCrLf,"<br>") %></td>
				</tr>
				<tr>
					<td class="bp5px"><img src="http://fiximage.10x10.co.kr/web2010/mytenbyten/tester_review_04.gif"></td>
				</tr>
				<tr>
					<td class="bp10px" style="padding-bottom:10;"><% = Replace(db2html(oGoodUsing.FItemList(lp).FUseETC),vbCrLf,"<br>") %></td>
				</tr>
				<tr>
					<td style="padding-top:10px">
					<% if oGoodUsing.FItemList(lp).FImageIcon1<>"" then %>
					<img src="<%= oGoodUsing.FItemList(lp).FImageIcon1 %>" id="file1<% = lp %>"><br>
					<% end if %>
					<% if oGoodUsing.FItemList(lp).FImageIcon2<>"" then %>
					<img src="<% = oGoodUsing.FItemList(lp).FImageIcon2 %>" id="file2<% = lp %>">
					<% end if %>
					<% if oGoodUsing.FItemList(lp).FImageIcon3<>"" then %>
					<img src="<% = oGoodUsing.FItemList(lp).FImageIcon3 %>" id="file3<% = lp %>">
					<% end if %>
					</td>
				</tr>
				</table>
			</td>
			<td width="110" align="center" valign="top" style="padding-top:25px;">
				<table border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td align="center" class="gray11px02">
					작성기간<br>
					<%=oGoodUsing.FitemList(lp).FUsewriteSdate%> ~ <%=oGoodUsing.FitemList(lp).FUsewriteEdate%>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<!-- 페이지 시작 -->
<tr>
	<td colspan="9" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<!-- 페이지 시작 -->
	<%
		if oGoodUsing.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oGoodUsing.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oGoodUsing.StartScrollPage to oGoodUsing.FScrollCount + oGoodUsing.StartScrollPage - 1

			if lp>oGoodUsing.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>[" & lp & "]</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
			end if

		next

		if oGoodUsing.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- 페이지 끝 -->
	</td>
</tr>
</form>
</table>
<!-- 페이지 끝 -->

<%
set oGoodUsing = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->