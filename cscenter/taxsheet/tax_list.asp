<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 세금계산서 발행정보
' History : 서동석 생성
'			2022.10.31 한용민 수정(위하고 세금계산서 발행 api 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<%
dim taxIdx, page, searchDiv, searchBilldiv, searchKey, searchString, param, sdate, edate, chkTerm
dim oTax, i, bgcolor, strIsue, chkDel, consignYN
	taxIdx = requestcheckvar(getNumeric(trim(request("taxIdx"))),10)
	page = requestcheckvar(getNumeric(trim(request("page"))),10)
	searchDiv = requestcheckvar(trim(request("searchDiv")),1)
	searchBilldiv = requestcheckvar(trim(request("searchBilldiv")),2)
	searchKey = requestcheckvar(trim(request("searchKey")),32)
	searchString = requestcheckvar(Trim(request("searchString")),128)
	sdate = requestcheckvar(trim(request("sdate")),10)
	edate = requestcheckvar(trim(request("edate")),10)
	chkTerm = requestcheckvar(trim(request("chkTerm")),1)
    chkDel = requestcheckvar(trim(request("chkDel")),1)
	consignYN = requestcheckvar(trim(request("consignYN")),1)

	if ((session("ssAdminLsn") = "1") or (session("ssAdminLsn") = "2") or (session("ssAdminLsn") = "3") or (session("ssAdminPsn") = "8")) then
		'파트선임이상
	else
		'기타 - 자기가 작성한 계산서만 조회가능
		''searchKey = "t.userid"
		''searchString = session("ssBctId")
	end if

	if page="" then
		page=1
		searchDiv = "N"
		chkDel = "N"
	end if
	''if searchKey="" then searchKey="t.orderserial"
	if sdate="" then	sdate = dateadd("m",-1,date)
	if edate="" then	edate = date()

	param = "&menupos=" & menupos & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString & "&sdate=" & sdate & "&edate=" & edate & "&chkTerm=" & chkTerm & "&consignYN=" & consignYN

	'// 클래스 선언
	set oTax = new CTax
	oTax.FCurrPage = page
	oTax.FPageSize = 20
	oTax.FRectsearchDiv = searchDiv
	oTax.FRectsearchBilldiv = searchBilldiv
	oTax.FRectsearchKey = searchKey
	oTax.FRectsearchString = searchString
	oTax.FRectSdate = sdate
	oTax.FRectEdate = edate
	oTax.FRectchkTerm = chkTerm
	oTax.FRectDelYn = chkDel
	oTax.FRectConsignYN = consignYN
	oTax.GetTaxList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function chk_form()
{
	var frm = document.frm_search;

	/*
	if(!frm.searchKey.value)
	{
		alert("검색 조건을 선택해주십시오.");
		frm.searchKey.focus();
		return;
	}
	*/
	frm.submit();
}

function goPage(pg)
{
	var frm = document.frm_search;

	frm.page.value= pg;
	frm.submit();
}

function chgDiv()
{
	var frm = document.frm_search;
	frm.submit();
}

function switchPrintBox()
{
	var form=document.frm_list;

	if(form.chkSelect.length>1)
	{
		for(i=0;i<form.chkSelect.length;i++)
		{
			if(form.switchPrint.checked)
				form.chkSelect[i].checked=true;
			else
				form.chkSelect[i].checked=false;
		}
	}
	else
	{
		if(form.switchPrint.checked)
			form.chkSelect.checked=true;
		else
			form.chkSelect.checked=false;
	}
}

function wordPrint()
{
	var form=document.frm_list;
	var chk = 0;

	if(form.chkSelect.length>1)
	{
		for(i=0;i<form.chkSelect.length;i++)
		{
			if(form.chkSelect[i].checked)
				chk++;
		}
	}
	else
	{
		if(form.chkSelect.checked)
			chk++;
	}

	if(chk==0)
	{
		alert("출력을 원하시는 요청서를 선택해주십시요.");
		return false;
	}
	else
	{
		form.action="tax_print.asp";
		form.submit();
	}
}

function BatchTaxPrint()
{
	var form=document.frm_list;
	var chk = 0;

	if(form.chkSelect.length>1)
	{
		for(i=0;i<form.chkSelect.length;i++)
		{
			if(form.chkSelect[i].checked)
				chk++;
		}
	}
	else
	{
		if(form.chkSelect.checked)
			chk++;
	}

	if(chk==0)
	{
		alert("출력을 원하시는 요청서를 선택해주십시요.");
		return false;
	}
	else
	{
		form.action="taxsheet_process.asp";
		form.mode.value="BatchOk";
		form.submit();
	}
}

function register_new() {
	document.location.href = 'tax_view.asp?menupos=<%= menupos %>';
}

$(document).ready(function() {
	var linkTr = "tr.linkTr";

	$(linkTr).click(function() {
		window.location = $(this).attr("url");
	});

	$(linkTr).hover(
		function() {
        	$(this).css('cursor','pointer');
			$(this).css("background-color","#F1F1F1");
    	},
		function() {
			$(this).css("background-color","#FFFFFF");
    	}
	);
});

</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<!-- 검색 시작 -->
<form name="frm_search" method="GET" action="tax_list.asp" onSubmit="return false" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			위수탁구분:
			<select class="select" name="consignYN" onchange="chgDiv()">
				<option value="">전체</option>
				<option value="N" <%if consignYN = "N" then %>selected<% end if %>>정상</option>
				<option value="Y" <%if consignYN = "Y" then %>selected<% end if %>>위수탁</option>
			</select>
			&nbsp;
			발급여부:
			<select class="select" name="searchDiv" onchange="chgDiv()">
				<option value="">전체</option>
				<option value="Y" <%if searchDiv = "Y" then %>selected<% end if %>>발급</option>
				<option value="N" <%if searchDiv = "N" then %>selected<% end if %>>미발급</option>
			</select>
			&nbsp;
			발행구분:
			<select class="select" name="searchBilldiv" onchange="chgDiv()">
				<option value="">전체</option>
				<option value="">-------</option>
				<option value="01" <%if searchBilldiv = "01" then %>selected<% end if %>>소비자(customer)</option>
				<option value="02" <%if searchBilldiv = "02" then %>selected<% end if %>>가맹점(accounts)</option>
				<option value="03" <%if searchBilldiv = "03" then %>selected<% end if %>>프로모션(promotion)</option>
				<option value="51" <%if searchBilldiv = "51" then %>selected<% end if %>>기타매출(accounts)</option>
				<option value="">-------</option>
				<option value="99" <%if searchBilldiv = "99" then %>selected<% end if %>>3PL</option>
				<option value="">-------</option>
				<option value="52" <%if searchBilldiv = "52" then %>selected<% end if %>>유아러걸(발행금지)</option>
				<option value="53" <%if searchBilldiv = "53" then %>selected<% end if %>>아이띵소(발행금지)</option>
				<option value="54" <%if searchBilldiv = "54" then %>selected<% end if %>>텐바이텐 리빙(발행금지)</option>
				<option value="55" <%if searchBilldiv = "55" then %>selected<% end if %>>에이플러스비(발행금지)</option>
			</select>
			&nbsp;
			검색조건:
			<select class="select" name="searchKey">
				<option value="">선택</option>
				<option value="">-------</option>
				<option value="t.orderserial">주문번호</option>
				<option value="t.userid">등록자아이디</option>
				<option value="">-------</option>
				<option value="s.busiName">업체명(공급자)</option>
				<option value="s.busiNo">사업자번호(공급자)</option>
				<option value="">-------</option>
				<option value="c.busiName">업체명(공급받는자)</option>
				<option value="c.busiNo">사업자번호(공급받는자)</option>
			</select>
			<script language="javascript">
				document.frm_search.searchDiv.value="<%=searchDiv%>";
				document.frm_search.searchKey.value="<%=searchKey%>";
			</script>
			<input type="text" class="text" name="searchString" size="20" value="<%= searchString %>">
			&nbsp;
			삭제구분:
			<select class="select" name="chkDel">
			    <option value="">전체</option>
				<option value="N" <%=CHKIIF(chkDel="N","selected","") %> >정상</option>
				<option value="Y" <%=CHKIIF(chkDel="Y","selected","") %> >삭제</option>
			</select>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="chk_form()">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="checkbox" name="chkTerm" value="Y" <% if chkTerm="Y" then Response.Write "checked"%>>기간검색
			(발행일)
			<input id="sdate" name="sdate" value="<%=sdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
			<input id="edate" name="edate" value="<%=edate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="edate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script language="javascript">
				var CAL_Start = new Calendar({
					inputField : "sdate", trigger    : "sdate_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				var CAL_End = new Calendar({
					inputField : "edate", trigger    : "edate_trigger",
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
</table>
</form>

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<% if (session("ssAdminLsn") = "1") then %>
				<input type="button" class="button_s" value="신규발행" onClick="register_new()"> (관리자권한)
				<br>* 매장 계산서 발행 (exec [db_shop].[dbo].[usp_Ten_TaxReg_OFF] 'tozzinet', '1510160110253686')
			<% else %>
				* 관리자 외 계산서 수기발행 불가
			<% end if %>
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<form name="frm_list" method="Post" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			검색결과 : <b><%= oTax.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= oTax.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<% if FALSE and searchDiv="N" and oTax.FTotalCount>0 then %><td align="center" width="10"><input type="checkbox" name="switchPrint" onClick="switchPrintBox()"></td><% end if %>
		<td width="50">IDX</td>
		<td>발행구분</td>
		<td width="40">위수탁<br>구분</td>
		<td width="95">사업자번호</td>
		<td><b>공급자</b></td>
		<td width="95">사업자번호</td>
		<td><b>공급받는자</b></td>
		<td width="80">관련IDX</td>
		<td>상품명</td>
		<td width="75">발행일</td>
		<td width="30">과세<br>구분</td>
		<td width="65">공급가액</td>
		<td width="50">세액</td>
		<td width="75">합계</td>

		<!--
		<td>부서</td>
		-->
		<td>계정</td>
		<!--
		<td width="80">공급자<br>그룹코드</td>
		<td width="80">공급받는자<br>그룹코드</td>
		-->
		<td width="50">삭제<br>여부</td>
		<td width="80">발급여부</td>
		<!--
		<td width="80">등록자</td>
		<td width="65">등록일</td>
		-->
	</tr>
	<%
		for i=0 to oTax.FResultCount - 1
			'발급여부
			if oTax.FTaxList(i).FisueYn="Y" then
				strIsue = "<font color=darkblue>발급</font>"
			else
				strIsue = "<font color=darkred>미발급</font>"
			end if
	%>
	<tr class="linkTr" url="tax_view.asp?taxIdx=<%= oTax.FTaxList(i).FtaxIdx %>&page=<%=page & param%>" height="25" align="center" bgcolor="#FFFFFF">
		<td><a href="tax_view.asp?taxIdx=<%= oTax.FTaxList(i).FtaxIdx %>&page=<%=page & param%>"><%= oTax.FTaxList(i).FtaxIdx %></a></td>
		<td align="left"><%= oTax.FTaxList(i).BillDivString %></td>
		<td><%= oTax.FTaxList(i).GetConsignmentYN %></td>
		<td>
			<% if (oTax.FTaxList(i).FsupplyBusiNo <> "211-87-00620") then %><font color="blue"><% end if %>
			<b><%= oTax.FTaxList(i).FsupplyBusiNo %></b>
		</td>
		<td align="left">&nbsp;<%= oTax.FTaxList(i).FsupplyBusiName %></td>
		<td><b><%= oTax.FTaxList(i).FBusiNo %></b></td>
		<td align="left">&nbsp;<%= oTax.FTaxList(i).FBusiName %></td>
		<td>
			<% if (Trim(oTax.FTaxList(i).Forderserial) <> "") then %>
				<%=oTax.FTaxList(i).Forderserial%>
			<% else %>
				<% if (oTax.FTaxList(i).Forderidx <> 0) then %>
					<%=oTax.FTaxList(i).Forderidx %>
				<% else %>
					<%=oTax.FTaxList(i).GetMultiOrderIdxSUM %>
				<% end if %>
			<% end if %>
		</td>
		<td align="left">&nbsp;<a href="tax_view.asp?taxIdx=<%= oTax.FTaxList(i).FtaxIdx %>&page=<%=page & param%>"><%= db2html(oTax.FTaxList(i).Fitemname) %>&nbsp;</a></td>
		<td>
			<b><%= FormatDate(oTax.FTaxList(i).FisueDate,"0000-00-00") %></b>
		</td>
		<td><%= oTax.FTaxList(i).TaxTypeString %></td>
		<td align="right"><%= CurrFormat(oTax.FTaxList(i).FtotalPrice - oTax.FTaxList(i).FtotalTax) %></td>
		<td align="right"><%= CurrFormat(oTax.FTaxList(i).FtotalTax) %></td>
		<td align="right"><b><%= CurrFormat(oTax.FTaxList(i).FtotalPrice) %></b></td>

		<!--
		<td><%= oTax.FTaxList(i).FsellBizNm %></td>
		-->
		<td><%= oTax.FTaxList(i).FselltypeNm %></td>

		<!--
		<td>
			<% if oTax.FTaxList(i).FsupplyGroupidCnt>1 then %>
				중복(<%= oTax.FTaxList(i).FsupplyGroupidCnt %>)
			<% elseif oTax.FTaxList(i).FsupplyGroupidCnt=1 then %>
				<%= oTax.FTaxList(i).FsupplyGroupid %>
			<% end if %>
		</td>

		<td>
			<% if oTax.FTaxList(i).FgroupidCnt>1 then %>
				중복(<%= oTax.FTaxList(i).FgroupidCnt %>)
			<% elseif oTax.FTaxList(i).FgroupidCnt=1 then %>
				<%= oTax.FTaxList(i).Fgroupid %>
			<% end if %>
		</td>
		-->

		<td><%= CHKIIF(oTax.FTaxList(i).FDelYn="Y","<font color=red>삭제</font>","") %></td>
		<td><%= strIsue %></td>
		<!--
		<td><%= oTax.FTaxList(i).Fuserid %></td>
		<td><%= FormatDate(oTax.FTaxList(i).Fregdate,"0000-00-00") %></td>
		-->

	</tr>
	<%
		next
	%>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<%
			if oTax.HasPreScroll then
				Response.Write "<a href='javascript:goPage(" & oTax.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
			else
				Response.Write "[pre] &nbsp;"
			end if

			for i=0 + oTax.StartScrollPage to oTax.FScrollCount + oTax.StartScrollPage - 1

				if i>oTax.FTotalpage then Exit for

				if CStr(page)=CStr(i) then
					Response.Write " <font color='red'>[" & i & "]</font> "
				else
					Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
				end if

			next

			if oTax.HasNextScroll then
				Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
			else
				Response.Write "&nbsp; [next]"
			end if
			%>
		</td>
	</tr>
</table>
</form>

<%
set oTax = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
