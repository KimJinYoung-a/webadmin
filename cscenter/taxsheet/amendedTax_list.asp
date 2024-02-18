<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<%

'// 사용안함
1

	'// 변수 선언 //
	dim taxIdx
	dim page, searchDiv, searchBilldiv, searchKey, searchString, param
	dim sdate, edate, chkTerm
	dim oTax, i, lp, bgcolor, strIsue
    dim chkDel

	'// 파라메터 접수 //
	taxIdx = request("taxIdx")
	page = request("page")
	searchDiv = request("searchDiv")
	searchBilldiv = request("searchBilldiv")
	searchKey = request("searchKey")
	searchString = request("searchString")
	sdate = request("sdate")
	edate = request("edate")
	chkTerm = request("chkTerm")
    chkDel = request("chkDel")

	if ((session("ssAdminLsn") = "1") or (session("ssAdminLsn") = "2") or (session("ssAdminLsn") = "3") or (session("ssAdminPsn") = "8")) then
		'파트선임이상
	else
		'기타 - 자기가 작성한 계산서만 조회가능
		''searchKey = "t1.userid"
		''searchString = session("ssBctId")
	end if

	if page="" then
		page=1
		searchDiv = "Y"
		chkDel = "N"
	end if
	if searchKey="" then searchKey="t1.orderserial"
	if sdate="" then	sdate = dateadd("m",-1,date)
	if edate="" then	edate = date()

	param = "&menupos=" & menupos & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString & "&sdate=" & sdate & "&edate=" & edate & "&chkTerm=" & chkTerm

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

	'// 수정세금계산서 발행 대상 목록
	oTax.GetAmendedTaxList

dim IsNewOrderSerial, IsNewTaxSheet

%>
<script language='javascript'>
<!--
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

	function swTermFd(ckV) {
		if(ckV.checked) {
			document.all.fdTerm.style.display='';
		} else {
			document.all.fdTerm.style.display='none';
		}
	}

	function register_new() {
		document.location.href = 'tax_register_new.asp?menupos=<%= menupos %>';
	}

//-->
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm_search" method="GET" action="amendedTax_list.asp" onSubmit="return false">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<b>일반계산서 :</b> 발급여부:
			<select class="select" name="searchDiv" onchange="chgDiv()">
				<option value="">전체</option>
				<option value="Y" <%if searchDiv = "Y" then %>selected<% end if %>>발급</option>
				<option value="N" <%if searchDiv = "N" then %>selected<% end if %>>미발급</option>
			</select>
			발행구분:
			<select class="select" name="searchBilldiv" onchange="chgDiv()">
				<option value="">전체</option>
				<option value="01" <%if searchBilldiv = "01" then %>selected<% end if %>>소비자(customer)</option>
				<option value="02" <%if searchBilldiv = "02" then %>selected<% end if %>>가맹점(accounts)</option>
				<option value="03" <%if searchBilldiv = "03" then %>selected<% end if %>>프로모션(promotion)</option>
				<option value="51" <%if searchBilldiv = "51" then %>selected<% end if %>>기타매출(accounts)</option>
				<option value="52" <%if searchBilldiv = "52" then %>selected<% end if %>>유아러걸(발행금지)</option>
				<option value="53" <%if searchBilldiv = "53" then %>selected<% end if %>>아이띵소</option>
				<option value="54" <%if searchBilldiv = "54" then %>selected<% end if %>>텐바이텐 리빙</option>
				<option value="55" <%if searchBilldiv = "55" then %>selected<% end if %>>에이플러스비</option>
			</select>
			검색조건:
			<select class="select" name="searchKey">
				<option value="">선택</option>
				<option value="t.orderserial">주문번호</option>
				<option value="t.userid">아이디</option>
				<option value="b.busiName">거래처</option>
				<option value="b.busiNo">사업자번호</option>
			</select>
			<script language="javascript">
				document.frm_search.searchDiv.value="<%=searchDiv%>";
				document.frm_search.searchKey.value="<%=searchKey%>";
			</script>
			<input type="text" class="text" name="searchString" size="20" value="<%= searchString %>">

			삭제구분
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
			<input type="checkbox" name="chkTerm" value="Y" <% if chkTerm="Y" then Response.Write "checked"%> onClick="swTermFd(this)">기간검색
			<span id="fdTerm" <% if chkTerm<>"Y" then %>style="display:none;"<% end if %>>
				(작성일기준)
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
			</span>
		</td>
	</tr>
	</form>
</table>

<p>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button_s" value="신규발행" onClick="register_new()" disabled>
			<font color=red>작업중입니다.</font>
		</td>
		<td align="right">
			<!--
			<% if searchDiv="Y" then %>
			<img src="/images/btn_word.gif" width="70" height="20" border="0" align="absmiddle" onClick="wordPrint()" style="cursor:pointer">
			<% elseif searchDiv="N" then %>
			<input type="button" value="계산서발행" onClick="BatchTaxPrint()" style="cursor:pointer" class="button">
			<% end if %>
			-->
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm_list" method="Post" action="">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oTax.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= oTax.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<% if FALSE and searchDiv="N" and oTax.FTotalCount>0 then %><td align="center" width="10"><input type="checkbox" name="switchPrint" onClick="switchPrintBox()"></td><% end if %>
		<td width="70">주문번호</td>
		<td width="50">주문<br>상태</td>
		<td width="60">결제총액</td>
		<td width="60">보조결제</td>
		<td width="65">작성일</td>
		<td width="40">일반<br>계산서</td>
		<td width="60">일반<br>합계</td>
		<td width="50">삭제<br>여부</td>
		<td>관련CS</td>


		<td>거래처명</td>
		<td width="80">사업자번호</td>
		<td width="60">공급가액</td>
		<td width="50">세액</td>



	</tr>
	<%
		for lp=0 to oTax.FResultCount - 1
			'발급여부
			if oTax.FTaxList(lp).FisueYn="Y" then
				strIsue = "<font color=darkblue>발급</font>"
			else
				strIsue = "<font color=darkred>미발급</font>"
			end if

			IsNewOrderSerial = False
			if (lp = 0) then
				IsNewOrderSerial = True
			elseif (oTax.FTaxList(lp - 1).Forderserial <> oTax.FTaxList(lp).Forderserial) then
				IsNewOrderSerial = True
			end if

			IsNewTaxSheet = False
			if (lp = 0) then
				IsNewTaxSheet = True
			elseif (CStr(oTax.FTaxList(lp - 1).FtaxIdx) <> CStr(oTax.FTaxList(lp).FtaxIdx)) then
				IsNewTaxSheet = True
			end if

	%>
	<tr align="center" bgcolor="#FFFFFF">
		<% if FALSE and searchDiv="N" then %><td><input type="checkbox" name="chkSelect" value="<%= oTax.FTaxList(lp).FtaxIdx %>"></td><% end if %>
		<td>
			<% if (IsNewOrderSerial = True) then %>
				<%= oTax.FTaxList(lp).Forderserial %>
			<% end if %>
		</td>
		<td>
			<% if (IsNewOrderSerial = True) then %>
				<% if oTax.FTaxList(lp).Fcancelyn="Y" then %>
					<font color=red>취소</font>
				<% else %>
					<font color=black>정상</font>
				<% end if %>
			<% end if %>
		</td>
		<td align="right">
			<% if (IsNewOrderSerial = True) then %>
				<%= CurrFormat(oTax.FTaxList(lp).Fsubtotalprice) %>
			<% end if %>
		</td>
		<td align="right">
			<% if (IsNewOrderSerial = True) then %>
				<%= CurrFormat(oTax.FTaxList(lp).FsumPaymentEtc) %>
			<% end if %>
		</td>
		<td>
			<% if (IsNewOrderSerial = True) or (IsNewTaxSheet = True) then %>
				<% if (oTax.FTaxList(lp).FisueYn="Y") then %>
				<%= FormatDate(oTax.FTaxList(lp).FisueDate,"0000-00-00") %>
				<% elseif (Not IsNull(oTax.FTaxList(lp).FisueDate)) then %>
				<font color="<%= adminColor("dgray") %>"><%= FormatDate(oTax.FTaxList(lp).FisueDate,"0000-00-00") %></font>
				<% end if %>
			<% end if %>
		</td>
		<td>
			<% if (IsNewOrderSerial = True) or (IsNewTaxSheet = True) then %>
				<%= strIsue %>
			<% end if %>
		</td>
		<td align="right">
			<% if (IsNewOrderSerial = True) or (IsNewTaxSheet = True) then %>
				<% if (oTax.FTaxList(lp).FDelYn="Y") then %>
					<font color=gray><%= CurrFormat(oTax.FTaxList(lp).FtotalPrice) %></font>
				<% else %>
					<% if (((oTax.FTaxList(lp).Fsubtotalprice + oTax.FTaxList(lp).FsumPaymentEtc) <> oTax.FTaxList(lp).FtotalPrice) or (oTax.FTaxList(lp).Fcancelyn="Y")) then %>
						<font color=red><%= CurrFormat(oTax.FTaxList(lp).FtotalPrice) %></font>
					<% else %>
						<%= CurrFormat(oTax.FTaxList(lp).FtotalPrice) %>
					<% end if %>
				<% end if %>
			<% end if %>
		</td>
		<td><%= CHKIIF(oTax.FTaxList(lp).FDelYn="Y","<font color=red>삭제</font>","") %></td>
		<td><%= db2html(oTax.FTaxList(lp).Fcstitle)%></td>
		<td><%= db2html(oTax.FTaxList(lp).FbusiName)%></td>
		<td><%= oTax.FTaxList(lp).FbusiNo %></td>
		<td align="right"><%= CurrFormat(oTax.FTaxList(lp).FtotalPrice - oTax.FTaxList(lp).FtotalTax) %></td>
		<td align="right"><%= CurrFormat(oTax.FTaxList(lp).FtotalTax) %></td>




	</tr>
	<%
		next
	%>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
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
	</form>
</table>
<%
set oTax = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
