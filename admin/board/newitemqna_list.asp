<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  서동석 생성
'			   2009.02.13 한용민 수정
'              2012.02.13 허진원 - 미니달력 교체
'			   2017-04-25 이종화 - 비밀글여부 추가
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/item_qnacls.asp" -->
<%
dim notupbea, mifinish, makerid, research, page, contents
dim cdl ,cdm,cds, dispCate, sDt , edt , chkTerm , userid, itemid , secretYN
dim dplusday, issoldout
	sDt = trim(Request("sDt"))
	eDt = trim(Request("eDt"))
	notupbea = trim(request("notupbea"))
	mifinish = trim(request("mifinish"))
	makerid = trim(request("makerid"))
	research = trim(request("research"))
	userid = trim(request("userid"))
	page = getNumeric(trim(request("page")))
	cdl = trim(Request("cdl"))
	cdm = trim(Request("cdm"))
	cds = trim(Request("cds"))
	chkTerm = trim(Request("chkTerm"))
	dplusday = trim(Request("dplusday"))
	itemid = requestCheckVar(trim(getNumeric(Request("itemid"))),10)
	secretYN = requestCheckVar(trim(request("secretYN")),1) '//공개여부
	issoldout = requestCheckVar(trim(request("issoldout")),2)
	contents = requestCheckVar(trim(request("contents")),800)
	dispCate = requestCheckvar(trim(request("disp")),16)

	if page="" then page=1
	if research="" and mifinish="" then mifinish="on"
	if sDt="" and chkTerm="" then sDt = DateAdd("m",-1,date())
	if eDt="" and chkTerm="" then eDt = date()

dim itemqna
set itemqna = new CItemQna
	itemqna.FPageSize = 20
	itemqna.FCurrpage = page
	itemqna.FRectMakerid = makerid
	itemqna.FRectOnlyTenBeasong = notupbea
	itemqna.FRectCDL = cdl
	itemqna.FRectcdm = cdm
	itemqna.FRectcds = cds
	itemqna.FRectCateCode = dispCate

	itemqna.FRectuserid = userid
	itemqna.FRectDPlusDay = dplusday

	itemqna.FRectItemID = itemid

	itemqna.FReckMiFinish = mifinish
	itemqna.frectstartdate = sDt
	itemqna.frectenddate = eDt

	itemqna.FRectsecretYN = secretYN '//비밀글 추가
	itemqna.FRectissoldout = issoldout
	itemqna.FRectcontents = contents
	itemqna.ItemQnaList

dim i
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript' src="/js/jsCal/js/jscal2.js"></script>
<script type='text/javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>

	function NextPage(page){
		frm.page.value=page;
		frm.submit();
	}

	// 전체기간 설정
	function swChkTerm(ckt)	{
		if(ckt.checked) {
			frm.sDt.value="";
			frm.eDt.value="";
		}
	}

	// 카테고리 변경시 명령
	function changecontent(){
	}
	//-->
</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="chkTerm" value="<%=chkTerm%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			검색기간
	        <input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
	        <input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			&nbsp;
			고객ID : <input type="text" class="text" name="userid" size="12" value="<%=userid%>" >
			&nbsp;
			브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
			&nbsp;
			상품코드 : <input type="text" class="text" name="itemid" size="12" value="<%=itemid%>" >
		</td>
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="검색" onclick="frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			배송구분 :
			<input type="radio" name="notupbea" value="" <%if (notupbea = "") then %>checked<% end if %> > 전체
			<input type="radio" name="notupbea" value="Y" <%if (notupbea = "Y") then %>checked<% end if %> > 텐배송
			<input type="radio" name="notupbea" value="N" <%if (notupbea = "N") then %>checked<% end if %> > 업체배송
			&nbsp;
			<input type=checkbox name=dplusday value="3" <% if dplusday="3" then response.write "checked" %> > D+3 초과 문의만
			<input type=checkbox name=mifinish <% if mifinish="on" then response.write "checked" %> > 미처리만검색
			<script language="javascript">
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
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			카테고리 :
			<!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
			&nbsp;
			공개여부 :
			<select name="secretYN">
				<option value="" <%=chkiif(secretYN="","selected","")%>>전체</option>
				<option value="N" <%=chkiif(secretYN="N","selected","")%>>공개글</option>
				<option value="Y" <%=chkiif(secretYN="Y","selected","")%>>비밀글</option>
			</select>
			&nbsp;
			<input type="checkbox" name="issoldout" <% if issoldout="on" then response.write " checked" %> >품절상품 제외
			&nbsp;
			내용검색 : <input type="text" class="text" name="contents" size="20" value="<%= contents %>" >
		</td>
	</tr>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">

		</td>
		<td align="right">

		</td>
	</tr>
</table>
</form>
<!-- 액션 끝 -->
<br>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if itemqna.FresultCount>0 then %>
	<tr height="30" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= itemqna.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= itemqna.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td height="25" align="center">고객명(아이디)</td>
    <td align="center">내용</td>
    <td width="60" align="center">상품ID</td>
    <td align="center">브랜드</td>
    <td width="45" align="center">배송</td>
    <td width="80" align="center">작성일</td>
    <td width="80" align="center">답변자</td>
    <td width="80" align="center">답변일</td>
    </tr>

	<% for i = 0 to (itemqna.FResultCount - 1) %>
		<tr height="25" bgcolor="#FFFFFF" >
			<td>
				<%= itemqna.FItemList(i).Fusername %><%'= printUserId(itemqna.FItemList(i).Fusername, 1, "*") %> (<%= printUserId(itemqna.FItemList(i).Fuserid, 2, "*") %>)
			</td>
			<td >
				&nbsp;
				<a href="newitemqna_view.asp?id=<%= itemqna.FItemList(i).Fid %>&menupos=<%= menupos %>&makerid=<%= makerid %>&page=<%= page %>&notupbea=<%= notupbea %>&mifinish=<%=  mifinish%>&research=<%= research %>">
				<%=chkiif(itemqna.FItemList(i).FSecretYN="Y","<font color='red'>&lt;비밀글&gt;</font>","")%>
				<%= db2html(itemqna.FItemList(i).Ftitle) %></a>
			</td>
			<td align="center"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= itemqna.FItemList(i).FItemID %>" target=_blank><%= itemqna.FItemList(i).FItemID %></a></td>
			<td align="center"><%= itemqna.FItemList(i).Fmakerid %></td>
			<td align="center"><font color="<%= itemqna.FItemList(i).GetDeliveryTypeColor %>"><%= itemqna.FItemList(i).GetDeliveryTypeName %></font></td>
			<td align="center">
                <acronym title="<%= itemqna.FItemList(i).Fregdate %>"><%= FormatDate(itemqna.FItemList(i).Fregdate, "0000-00-00") %></acronym>
            </td>
			<td align="center"><%= itemqna.FItemList(i).Freplyuser %></td>
			<td align="center">
				<% if Not IsNULL(itemqna.FItemList(i).FReplydate) then %>
                	<acronym title="<%= itemqna.FItemList(i).FReplydate %>"><%= FormatDate(itemqna.FItemList(i).FReplydate, "0000-00-00") %></acronym>
				<% end if %>
			</td>
		</tr>
	<% next %>

	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>

    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if itemqna.HasPreScroll then %>
				<a href="javascript:NextPage('<%= CStr(itemqna.StartScrollPage - 1) %>')">[prev]</a>
			<% else %>
				[prev]
			<% end if %>
			<% for i = itemqna.StartScrollPage to (itemqna.StartScrollPage + itemqna.FScrollCount - 1) %>
			  <% if (i > itemqna.FTotalPage) then Exit For %>
			  <% if CStr(i) = CStr(itemqna.FCurrPage) then %>
				 <font color="red">[<%= i %>]</font>
			  <% else %>
				 <a href="javascript:NextPage('<%= i %>')" class="id_link">[<%= i %>]</a>
			  <% end if %>
			<% next %>
			<% if itemqna.HasNextScroll then %>
				<a href="javascript:NextPage('<%= CStr(itemqna.StartScrollPage + itemqna.FScrollCount) %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
set itemqna = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp" -->
