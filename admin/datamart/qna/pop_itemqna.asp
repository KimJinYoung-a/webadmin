<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/datamart/qna/qna_summaryCls.asp"-->
<%
'DATAMART>>Q&A 통계에서 넘어온 값''''''''''''''''''''''''''''''''''''''''
Dim sType, sTypeVal, iSD, iED
sType		= requestCheckVar(request("sType"),10)
sTypeVal	= requestCheckVar(request("sTypeVal"),32)
iSD			= requestCheckVar(request("iSD"),10)
iED			= requestCheckVar(request("iED"),10)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim notupbea, mifinish, makerid, research, page, i
Dim cdl ,cdm,cds, sDt , edt , chkTerm , userid, secretYN
Dim dplusday
Dim oQna, dispCate, maxDepth, searchType, searchString
sDt				= Request("sDt")
eDt				= Request("eDt")
notupbea		= request("notupbea")
mifinish		= request("mifinish")
makerid			= request("makerid")
research		= request("research")
userid			= request("userid")
page			= request("page")
cdl				= Request("cdl")
cdm				= Request("cdm")
cds				= Request("cds")
chkTerm			= Request("chkTerm")
dplusday		= Request("dplusday")
secretYN		= requestCheckVar(request("secretYN"),1) '//공개여부
dispCate		= requestCheckVar(Request("disp"),10) 		'전시 카테고리
searchType		= requestCheckVar(Request("searchType"),15)
searchString	= requestCheckVar(Request("searchString"),150)

If page = "" Then page = 1
If research = "" and mifinish = "" Then mifinish = "N"
If sDt = "" Then sDt = iSD
If eDt = "" Then eDt = iED
If sType = "brand" and makerid = "" Then
	makerid = sTypeVal
ElseIf sType = "category" and dispCate = "" Then
	dispCate = sTypeVal
End If
maxDepth = 5

SET oQna = new cQnaSummary
	oQna.FPageSize = 20
	oQna.FCurrpage = page
	oQna.FRectMakerid = makerid
	oQna.FRectOnlyTenBeasong = notupbea
	oQna.FRectCateCode = dispCate
	oQna.FRectuserid = userid
	oQna.FRectDPlusDay = dplusday
	oQna.FReckMiFinish = mifinish
	oQna.frectstartdate = sDt
	oQna.frectenddate = eDt
	oQna.FRectSecretYN = secretYN '//비밀글 추가
	oQna.FRectSearchType = searchType
	oQna.FRectSearchString = searchString
	oQna.ItemQnaList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
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
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="chkTerm" value="<%=chkTerm%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" >
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		기간
        <input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		&nbsp;
		고객ID : <input type="text" class="text" name="userid" size="12" value="<%=userid%>" >
		&nbsp;
		브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
		&nbsp;
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
		<input type=checkbox name=dplusday value="3" <% if dplusday="3" then response.write "checked" %> > 작성일(D+3)
		&nbsp;
		답변여부 : 
			<select name="mifinish" class="select">
			<option value="" <%=chkiif(mifinish="","selected","")%>>전체</option>
			<option value="N" <%=chkiif(mifinish="N","selected","")%>>미답변</option>
			<option value="Y" <%=chkiif(mifinish="Y","selected","")%>>답변</option>
		</select>
		&nbsp;
		공개여부 : 
		<select name="secretYN" class="select">
			<option value="" <%=chkiif(secretYN="","selected","")%>>전체</option>
			<option value="N" <%=chkiif(secretYN="N","selected","")%>>공개글</option>
			<option value="Y" <%=chkiif(secretYN="Y","selected","")%>>비밀글</option>
		</select>
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
		<!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
		&nbsp;&nbsp;
		<select name="searchType" class="select">
			<option value="">-Choice-</option>
			<option value="itemid" 			<%= Chkiif(searchType = "itemid", "selected", "") %> >상품코드</option>
			<option value="qnaContent"		<%= Chkiif(searchType = "qnaContent", "selected", "") %> >질문내용</option>
			<option value="replyContent"	<%= Chkiif(searchType = "replyContent", "selected", "") %> >답변내용</option>
		</select>
		<input type="text" class="text" name="searchString" value="<%= searchString %>" size="50" />
	</td>
</tr>
</table>
<p />
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oQna.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oQna.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td height="25" align="center">고객명(아이디)</td>
    <td align="center">내용</td>
    <td width="60" align="center">상품ID</td>
    <td align="center">브랜드</td>
    <td width="45" align="center">배송</td>
    <td width="80" align="center">작성일</td>
    <td width="150" align="center">답변자</td>
    <td width="80" align="center">답변일</td>
</tr>
<% For i = 0 to (oQna.FResultCount - 1) %>
<tr height="25" bgcolor="#FFFFFF" >
	<td >&nbsp;<%= oQna.FItemList(i).Fusername %>(<%= oQna.FItemList(i).Fuserid %>)</td>
	<td >&nbsp;
		<a href="newitemqna_view.asp?id=<%= oQna.FItemList(i).Fid %>&menupos=<%= menupos %>&makerid=<%= makerid %>&page=<%= page %>&notupbea=<%= notupbea %>&mifinish=<%=  mifinish%>&research=<%= research %>&sType=<%= sType %>&sTypeVal=<%= sTypeVal %>&iSD=<%= iSD %>&iED=<%= iED %>"><%=chkiif(oQna.FItemList(i).FSecretYN="Y","<font color='red'>&lt;비밀글&gt;</font>","")%>
			<%= db2html(oQna.FItemList(i).Ftitle) %>
		</a>
	</td>
	<td align="center"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oQna.FItemList(i).FItemID %>" target=_blank><%= oQna.FItemList(i).FItemID %></a></td>
	<td align="center"><%= oQna.FItemList(i).Fmakerid %></td>
	<td align="center"><font color="<%= oQna.FItemList(i).GetDeliveryTypeColor %>"><%= oQna.FItemList(i).GetDeliveryTypeName %></font></td>
	<td align="center"><%= FormatDate(oQna.FItemList(i).Fregdate, "0000-00-00") %></td>
	<td align="center">
		<%
		If oQna.FItemList(i).FCSusername = "" AND oQna.FItemList(i).Freplyuser = "" Then
		ElseIf oQna.FItemList(i).FCSusername <> "" Then
			response.write oQna.FItemList(i).FCSusername
		ElseIf oQna.FItemList(i).Freplyuser <> "" Then
			response.write "[업체]"&oQna.FItemList(i).Freplyuser
		End If
		%>
	</td>
	<td align="center">
	<%
		If Not IsNULL(oQna.FItemList(i).FReplydate) then
			response.write FormatDate(oQna.FItemList(i).FReplydate, "0000-00-00")
		End If
	%>
	</td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if oQna.HasPreScroll then %>
			<a href="javascript:NextPage('<%= CStr(oQna.StartScrollPage - 1) %>')">[prev]</a>
		<% else %>
			[prev]
		<% end if %>
		<% for i = oQna.StartScrollPage to (oQna.StartScrollPage + oQna.FScrollCount - 1) %>
		  <% if (i > oQna.FTotalPage) then Exit For %>
		  <% if CStr(i) = CStr(oQna.FCurrPage) then %>
			 <font color="red">[<%= i %>]</font>
		  <% else %>
			 <a href="javascript:NextPage('<%= i %>')" class="id_link">[<%= i %>]</a>
		  <% end if %>
		<% next %>
		<% if oQna.HasNextScroll then %>
			<a href="javascript:NextPage('<%= CStr(oQna.StartScrollPage + oQna.FScrollCount) %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
<% SET oQna = nothing  %>
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->