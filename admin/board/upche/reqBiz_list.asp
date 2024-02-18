<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
'###########################################################
' Description : 업체 사업제휴서
' History : 2008.09.01 한용민 수정/추가
'			2014.05.13 정윤정 수정
'###########################################################
%>

<%
dim i, j,ix
dim page,gubun, onlymifinish
dim research, searchkey,catevalue, dispCate,maxDepth
dim ipjumYN , catemid ,catelarge, sellgubun
Dim workid
Dim iid
	page 			= requestCheckvar(request("pg"),10)
	gubun 			= requestCheckvar(request("gubun"),2)
	onlymifinish 	= requestCheckvar(request("onlymifinish"),3)
	research 		= requestCheckvar(request("research"),3)
	searchkey 		= requestCheckvar(request("searchkey"),32)
	catevalue		= requestCheckvar(request("catevalue"),3)
	ipjumYN			= requestCheckvar(request("ipjumYN"),1)
	catemid 		= requestCheckvar(request("catemidbox"),3)
	catelarge 		= requestCheckvar(request("catelargebox"),3)
	dispCate		= requestCheckVar(Request("disp"),16) 
	maxDepth		= 2
	sellgubun			= requestCheckvar(request("sellgubun"),1)
	workid			= requestCheckvar(request("workid"),34)
	iid             = requestCheckVar(Request("iid"),9) 
	
 
	 gubun="02"
	if research="" and onlymifinish="" then onlymifinish="on"		
	if (page = "") then page = "1"
 	 

dim companyrequest
set companyrequest = New CCompanyRequest
	companyrequest.PageSize = 20
	companyrequest.CurrPage = CInt(page)
	companyrequest.ScrollCount = 10
	companyrequest.FReqcd=gubun
	companyrequest.FOnlyNotFinish = onlymifinish
	companyrequest.FRectSearchKey = searchkey
	companyrequest.FRectCatevalue = catevalue
	companyrequest.FipjumYN = ipjumYN
	companyrequest.FRectDispCate = dispCate
	companyrequest.FRectSellgubun = sellgubun
	companyrequest.FRectWorkid = workid
	companyrequest.FRectID=iid
	companyrequest.list

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function checkComp(comp){
 
		document.location.href='/admin/board/upche/req_list.asp?menupos=<%=menupos%>&gubun=02&disp=&catevalue=';
	 
}

//프린트
function printpage(id){
	
	var printpage;
	printpage = window.open("/admin/board/upche/req_print.asp?id=" +id, "printpage","width=1024,height=768,scrollbars=yes,resizable=yes");
	printpage.focus();

}

function delitem(id){
	
	if (confirm("삭제하시겠습니까?.") ==true)
		frmdel.mode.value="del";
		frmdel.id.value=id;
		frmdel.submit();
}
function MovePage(page){
	frm.pg.value=page;
	frm.research.value="<%=research %>";
	//frm.gubun.value="<%=gubun%>";
	frm.onlymifinish.value="<%=onlymifinish%>";
	frm.catevalue.value="<%=catevalue%>";
	frm.ipjumYNvalue="<%=ipjumYN%>";
	frm.searchkey.value="<%=searchkey%>";
	frm.action="/admin/board/upche/req_list.asp";
	frm.submit();
}

function ViewPage(id){
 
		var winView = window.open("/admin/board/upche/req_view2.asp?id="+id,"popReq","width=1024,height=768,scrollbars=yes,resizable=yes");
	 
	winView.focus();
/*
	var winView = window.open("about:blank;","popReq","width=1024,height=768,scrollbars=yes,resizable=yes");
	frm.id.value=id;
	frm.pg.value=<%=page%>;
	frm.research.value="<%=research %>";
	//frm.gubun.value="<%=gubun%>";
	frm.onlymifinish.value="<%=onlymifinish%>";
	frm.catevalue.value="<%=catevalue%>";
	frm.ipjumYNvalue="<%=ipjumYN%>";
	frm.searchkey.value="<%=searchkey%>";
	frm.target = "popReq";
	 
		frm.action="/admin/board/upche/req_view2.asp";
	 
	frm.submit();
*/
}

function DownPage(id,sFN){
	  var winFD = window.open("<%=uploadImgUrl%>/linkweb/company/downcorequest.asp?idx="+id+"&sFN="+sFN,"popFD","");
    winFD.focus();
} 

function changecontent() {
	frm.pg.value="1";
	frm.submit();
}

</script> 
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="id" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="pg" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
	 
		<input type="hidden" name="catevalue" value="">
		제휴종류 : 
		<select name="sellgubun" class="select">
			<option value="">전체</option>
			<option value="1" <%= Chkiif(sellgubun="1", "selected", "") %> >공급제휴</option>
			<option value="2" <%= Chkiif(sellgubun="2", "selected", "") %> >컨텐츠제휴</option>
			<option value="3" <%= Chkiif(sellgubun="3", "selected", "") %> >공동마케팅 및 프로모션 제휴</option>
			<option value="4" <%= Chkiif(sellgubun="4", "selected", "") %> >문화이벤트 제휴</option>
			<option value="5" <%= Chkiif(sellgubun="5", "selected", "") %> >기술 및 솔루션 관련 제휴</option>
			<option value="6" <%= Chkiif(sellgubun="6", "selected", "") %> >광고문의</option>
		</select>&nbsp;&nbsp;
		담당자 : 
		<% DrawWorkIdCombo "workid", workid %>
	 
		<select name="ipjumYN" class="a">
			<option value="">완료구분</option>
			<option value="Y" <% if ipjumYN="Y" then response.write "selected" %>>입점완료</option>
			<option value="N" <% if ipjumYN="N" then response.write "selected" %>>미완료</option>
		</select>
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="changecontent();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		 
		&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="onlymifinish" <% if onlymifinish="on" then response.write "checked" %> >처리안된목록
		&nbsp;&nbsp;&nbsp;&nbsp;
		업체명 <input type="text" name="searchkey" value="<%= searchkey %>">	
		
		&nbsp;&nbsp;&nbsp;&nbsp;
		글번호 <input type="text" name="iid" value="<%= iid %>" size=6>		
		</td>
	</tr>
	</form>
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
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if companyrequest.resultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= companyrequest.TotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= companyrequest.TotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">번호</td>
    <td align="center">신청일</td>
    <td align="center">제목</td>
    <td align="center">처리일</td>
    <td align="center">입점여부</td> 
    <td align="center">회사URL</td>
    <td align="center">답변여부</td>
    <td align="center">비고</td>
    </tr>
	<% for i = 0 to (companyrequest.ResultCount - 1) %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%=companyrequest.results(i).id%></td>
	    <td align="center" nowrap><%= FormatDate(companyrequest.results(i).regdate, "0000-00-00") %></td>
	    <td>[<%= companyrequest.code2name(companyrequest.results(i).reqcd) %>] <%= companyrequest.results(i).companyname %></td>
	    <td align="center" nowrap>
	        <% if (IsNull(companyrequest.results(i).finishdate) = true) then %>
	      <font color="red">미완료</font>
	        <% else %>
	      <%= FormatDate(companyrequest.results(i).finishdate, "0000-00-00") %>
	        <% end if %>
	    </td>
	    <td align="center">
	    	<%if companyrequest.results(i).ipjumYN="Y" then response.write "입점완료" %>
	    	<%if companyrequest.results(i).ipjumYN="N" then response.write "N" %>
	    	</td>
	  	 
	  	<td align="center">
	  		<a href="<%IF left(companyrequest.results(i).companyurl,4)<>"http" then%>http://<%END IF%><%= companyrequest.results(i).companyurl%>" target="_blank"><%= companyrequest.results(i).companyurl%></a>
	  	</td> 
	  	<td align="center">
	  		<% if companyrequest.commentcheck(companyrequest.results(i).replycomment)="Y" then %>
	  		Y
	  		<% else %>
	  		<font color="red">N</font>
	  		<% end if %>
	  	</td>
	  	<td align="center" nowrap>
		  	<input type="button" value="보기" class="button" onclick="javascript:ViewPage(<%= companyrequest.results(i).id %>);">
		   
		  	<%if companyrequest.results(i).attachfile <> "" then%><input type="button" value="첨부파일다운" class="button" onclick="javascript:DownPage(<%= companyrequest.results(i).id %>,'<%=companyrequest.results(i).attachfile%>');"><%end if%>
	  	</td>
    </tr>   
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
		<% if companyrequest.HasPreScroll then %>
			<a href="javascript:MovePage(<%= companyrequest.StartScrollPage-1 %>);">[prev]</a>
		<% else %>
			[prev]
		<% end if %>

		<% for ix=0 + companyrequest.StartScrollPage to companyrequest.ScrollCount + companyrequest.StartScrollPage - 1 %>
			<% if ix>companyrequest.Totalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
					<font color="red">[<%= ix%>]</font>
				<% else %>
					<a href="javascript:MovePage(<%=ix%>);">[<%= ix %>]</a>
				<% end if %>
		<% next %>

		<% if companyrequest.HasNextScroll then %>
			<a href="javascript:MovePage(<%=ix%>);">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>

<form name="frmdel" method="get" action="cscenter_req_board_act.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="id" value="">
<input type="hidden" name="page" value="<%=page%>">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->